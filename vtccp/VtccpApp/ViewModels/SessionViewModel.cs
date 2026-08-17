namespace VtccpApp.ViewModels;

using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using ConfigEngine;
using ConfigEngine.Models;
using DeviceInterface;
using DeviceInterface.Dmst;
using DeviceInterface.Reports;
using DeviceInterface.Rfid;
using DeviceInterface.Rfid.Gcp;
using DeviceInterface.Rfid.Models;
using ExcelEngine.Models;
using ExcelEngine.Schema;
using ExcelEngine.Session;
using OcrEngine;
using VtccpApp.Commands;
using VtccpApp.Views;

/// <summary>
/// Drives the Session Launcher page.
///
/// Scan modes:
///   Manual    — operator presses ⚡ Trigger Scan for each scan.
///   AutoPoll  — background loop fires TriggerAndGetResultAsync at <see cref="AutoPollIntervalMs"/> ms.
///   Push      — device pushes DMST XML via TCP; requires DmstListenPort &gt; 0 in the device profile.
/// </summary>
public sealed class SessionViewModel : ViewModelBase
{
    // ── Build label ───────────────────────────────────────────────────────────

    public static string BuildLabel =>
        "Build: " + (System.Reflection.Assembly
            .GetExecutingAssembly()
            .GetName().Version?.ToString(3) ?? "?");

    // ── Scan mode ─────────────────────────────────────────────────────────────

    public enum ScanMode { Manual, AutoPoll, Push }

    // ── Dependencies ──────────────────────────────────────────────────────────

    private readonly ConfigRepository  _repo;
    private readonly HistoryViewModel  _history;
    private readonly VerificationXmlMap _xmlMap = new();

    // ── Runtime state ─────────────────────────────────────────────────────────

    private DeviceSession?                              _deviceSession;
    private DeviceInterface.Dmst.HttpEventSubscriber?  _pushHttpSubscriber;
    private SessionManager?                             _sessionMgr;
    private System.Threading.CancellationTokenSource? _pollCts;
    private bool                                _isRunning;
    private string                              _statusMessage = "Ready.";
    private int                                 _recordCount;
    private LiveFeedWindow?                     _liveFeedWindow;
    private StitchingWindow?                    _stitchingWindow;

    // ── RFID ──────────────────────────────────────────────────────────────────
    private RfidScanCoordinator? _rfidCoordinator;
    private GcpValidator?        _gcpValidator;

    // ── Session output directory ───────────────────────────────────────────────
    // Captured at OnStartAsync so fire-and-forget report generators have the path
    // without relying on mutable session-manager state after the session closes.
    private string _sessionOutputDir = string.Empty;

    /// <summary>
    /// FileSystem watcher (DmstHtmlScraper) started only in Push mode when
    /// <see cref="ConfigEngine.Models.HybridReportMode.Replace"/> is active.
    /// Watches the CodeQuality folder, parses and deletes the Webscan HTML files,
    /// and makes them available (by timestamp correlation) so AcceptRecordInnerAsync
    /// can write the hybrid report back to the original file path.
    /// Not used in Manual/AutoPoll mode — DeviceSession owns the scraper there.
    /// </summary>
    private DeviceInterface.Dmst.DmstHtmlScraper? _htmlWatcher;

    // ── OCR ───────────────────────────────────────────────────────────────────

    private readonly DualEngineOcrRunner _ocrRunner = new();

    /// <summary>
    /// Controls whether OCR runs on the L1 barcode-crop image for each accepted scan.
    /// TODO: bind to a per-session UI toggle (checkbox in SessionView) before shipping.
    ///       Defaulting true so OCR results populate during development and testing.
    /// </summary>
    private bool _ocrEnabled = true;

    // ── UPC/EAN Supplemental state ────────────────────────────────────────────
    private int    _upcEanSupplementalMode = 0;
    private string _supplementalStatus     = "Not yet read from device.";
    private bool   _isApplyingSupplemental;

    /// <summary>
    /// Count of AcceptRecordAsync calls currently in flight.
    /// Incremented before AddRecord; decremented in finally.
    /// OnStopAsync drains to zero before calling CloseSession so no record
    /// is lost to a close/write race.
    /// </summary>
    private int _pendingAccept;

    // ── Selection ─────────────────────────────────────────────────────────────

    private DeviceProfile? _selectedDevice;
    private JobTemplate?   _selectedTemplate;
    private string         _operatorOverride = string.Empty;
    private ScanMode       _scanMode         = ScanMode.Push;
    private int            _autoPollIntervalMs = 500;

    // ── Bindable collections ──────────────────────────────────────────────────

    public ObservableCollection<DeviceProfile> AvailableDevices   { get; } = [];
    public ObservableCollection<JobTemplate>   AvailableTemplates { get; } = [];

    // ── Bindable properties ───────────────────────────────────────────────────

    public bool HasSelectedDevice => _selectedDevice is not null;

    public DeviceProfile? SelectedDevice
    {
        get => _selectedDevice;
        set
        {
            Set(ref _selectedDevice, value);
            OnPropertyChanged(nameof(IsPushAvailable));
            OnPropertyChanged(nameof(HasSelectedDevice));
            RelayCommand.Refresh();

            // Auto-select the best mode for the newly chosen device:
            //   • If the device has a push port → Push mode
            //   • Otherwise fall back to Manual (never leave mode on Push when unavailable)
            if (!IsRunning)
                ActiveScanMode = value?.DmstListenPort > 0 ? ScanMode.Push : ScanMode.Manual;
        }
    }

    public JobTemplate? SelectedTemplate
    {
        get => _selectedTemplate;
        set { Set(ref _selectedTemplate, value); RelayCommand.Refresh(); }
    }

    public string OperatorOverride
    {
        get => _operatorOverride;
        set => Set(ref _operatorOverride, value);
    }

    public ScanMode ActiveScanMode
    {
        get => _scanMode;
        set
        {
            Set(ref _scanMode, value);
            OnPropertyChanged(nameof(IsManualMode));
            OnPropertyChanged(nameof(IsAutoPollMode));
            OnPropertyChanged(nameof(IsPushMode));
            OnPropertyChanged(nameof(ShowTriggerButton));
            RelayCommand.Refresh();
        }
    }

    public bool IsManualMode     => _scanMode == ScanMode.Manual;
    public bool IsAutoPollMode   => _scanMode == ScanMode.AutoPoll;
    public bool IsPushMode       => _scanMode == ScanMode.Push;

    /// <summary>True in Manual and Push modes — both support a software trigger.</summary>
    public bool ShowTriggerButton => _scanMode is ScanMode.Manual or ScanMode.Push;

    /// <summary>True when the selected device has a non-zero DmstListenPort.</summary>
    public bool IsPushAvailable => _selectedDevice?.DmstListenPort > 0;

    public int AutoPollIntervalMs
    {
        get => _autoPollIntervalMs;
        set => Set(ref _autoPollIntervalMs, Math.Max(100, value));
    }

    public bool IsRunning
    {
        get => _isRunning;
        private set { Set(ref _isRunning, value); RelayCommand.Refresh(); }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => Set(ref _statusMessage, value);
    }

    public int RecordCount => _recordCount;

    // ── UPC/EAN Supplemental properties ──────────────────────────────────────

    /// <summary>
    /// Current UPC/EAN supplemental mode selection (0–4).
    /// Confirmed against DMCC Reference 6.1.16_sr4 (UPC-EAN.SUPPLEMENT):
    ///   0=Ignore, 1=Required (any), 2=Required 2-digit,
    ///   3=Required 5-digit, 4=Not Required (optional).
    /// Change the value via the IsSupplemental* radio-button helpers.
    /// </summary>
    public int UpcEanSupplementalMode
    {
        get => _upcEanSupplementalMode;
        set
        {
            if (Set(ref _upcEanSupplementalMode, value))
            {
                OnPropertyChanged(nameof(IsSupplementalIgnore));
                OnPropertyChanged(nameof(IsSupplementalRequired));
                OnPropertyChanged(nameof(IsSupplementalRequired2));
                OnPropertyChanged(nameof(IsSupplementalRequired5));
                OnPropertyChanged(nameof(IsSupplementalNotRequired));
            }
        }
    }

    public bool IsSupplementalIgnore
    {
        get => _upcEanSupplementalMode == 0;
        set { if (value) UpcEanSupplementalMode = 0; }
    }
    public bool IsSupplementalRequired
    {
        get => _upcEanSupplementalMode == 1;
        set { if (value) UpcEanSupplementalMode = 1; }
    }
    public bool IsSupplementalRequired2
    {
        get => _upcEanSupplementalMode == 2;
        set { if (value) UpcEanSupplementalMode = 2; }
    }
    public bool IsSupplementalRequired5
    {
        get => _upcEanSupplementalMode == 3;
        set { if (value) UpcEanSupplementalMode = 3; }
    }
    public bool IsSupplementalNotRequired
    {
        get => _upcEanSupplementalMode == 4;
        set { if (value) UpcEanSupplementalMode = 4; }
    }

    public string SupplementalStatus
    {
        get => _supplementalStatus;
        private set => Set(ref _supplementalStatus, value);
    }

    // ── Commands ──────────────────────────────────────────────────────────────

    public RelayCommand StartCommand            { get; }
    public RelayCommand StopCommand             { get; }
    public RelayCommand TriggerCommand          { get; }
    public RelayCommand SetManualCommand        { get; }
    public RelayCommand SetAutoPollCommand      { get; }
    public RelayCommand SetPushCommand          { get; }
    public RelayCommand ReadSupplementalCommand  { get; }
    public RelayCommand WriteSupplementalCommand { get; }
    public RelayCommand OpenLiveFeedCommand      { get; }
    public RelayCommand OpenStitchingCommand     { get; }

    public SessionViewModel(ConfigRepository repo, HistoryViewModel history)
    {
        _repo    = repo;
        _history = history;

        StartCommand   = new RelayCommand(async () => await OnStartAsync(),
            () => !IsRunning && SelectedDevice is not null && SelectedTemplate is not null);
        StopCommand    = new RelayCommand(async () => await OnStopAsync(),
            () => IsRunning);
        TriggerCommand = new RelayCommand(async () => await OnTriggerAsync(),
            () => IsRunning && (_scanMode == ScanMode.Manual || _scanMode == ScanMode.Push));

        SetManualCommand   = new RelayCommand(() => ActiveScanMode = ScanMode.Manual,   () => !IsRunning);
        SetAutoPollCommand = new RelayCommand(() => ActiveScanMode = ScanMode.AutoPoll, () => !IsRunning);
        SetPushCommand     = new RelayCommand(() => ActiveScanMode = ScanMode.Push,     () => !IsRunning && IsPushAvailable);

        ReadSupplementalCommand  = new RelayCommand(
            async () => await OnReadSupplementalAsync(),
            () => SelectedDevice is not null && !_isApplyingSupplemental);

        WriteSupplementalCommand = new RelayCommand(
            async () => await OnWriteSupplementalAsync(),
            () => SelectedDevice is not null && !_isApplyingSupplemental);

        OpenLiveFeedCommand = new RelayCommand(
            OnOpenLiveFeed,
            () => SelectedDevice is not null);

        OpenStitchingCommand = new RelayCommand(
            OnOpenStitching,
            () => SelectedDevice is not null);

        Reload();
    }

    // ── Reload from repository ────────────────────────────────────────────────

    public void Reload()
    {
        AvailableDevices.Clear();
        AvailableTemplates.Clear();
        foreach (var d in _repo.Devices)   AvailableDevices.Add(d);
        foreach (var t in _repo.Templates) AvailableTemplates.Add(t);

        if (SelectedDevice is null) SelectedDevice = _repo.DefaultDevice;

        // Keep the current template if it still exists in the repo (id-based
        // lookup survives edits to template properties).  Reset to the default
        // when the selection is gone or was never set — this fires correctly
        // after the user changes which template is marked as the default.
        var currentId = _selectedTemplate?.Id;
        SelectedTemplate = (currentId is not null
            ? AvailableTemplates.FirstOrDefault(t => t.Id == currentId)
            : null)
            ?? _repo.DefaultTemplate;

        // Pre-fill Operator ID with the value typed at the last session start.
        // The user can clear or change it before each session; the new value is
        // saved back to AppSettings when the session begins.
        if (string.IsNullOrWhiteSpace(OperatorOverride) &&
            !string.IsNullOrWhiteSpace(_repo.Settings.LastOperatorId))
        {
            OperatorOverride = _repo.Settings.LastOperatorId;
        }
    }

    // ── Start / Stop ──────────────────────────────────────────────────────────

    private async Task OnStartAsync()
    {
        if (SelectedDevice is null || SelectedTemplate is null) return;

        string outputDir = !string.IsNullOrWhiteSpace(SelectedTemplate.OutputDirectory)
            ? SelectedTemplate.OutputDirectory
            : _repo.Settings.DefaultOutputDirectory;
        _sessionOutputDir = outputDir;   // captured for hybrid report generation

        SessionState state = SelectedTemplate.ToSessionState(outputDir);
        if (!string.IsNullOrWhiteSpace(OperatorOverride))
            state.OperatorId = OperatorOverride.Trim();

        // Persist the operator ID so it pre-fills automatically next time.
        // Fire-and-forget is fine — a failed settings write is never session-critical.
        if (!string.IsNullOrWhiteSpace(state.OperatorId) &&
            state.OperatorId != _repo.Settings.LastOperatorId)
        {
            _repo.Settings.LastOperatorId = state.OperatorId;
            _ = _repo.SaveSettingsAsync();
        }

        _sessionMgr = new SessionManager(TruCheckCompatibleSchema.Build());
        _pollCts    = new System.Threading.CancellationTokenSource();

        try
        {
            if (_scanMode == ScanMode.Push)
            {
                // Push mode: subscribe directly to the device's HTTP event channel on
                // port 44444.  No SDK connection is opened so DMST can remain fully
                // active alongside CP (live view, manual Verify button, positioning).
                //
                // Results from ANY scan source — physical reader button, DMST Verify,
                // or CP's own Trigger Scan — arrive via the device's HTTP pub-sub
                // channel (PUT /codes.xml with origin="common").  This replaces the
                // old DmstListener (TCP port 9004) approach which required DMST to
                // manage the push destination and broke whenever DMST was closed or
                // CONFIG.DEFAULT was issued.
                StatusMessage = "Subscribing to device push channel…";
                var cfg = SelectedDevice.ToDeviceConfig();
                var ctx = new VerificationRecord
                {
                    Symbology       = string.Empty,
                    DeviceSerial    = string.Empty,
                    DeviceName      = SelectedDevice.Name,
                    FirmwareVersion = string.Empty,
                    OperatorId      = state.OperatorId  ?? string.Empty,
                    JobName         = state.JobName      ?? string.Empty,
                    BatchNumber     = state.BatchNumber  ?? string.Empty,
                    CompanyName     = state.CompanyName  ?? string.Empty,
                    LogoPath        = state.LogoPath,
                };
                _pushHttpSubscriber = new DeviceInterface.Dmst.HttpEventSubscriber(
                    cfg.Host, cfg.Port, _xmlMap, ctx, OnPushRecord);
                await _pushHttpSubscriber.StartAsync(_pollCts.Token);

                // ── Replace mode: watch the CodeQuality folder in Push mode ───
                // In Push mode, DeviceSession (and its built-in DmstHtmlScraper) is
                // not used.  When Replace mode is active, start a standalone scraper
                // here so we can correlate the Webscan HTML file by timestamp, record
                // its path, and write the hybrid report back to the same location.
                if (_repo.Settings.GenerateHybridReport &&
                    _repo.Settings.HybridReportMode == ConfigEngine.Models.HybridReportMode.Replace &&
                    SelectedDevice.Name is { Length: > 0 } devName)
                {
                    var watchPath = DeviceInterface.Dmst.DmstHtmlScraper.BuildReportPath(devName);
                    _htmlWatcher = new DeviceInterface.Dmst.DmstHtmlScraper(watchPath);
                    _htmlWatcher.Start();
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-REPLACE] HTML watcher started: '{watchPath}'");
                }
            }
            else
            {
                // Manual / AutoPoll: open DMCC connection (requires DMST to be closed).
                StatusMessage = "Connecting to device…";
                var cfg = SelectedDevice.ToDeviceConfig();
                _deviceSession = new DeviceSession(cfg, _xmlMap);
                await _deviceSession.ConnectAsync();

                // Subscribe to the device's HTTP result push channel — same channel
                // DMST uses for all TC verification results (codes.xml origin="common").
                // ResultReceived fires for every TC verification regardless of trigger source
                // (DMST Verify button, physical reader button, or CP's own TRIGGER).
                _deviceSession.ResultReceived += (_, rec) => OnPushRecord(rec);
                await _deviceSession.StartHttpSubscriberAsync();
            }

            await Task.Run(() => _sessionMgr.StartSession(state));
            _history.SetSessionContext(state.JobName, state.OperatorId);
            _history.ClearHistory();
            _recordCount = 0; OnPropertyChanged(nameof(RecordCount));
            IsRunning    = true;

            // ── RFID coordinator (optional) ───────────────────────────────────
            // Enabled when AppSettings.RfidComPort is set; silently skipped otherwise.
            // Connection failure is non-fatal — session continues without RFID.
            string? rfidPort = _repo.Settings.RfidComPort;
            if (!string.IsNullOrWhiteSpace(rfidPort))
            {
                try
                {
                    var rfidSettings = new RfidScanCoordinatorSettings
                    {
                        Enabled              = true,
                        ScanWindowMs         = _repo.Settings.RfidScanWindowMs,
                        FlagMismatchInReport = _repo.Settings.RfidFlagMismatch,
                    };
                    // Load the GCP prefix table so the validator can confirm each
                    // company prefix against the GS1 registry.  Failure is non-fatal:
                    // the session continues without GCP validation.
                    //
                    // Path resolution order (first existing file wins):
                    //   1. AppSettings.GcpDataPath — set by GcpUpdateService after a download
                    //   2. Update-service install target (%AppData%\VCCS\gcpPrefixes.xml)
                    //   3. Bundled seed copy deployed next to the exe (data\gcp-prefix-format-list.xml)
                    //   4. User-data fallback (DefaultOutputDirectory\data\gcp-prefix-format-list.xml)
                    _gcpValidator = null;
                    string[] gcpCandidates =
                    [
                        _repo.Settings.GcpDataPath ?? string.Empty,
                        Services.GcpUpdateServiceFactory.DefaultInstallPath,
                        Path.Combine(AppContext.BaseDirectory, "data", "gcp-prefix-format-list.xml"),
                        Path.Combine(_repo.Settings.DefaultOutputDirectory, "data", "gcp-prefix-format-list.xml"),
                    ];
                    string? gcpPath = gcpCandidates.FirstOrDefault(
                        p => !string.IsNullOrWhiteSpace(p) && File.Exists(p));

                    if (gcpPath is not null)
                    {
                        try
                        {
                            var gcpTable = GcpLengthTable.LoadFromFile(gcpPath);
                            _gcpValidator = new GcpValidator(gcpTable);
                            System.Diagnostics.Debug.WriteLine(
                                $"[GCP] Loaded {gcpTable.EntryCount} entries from '{gcpPath}'; date={gcpTable.DataDate:yyyy-MM-dd}.");

                            // Persist the resolved path and date so the settings file is always
                            // current and the PDF provenance annotation reflects the live table date.
                            bool settingsChanged = false;
                            if (string.IsNullOrWhiteSpace(_repo.Settings.GcpDataPath))
                            {
                                _repo.Settings.GcpDataPath = gcpPath;
                                settingsChanged = true;
                            }
                            string? newDateStr = gcpTable.DataDate?.ToString("O");
                            if (newDateStr != _repo.Settings.GcpLastModified)
                            {
                                _repo.Settings.GcpLastModified = newDateStr;
                                settingsChanged = true;
                            }
                            if (settingsChanged)
                                _ = _repo.SaveSettingsAsync();
                        }
                        catch (Exception gcpEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"[GCP] Table load failed: {gcpEx.Message}");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[GCP] No GCP prefix list found; GCP validation skipped.");
                    }

                    var reader    = EpcReaderFactory.CreateAsReaderP35U();
                    var validator = new RfidValidator(_gcpValidator);
                    _rfidCoordinator = new RfidScanCoordinator(reader, validator, rfidSettings);
                    await reader.ConnectAsync(rfidPort, _pollCts.Token);
                    System.Diagnostics.Debug.WriteLine($"[RFID] Coordinator started on {rfidPort}.");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[RFID] Coordinator startup failed: {ex.Message}");
                    if (_rfidCoordinator is not null)
                    {
                        await _rfidCoordinator.DisposeAsync();
                        _rfidCoordinator = null;
                    }
                }
            }

            string triggerNote = string.Empty;
            if (_deviceSession is not null && _deviceSession.OriginalTriggerType is { } origTrig)
            {
                string trigName = origTrig switch
                {
                    "0" => "Single",
                    "1" => "Presentation",
                    "2" => "Manual",
                    "3" => "Burst",
                    "4" => "Self",
                    "5" => "Continuous",
                    _   => origTrig,
                };
                triggerNote = $"  [DMST trigger was: {trigName} ({origTrig}) → restored on Stop]";
            }

            string modeLabel = _scanMode switch
            {
                ScanMode.AutoPoll => $"Auto-Poll ({_autoPollIntervalMs} ms)",
                ScanMode.Push     => $"Push (DMST) — port {SelectedDevice.DmstListenPort}",
                _                 => "Manual Trigger",
            };
            StatusMessage = $"Session active — {SelectedDevice.Name} / {SelectedTemplate.Name}  [{modeLabel}]{triggerNote}";

            if (_scanMode == ScanMode.AutoPoll)
                _ = RunAutoPollLoopAsync(_pollCts.Token);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Start failed: {ex.Message}";
            await CleanupAsync();
        }
    }

    private async Task OnStopAsync()
    {
        StatusMessage = "Closing session…";
        _pollCts?.Cancel();

        // ── Step 1: stop the push listener so no new records arrive ──────────
        if (_pushHttpSubscriber is not null)
        {
            await _pushHttpSubscriber.StopAsync();
            _pushHttpSubscriber = null;
        }

        // ── Step 2: drain in-flight AcceptRecordAsync calls (max 2 s) ────────
        // OnPushRecord posts AcceptRecordAsync via fire-and-forget Dispatcher.InvokeAsync.
        // We must wait for all of them to finish before saving, otherwise a record that
        // arrived just before Stop is counted by the UI but missed from the XLSX.
        int drainMs = 0;
        while (System.Threading.Volatile.Read(ref _pendingAccept) > 0 && drainMs < 2000)
        {
            await Task.Delay(25);
            drainMs += 25;
        }

        // ── Step 3: save and close the session ────────────────────────────────
        // CloseSession() returns null on clean save, a rescue path if the primary
        // file was locked by Excel (all records still written to the rescue copy),
        // or "" if even the rescue failed (disk / permission problem).
        string? rescuePath = null;
        try
        {
            if (_sessionMgr is not null)
                rescuePath = await Task.Run(() => _sessionMgr.CloseSession());
        }
        finally
        {
            await CleanupAsync();
            IsRunning = false;
            StatusMessage = rescuePath switch
            {
                null => $"Session closed. {RecordCount} record(s) written.",
                ""   => $"⚠ Session closed — file locked by Excel and rescue save also failed. {RecordCount} record(s) may be lost.",
                _    => $"⚠ File was open in Excel — {RecordCount} record(s) saved to rescue copy: {rescuePath}",
            };
        }
    }

    // ── Manual / Push trigger ─────────────────────────────────────────────────

    private async Task OnTriggerAsync()
    {
        if (_sessionMgr is null) return;
        try
        {
            if (_scanMode == ScanMode.Push)
            {
                // Push mode has no persistent DMCC connection (DMST may be open for
                // live view). Open a brief connection, fire TRIGGER, then close.
                // The result arrives asynchronously via OnPushRecord.
                await SendPushTriggerAsync();
                return;
            }

            // Manual mode — synchronous DMCC trigger-and-wait.
            if (_deviceSession is null) return;
            var ctx = BuildContext();
            VerificationRecord? record = await _deviceSession.TriggerAndGetResultAsync(ctx);
            if (record is not null) await AcceptRecordAsync(record);
            else StatusMessage = "No read.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Trigger error: {ex.Message}";
        }
    }

    /// <summary>
    /// Fire a DMCC TRIGGER over a short-lived connection and immediately close.
    /// Used in Push mode where no persistent DMCC session is maintained.
    /// The push result arrives asynchronously via <see cref="OnPushRecord"/>.
    /// </summary>
    private async Task SendPushTriggerAsync()
    {
        if (SelectedDevice is null) return;

        var cfg = SelectedDevice.ToDeviceConfig();
        cfg.ConnectTimeoutMs  = 3_000;
        cfg.ResponseTimeoutMs = 5_000;

        // CRITICAL: raw DMCC text commands use port 23 (Telnet/DMCC), NOT port 44444.
        // Port 44444 is the DataMan SDK / HTTP-events port and uses the SDK's own binary
        // framing.  A bare TCP connection sending ||>TRIGGER ON\r\n to port 44444 is not
        // recognised as a DMCC session — the device silently ignores every command.
        // Port 23 is the standard Cognex DMCC raw text interface confirmed in the
        // DMCC Reference documentation.
        const int DmccRawPort = 23;

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-DMCC] Trigger attempt: {cfg.Host}:{DmccRawPort} (raw DMCC port)  " +
            $"connect={cfg.ConnectTimeoutMs}ms  response={cfg.ResponseTimeoutMs}ms  idle={cfg.IdleGapMs}ms");

        DeviceInterface.Dmcc.DmccResponse resp;
        try
        {
            using var tcp = new System.Net.Sockets.TcpClient();
            using var connectCts = new System.Threading.CancellationTokenSource(cfg.ConnectTimeoutMs);
            await tcp.ConnectAsync(cfg.Host, DmccRawPort, connectCts.Token);
            var stream = tcp.GetStream();

            // Drain welcome banner — port 23 on this device sends none; keep wait short.
            try
            {
                byte[] bannerBuf = new byte[512];
                using var bannerCts = new System.Threading.CancellationTokenSource(200);
                int nb = await stream.ReadAsync(bannerBuf, bannerCts.Token);
                if (nb > 0)
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-DMCC] Banner ({nb}B): " +
                        $"'{System.Text.Encoding.ASCII.GetString(bannerBuf, 0, nb).Trim()}'");
                else
                    System.Diagnostics.Debug.WriteLine("[VTCCP-DMCC] No banner received.");
            }
            catch (OperationCanceledException)
            {
                System.Diagnostics.Debug.WriteLine("[VTCCP-DMCC] Banner wait timed out (no banner).");
            }
            catch { }

            // Switch to Extended response mode.
            // The device starts in Silent mode (no ACK for any command including this one).
            // Mode takes effect immediately; subsequent commands will return ||[N]\r\n ACKs.
            await stream.WriteAsync(
                System.Text.Encoding.ASCII.GetBytes(
                    $"{DeviceInterface.Dmcc.DmccCommand.WireHeader}{DeviceInterface.Dmcc.DmccCommand.SetDmccResponseExtended}\r\n"));
            System.Diagnostics.Debug.WriteLine("[VTCCP-DMCC] SET COM.DMCC-RESPONSE 2 sent.");

            // Announce the trigger on the UI thread BEFORE firing it.
            // OnPushRecord will overwrite this message when the scan result arrives.
            // Setting it here (not after the ACK read) prevents the idle-gap timer
            // from expiring and overwriting a record message that arrived first.
            StatusMessage = "Trigger sent — waiting for push result…";

            // Send TRIGGER ON.
            // TRIGGER.TYPE=0 (Single/external) accepts software TRIGGER ON directly —
            // no trigger-type manipulation needed.
            await stream.WriteAsync(
                System.Text.Encoding.ASCII.GetBytes(
                    $"{DeviceInterface.Dmcc.DmccCommand.WireHeader}{DeviceInterface.Dmcc.DmccCommand.TriggerOn}\r\n"));
            System.Diagnostics.Debug.WriteLine("[VTCCP-DMCC] TRIGGER ON sent.");

            // Read ACK: expect ||[0]\r\n in Extended mode.
            // The scan result arrives separately via the HTTP subscriber.
            var sb = new System.Text.StringBuilder();
            byte[] buf = new byte[256];
            using var readCts = new System.Threading.CancellationTokenSource(cfg.ResponseTimeoutMs);
            try
            {
                while (true)
                {
                    using var idleCts = System.Threading.CancellationTokenSource
                        .CreateLinkedTokenSource(readCts.Token);
                    idleCts.CancelAfter(cfg.IdleGapMs);
                    int n = await stream.ReadAsync(buf, idleCts.Token);
                    if (n == 0) break;
                    sb.Append(System.Text.Encoding.ASCII.GetString(buf, 0, n));
                    if (sb.ToString().Contains("\r\n")) break;
                }
            }
            catch (OperationCanceledException) { }

            string raw = sb.ToString();
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-DMCC] TRIGGER raw response ({raw.Length}B): " +
                $"'{raw.Replace("\r", "\\r").Replace("\n", "\\n")}'");

            resp = DeviceInterface.Dmcc.DmccResponse.Parse(raw);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-DMCC] TRIGGER raw TCP exception: {ex.GetType().Name}: {ex.Message}");
            resp = DeviceInterface.Dmcc.DmccResponse.Parse(string.Empty);
        }

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-DMCC] TRIGGER response: code={resp.StatusCode}  " +
            $"body='{(resp.Body.Length > 80 ? resp.Body[..80] + "…" : resp.Body)}'");

        // On success (Ok) the message was already set to "Trigger sent — waiting…"
        // before the trigger fired, so OnPushRecord can overwrite it with the record
        // summary without being raced by the idle-gap timer expiring here.
        // Only update StatusMessage for non-Ok outcomes (errors the operator must see).
        if (resp.StatusCode != DeviceInterface.Dmcc.DmccStatus.Ok)
        {
            StatusMessage = resp.StatusCode switch
            {
                DeviceInterface.Dmcc.DmccStatus.NoRead =>
                    "Trigger fired — no symbol in field of view.",

                DeviceInterface.Dmcc.DmccStatus.Busy =>
                    "Device busy — trigger rejected. Wait a moment and retry.",

                DeviceInterface.Dmcc.DmccStatus.NoResponse =>
                    "Trigger: device sent no reply (code -2). " +
                    "Check VS Output for [VTCCP-DMCC] lines.",

                DeviceInterface.Dmcc.DmccStatus.Timeout =>
                    "Trigger: connection timed out (code -3). " +
                    "Verify the device IP/port and that the device is online.",

                DeviceInterface.Dmcc.DmccStatus.ParseError =>
                    "Trigger: unexpected response format from device (code -1). " +
                    "Check firmware version or DMCC port setting.",

                _ => string.IsNullOrWhiteSpace(resp.Body)
                        ? $"Trigger: device returned code {resp.StatusCode}."
                        : $"Trigger: device returned code {resp.StatusCode} — {resp.Body}",
            };
        }
    }

    // ── Auto-Poll background loop ─────────────────────────────────────────────

    private async Task RunAutoPollLoopAsync(System.Threading.CancellationToken ct)
    {
        while (!ct.IsCancellationRequested && IsRunning)
        {
            try
            {
                var ctx    = BuildContext();
                var record = await _deviceSession!.TriggerAndGetResultAsync(ctx, ct);
                if (record is not null)
                    await Application.Current.Dispatcher.InvokeAsync(() => _ = AcceptRecordAsync(record));
                else
                    Application.Current.Dispatcher.Invoke(() => StatusMessage = "No read — waiting…");
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() =>
                    StatusMessage = $"Poll error: {ex.Message}");
            }

            try { await Task.Delay(_autoPollIntervalMs, ct); }
            catch (OperationCanceledException) { break; }
        }
    }

    // ── Push (DMST) mode callback ─────────────────────────────────────────────

    // Called on thread-pool by DmstListener after each parsed push result.
    // The push script (v1.5) sends the complete quality XML in one TCP push, so no
    // secondary DMCC GET SYMBOL.RESULT fetch is needed or possible on this firmware.
    private void OnPushRecord(VerificationRecord pushRecord)
    {
        System.Threading.Interlocked.Increment(ref _pendingAccept);
        _ = Application.Current.Dispatcher.InvokeAsync(async () =>
        {
            try
            {
                await AcceptRecordInnerAsync(pushRecord);
            }
            catch { /* non-fatal write error */ }
            finally
            {
                System.Threading.Interlocked.Decrement(ref _pendingAccept);
            }
        });
    }

    // ── Shared record acceptance ──────────────────────────────────────────────

    /// <summary>
    /// Full acceptance: increments/decrements _pendingAccept around the inner write.
    /// Used by Manual and AutoPoll modes.
    /// </summary>
    private async Task AcceptRecordAsync(VerificationRecord record)
    {
        if (_sessionMgr is null) return;
        System.Threading.Interlocked.Increment(ref _pendingAccept);
        try   { await AcceptRecordInnerAsync(record); }
        finally { System.Threading.Interlocked.Decrement(ref _pendingAccept); }
    }

    /// <summary>
    /// Inner write — no _pendingAccept tracking (caller owns it).
    /// Used directly by OnPushRecord (which manages the counter itself across
    /// the whole DMCC fetch + write span).
    /// </summary>
    private async Task AcceptRecordInnerAsync(VerificationRecord record)
    {
        if (_sessionMgr is null) return;

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-WRITER] AcceptRecordInnerAsync: symbology={record.Symbology}  grade={record.OverallGrade?.LetterGradeString}");

        // ── OCR pass (L1 barcode-crop image) ─────────────────────────────────
        // TODO: replace _ocrEnabled with a per-session UI toggle (checkbox in
        //       SessionView) before shipping.  Defaulting true for dev/testing.
        if (_ocrEnabled && record.JpegImageBase64 is { Length: > 0 } b64)
        {
            try
            {
                byte[] jpegBytes = Convert.FromBase64String(b64);

                // UPC-A and EAN-13/EAN-8 are the canonical L1 OCR use cases:
                // the HRI digits are part of the barcode symbol and visible in
                // the barcode crop JPEG.  Pass the hint so DualEngineOcrRunner
                // applies BarcodeHriParser pattern extraction to both engines.
                bool isUpcEan = record.Symbology is { } sym &&
                    (sym.StartsWith("UPC", StringComparison.OrdinalIgnoreCase) ||
                     sym.StartsWith("EAN", StringComparison.OrdinalIgnoreCase));
                HriSymbologyHint hint = isUpcEan
                    ? HriSymbologyHint.UpcEan
                    : HriSymbologyHint.None;

                OcrResult ocrResult = await _ocrRunner.RunAsync(
                    jpegBytes, OcrImageSource.BarcodeCrop, hint);
                record = record with { OcrResult = ToOcrDto(ocrResult, record.DecodedData) };
            }
            catch { /* OCR failure must never block record acceptance */ }
        }

        // ── RFID cross-validation (awaited before Excel write) ────────────────
        // Awaiting here means the RFID result lands in the same Excel row as the
        // barcode grade — no row-update pass needed later.
        // Non-fatal: an RFID error never blocks barcode record acceptance.
        RfidValidationResult? rfidResult = null;
        if (_rfidCoordinator is { } rfidCoord)
        {
            try
            {
                rfidResult = await rfidCoord.OnBarcodeScannedAsync(
                    record, _pollCts?.Token ?? default).ConfigureAwait(false);
            }
            catch { /* RFID failure must never block record acceptance */ }
        }

        if (rfidResult is not null)
        {
            // Resolve the GCP prefix table date from the loaded validator (available when the
            // GcpLengthTable was successfully loaded at session start).  Format as yyyy-MM-dd
            // for the PDF provenance annotation, e.g. "From GCP prefix table as of 2026-05-03".
            string? gcpTableDate = _gcpValidator?.DataDate?.ToString("yyyy-MM-dd");

            record = record with
            {
                RfidStatus         = rfidResult.Status.ToString(),
                RfidEpcHex         = rfidResult.SelectedRead?.EpcHex,
                RfidEpcTagUri      = BuildEpcTagUri(rfidResult.ParsedEpc),
                RfidGtin14         = rfidResult.RfidGtin14,
                RfidSerial         = rfidResult.RfidSerial,
                RfidMismatchDetail = rfidResult.MismatchDetail,
                RfidScanWindowMs   = rfidResult.ScanWindowMs,
                RfidGcpValid       = rfidResult.GcpValid,
                RfidGcpTableDate   = gcpTableDate,
            };
        }

        // AddRecord is kept in a try/catch so that a Save failure (e.g. Excel
        // file locked, or session not yet fully started) never silently drops the
        // record from the history and UI.  savedToDisk controls the status message.
        bool savedToDisk = false;
        try
        {
            savedToDisk = await Task.Run(() => _sessionMgr.AddRecord(record));
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-WRITER] AddRecord complete.  savedToDisk={savedToDisk}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-WRITER] AddRecord threw {ex.GetType().Name}: {ex.Message}");
        }

        _history.AddRecord(record);
        _recordCount++; OnPropertyChanged(nameof(RecordCount));

        // ── Hybrid HTML report (fire-and-forget) ──────────────────────────────
        // Generates a self-contained report combining barcode grades + RFID data.
        // Runs on the thread-pool; failures are silently swallowed so they never
        // interfere with the scan loop.
        //
        // Alongside mode (default):
        //   Report lands in the session output dir (or HybridReportOutputDirectory).
        //
        // Replace mode:
        //   The original Webscan HTML was parsed and deleted by DmstHtmlScraper.
        //   The hybrid report is written back to the same path (same folder, same
        //   filename) so downstream tools watching the CodeQuality folder see only
        //   the enriched version.  The original file path is captured synchronously
        //   here (before firing the task) to avoid a race with the next incoming scan.
        if (_repo.Settings.GenerateHybridReport && _sessionOutputDir is { } sessionDir)
        {
            var hybridSettings = _repo.Settings;
            bool isReplace = hybridSettings.HybridReportMode == ConfigEngine.Models.HybridReportMode.Replace;

            // Capture the matched source path synchronously (before the task fires)
            // to prevent a concurrent scan from overwriting LastMatchedSourcePath.
            // For Manual/AutoPoll: DeviceSession's scraper already ran TryMergeAsync.
            // For Push mode:       _htmlWatcher is used inside the task (not yet set).
            string? capturedSourcePath = isReplace
                ? _deviceSession?.LastMatchedSourcePath
                : null;

            // Snapshot the watcher reference so the task closure is safe even if
            // CleanupAsync nulls _htmlWatcher between scheduling and execution.
            var htmlWatcherSnap = isReplace ? _htmlWatcher : null;

            _ = Task.Run(async () =>
            {
                try
                {
                    string? targetPath = capturedSourcePath;

                    if (isReplace && targetPath is null && htmlWatcherSnap is not null)
                    {
                        // Push mode Replace: wait for the Webscan HTML file to land and
                        // be parsed by the watcher (up to DmstHtmlScraper.FileArrivalTimeout).
                        // We discard the enriched record (HTTP already delivered full data),
                        // but TryMergeAsync sets LastMatchedSourcePath on the watcher.
                        await htmlWatcherSnap.TryMergeAsync(record);
                        targetPath = htmlWatcherSnap.LastMatchedSourcePath;
                    }

                    if (targetPath is { Length: > 0 })
                    {
                        // Replace mode — write hybrid to the same dir + filename as original
                        string dir      = Path.GetDirectoryName(targetPath)!;
                        string baseName = Path.GetFileNameWithoutExtension(targetPath);
                        await HybridReportGenerator.SaveAsync(record, dir, baseName);

                        System.Diagnostics.Debug.WriteLine(
                            $"[VTCCP-REPLACE] Hybrid report written → '{targetPath}'");
                    }
                    else
                    {
                        // Alongside mode, or Replace mode with no correlated HTML file
                        // (e.g. DMST extension is still .pdf): fall back to configured dir.
                        string reportDir = !string.IsNullOrWhiteSpace(hybridSettings.HybridReportOutputDirectory)
                            ? hybridSettings.HybridReportOutputDirectory
                            : sessionDir;
                        await HybridReportGenerator.SaveAsync(record, reportDir);
                    }
                }
                catch { /* report write failure must never affect the scan loop */ }
            });
        }

        // ── VCCS PDF report (fire-and-forget) ────────────────────────────────────
        // Non-fatal: PDF generation failures are caught inside GenerateAsync and
        // logged to Debug output.  The record has already been written to Excel above.
        if (_repo.Settings.GenerateVccsReport)
        {
            string pdfDir = !string.IsNullOrWhiteSpace(_repo.Settings.VccsReportOutputDirectory)
                ? _repo.Settings.VccsReportOutputDirectory
                : _sessionOutputDir;

            // WebscanSourcePath is non-null only when DMST correlated an HTML/PDF file;
            // MergeAsync inside GenerateAsync checks for .pdf extension automatically.
            string? wsPath = record.WebscanSourcePath;

            _ = DeviceInterface.Reports.PdfReportGenerator.GenerateAsync(
                record, pdfDir, wsPath, _pollCts?.Token ?? default);
        }

        string grade     = record.OverallGrade?.LetterGradeString is { Length: > 0 } g ? g : "?";
        string num       = record.OverallGrade?.NumericGrade is { } n ? $" ({n:F1})" : string.Empty;
        string ocrSuffix = record.OcrResult?.Tier is { Length: > 0 } t ? $"  | OCR: {t}" : string.Empty;
        string rfidSuffix = rfidResult?.Status switch
        {
            RfidValidationStatus.Pass                 => "  | RFID: ✓",
            RfidValidationStatus.Fail                 => "  | RFID: ✗ mismatch",
            RfidValidationStatus.NoTag                => "  | RFID: no tag",
            RfidValidationStatus.ParseError           => "  | RFID: parse error",
            RfidValidationStatus.MultipleTagsDetected => "  | RFID: multi-tag",
            _                                         => string.Empty,
        } ?? string.Empty;

        if (savedToDisk)
            StatusMessage = $"Record {RecordCount}: {record.Symbology} — {grade}{num}{ocrSuffix}{rfidSuffix}";
        else
            StatusMessage = $"⚠ Record {RecordCount}: {record.Symbology} — {grade}{num}{ocrSuffix}{rfidSuffix}  [file open in Excel — close Excel before ending session]";
    }

    /// <summary>
    /// Builds the EPC Tag URI (urn:epc:tag:...) from a decoded EPC.
    /// Returns null when the scheme is unknown or any required field is missing.
    /// </summary>
    private static string? BuildEpcTagUri(ParsedEpc? epc)
    {
        if (epc is null) return null;
        string? scheme = epc.Scheme switch
        {
            EpcScheme.Sgtin96  => "sgtin-96",
            EpcScheme.Sgtin198 => "sgtin-198",
            _                  => null,
        };
        if (scheme is null || epc.Filter is null || epc.CompanyPrefix is null
                           || epc.ItemReference is null || epc.Serial is null)
            return null;
        return $"urn:epc:tag:{scheme}:{epc.Filter}.{epc.CompanyPrefix}.{epc.ItemReference}.{epc.Serial}";
    }

    private static ExcelEngine.Models.OcrResultDto ToOcrDto(OcrResult r, string? encodedData)
    {
        // Compute OCR-vs-encoded-data match for UPC/EAN records.
        // Both sides are stripped to digits only so spaces, dashes, and the EAN
        // right-margin '>' do not cause false mismatches.
        string? match = null;
        if (r.ParsedDigits is { Length: > 0 } ocrDigits && encodedData is { Length: > 0 })
        {
            string encDigits = BarcodeHriParser.ExtractDigits(encodedData);
            if (encDigits.Length > 0)
                match = string.Equals(ocrDigits, encDigits, StringComparison.Ordinal)
                    ? "MATCH" : "MISMATCH";
        }

        return new ExcelEngine.Models.OcrResultDto
        {
            AgreedText          = r.AgreedText,
            Tier                = r.Tier.ToString(),
            WindowsText         = r.WindowsEngineText,
            WindowsConfidence   = r.WindowsEngineConfidence,
            TesseractText       = r.TesseractText,
            TesseractConfidence = r.TesseractConfidence,
            EditDistance        = r.EditDistance,
            EncodedDataMatch    = match,
        };
    }

    // ── UPC/EAN Supplemental read / write ─────────────────────────────────────

    // ── Live Feed window ──────────────────────────────────────────────────────

    /// <summary>
    /// Opens the Live Feed window for the currently selected device.
    /// If a window is already open for this device, brings it to the front instead
    /// of opening a second instance.
    /// </summary>
    private void OnOpenLiveFeed()
    {
        if (SelectedDevice is null) return;

        if (_liveFeedWindow is not null)
        {
            _liveFeedWindow.Activate();
            return;
        }

        var vm     = new LiveFeedViewModel(SelectedDevice.Host, SelectedDevice.Port);
        var window = new LiveFeedWindow(vm) { Owner = Application.Current.MainWindow };
        _liveFeedWindow = window;
        window.Closed += (_, _) => _liveFeedWindow = null;
        window.Show();
    }

    /// <summary>
    /// Opens the Symbol Stitching window for the currently selected device.
    /// Passes <see cref="_deviceSession"/> so Verify can call IMAGE.LOAD;
    /// null is accepted — capture and preview still work, Verify is disabled.
    /// </summary>
    private void OnOpenStitching()
    {
        if (SelectedDevice is null) return;

        if (_stitchingWindow is not null)
        {
            _stitchingWindow.Activate();
            return;
        }

        var vm     = new StitchingViewModel(SelectedDevice.Host, _deviceSession);
        var window = new StitchingWindow(vm) { Owner = Application.Current.MainWindow };
        _stitchingWindow = window;
        window.Closed += (_, _) => _stitchingWindow = null;
        window.Show();
    }

    // ── Supplemental commands ─────────────────────────────────────────────────

    private async Task OnReadSupplementalAsync()
    {
        if (SelectedDevice is null) return;
        _isApplyingSupplemental = true;
        RelayCommand.Refresh();
        SupplementalStatus = "Reading from device…";
        try
        {
            var cfg = SelectedDevice.ToDeviceConfig();
            cfg.ConnectTimeoutMs  = 3_000;
            cfg.ResponseTimeoutMs = 3_000;

            await using var client = new DeviceInterface.Dmcc.DataManSdkClient(cfg);
            await client.ConnectAsync();
            var resp = await client.SendAsync(DeviceInterface.Dmcc.DmccCommand.GetUpcEanSupplemental);

            if (resp.StatusCode == DeviceInterface.Dmcc.DmccStatus.Ok
                && int.TryParse(resp.Body.Trim(), out int mode)
                && mode is >= 0 and <= 5)
            {
                UpcEanSupplementalMode = mode;
                SupplementalStatus = $"Read OK — firmware: {SupplementalModeLabel(mode)}";
            }
            else
            {
                SupplementalStatus =
                    $"Read failed (code {resp.StatusCode}). " +
                    "DMCC key may differ on this firmware — see DmccCommand.cs note.";
            }
        }
        catch (Exception ex)
        {
            SupplementalStatus = $"Read error: {ex.Message}";
        }
        finally
        {
            _isApplyingSupplemental = false;
            RelayCommand.Refresh();
        }
    }

    /// <summary>
    /// Opens a short-lived DMCC connection and writes the selected supplemental
    /// mode to firmware (persistent — survives power cycle).
    /// Works independently of whether a session is running.
    /// </summary>
    private async Task OnWriteSupplementalAsync()
    {
        if (SelectedDevice is null) return;
        _isApplyingSupplemental = true;
        RelayCommand.Refresh();
        int mode = _upcEanSupplementalMode;
        SupplementalStatus = $"Writing \"{SupplementalModeLabel(mode)}\" to firmware…";
        try
        {
            var cfg = SelectedDevice.ToDeviceConfig();
            cfg.ConnectTimeoutMs  = 3_000;
            cfg.ResponseTimeoutMs = 3_000;

            await using var client = new DeviceInterface.Dmcc.DataManSdkClient(cfg);
            await client.ConnectAsync();
            var resp = await client.SendAsync(DeviceInterface.Dmcc.DmccCommand.SetUpcEanSupplemental(mode));

            SupplementalStatus = resp.StatusCode == DeviceInterface.Dmcc.DmccStatus.Ok
                ? $"Written — firmware now: {SupplementalModeLabel(mode)} (persistent)"
                : $"Write failed (code {resp.StatusCode}). " +
                  "DMCC key may differ on this firmware — see DmccCommand.cs note.";
        }
        catch (Exception ex)
        {
            SupplementalStatus = $"Write error: {ex.Message}";
        }
        finally
        {
            _isApplyingSupplemental = false;
            RelayCommand.Refresh();
        }
    }

    private static string SupplementalModeLabel(int mode) => mode switch
    {
        0 => "Ignore",
        1 => "Required",
        2 => "Required 2-digit",
        3 => "Required 5-digit",
        4 => "Not Required",
        _ => $"Unknown ({mode})",
    };

    // ── Helpers ───────────────────────────────────────────────────────────────

    private VerificationRecord BuildContext() => new()
    {
        Symbology       = string.Empty,
        DeviceSerial    = _deviceSession?.DeviceInfo.Serial    ?? string.Empty,
        DeviceName      = _deviceSession?.DeviceInfo.Name      ?? string.Empty,
        FirmwareVersion = _deviceSession?.DeviceInfo.FirmwareVersion ?? string.Empty,
        SoftwareVersion = _deviceSession?.DeviceInfo.SoftwareVersion,
        OperatorId      = string.IsNullOrWhiteSpace(OperatorOverride)
                              ? (SelectedTemplate?.OperatorId ?? string.Empty)
                              : OperatorOverride.Trim(),
        JobName         = SelectedTemplate?.JobName ?? string.Empty,
        LogoPath        = SelectedTemplate?.LogoPath,
    };

    /// <summary>
    /// Called by App.OnExit to ensure the DMCC connection is closed and
    /// TRIGGER.TYPE is restored to the DMST-native value before the process exits.
    /// Safe to call when no session is running (no-op).
    /// </summary>
    public async Task StopSessionOnExitAsync()
    {
        if (!IsRunning) return;
        _pollCts?.Cancel();
        await CleanupAsync();
        IsRunning = false;
    }

    private async Task CleanupAsync()
    {
        if (_deviceSession is not null)
        {
            await _deviceSession.DisposeAsync();
            _deviceSession = null;
        }
        if (_pushHttpSubscriber is not null)
        {
            await _pushHttpSubscriber.StopAsync();
            _pushHttpSubscriber = null;
        }
        if (_htmlWatcher is not null)
        {
            _htmlWatcher.Stop();
            _htmlWatcher = null;
        }
        if (_rfidCoordinator is not null)
        {
            await _rfidCoordinator.DisposeAsync();
            _rfidCoordinator = null;
        }
        _gcpValidator = null;
        _sessionMgr?.Dispose();
        _sessionMgr = null;
        _pollCts?.Dispose();
        _pollCts = null;
    }

}
