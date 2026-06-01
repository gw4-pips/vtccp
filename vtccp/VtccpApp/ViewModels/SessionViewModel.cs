namespace VtccpApp.ViewModels;

using System.Collections.ObjectModel;
using System.Windows;
using ConfigEngine;
using ConfigEngine.Models;
using DeviceInterface;
using DeviceInterface.Dmst;
using ExcelEngine.Models;
using ExcelEngine.Schema;
using ExcelEngine.Session;
using OcrEngine;
using VtccpApp.Commands;

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

        Reload();
    }

    // ── Reload from repository ────────────────────────────────────────────────

    public void Reload()
    {
        AvailableDevices.Clear();
        AvailableTemplates.Clear();
        foreach (var d in _repo.Devices)   AvailableDevices.Add(d);
        foreach (var t in _repo.Templates) AvailableTemplates.Add(t);

        if (SelectedDevice is null)   SelectedDevice   = _repo.DefaultDevice;
        if (SelectedTemplate is null) SelectedTemplate = _repo.DefaultTemplate;

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
                };
                _pushHttpSubscriber = new DeviceInterface.Dmst.HttpEventSubscriber(
                    cfg.Host, cfg.Port, _xmlMap, ctx, OnPushRecord);
                await _pushHttpSubscriber.StartAsync(_pollCts.Token);
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
                byte[]    jpegBytes = Convert.FromBase64String(b64);
                OcrResult ocrResult = await _ocrRunner.RunAsync(
                    jpegBytes, OcrImageSource.BarcodeCrop);
                record = record with { OcrResult = ToOcrDto(ocrResult) };
            }
            catch { /* OCR failure must never block record acceptance */ }
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
        string grade     = record.OverallGrade?.LetterGradeString is { Length: > 0 } g ? g : "?";
        string num       = record.OverallGrade?.NumericGrade is { } n ? $" ({n:F1})" : string.Empty;
        string ocrSuffix = record.OcrResult?.Tier is { Length: > 0 } t ? $"  | OCR: {t}" : string.Empty;
        if (savedToDisk)
            StatusMessage = $"Record {RecordCount}: {record.Symbology} — {grade}{num}{ocrSuffix}";
        else
            StatusMessage = $"⚠ Record {RecordCount}: {record.Symbology} — {grade}{num}{ocrSuffix}  [file open in Excel — close Excel before ending session]";
    }

    private static ExcelEngine.Models.OcrResultDto ToOcrDto(OcrResult r) =>
        new()
        {
            AgreedText          = r.AgreedText,
            Tier                = r.Tier.ToString(),
            WindowsText         = r.WindowsEngineText,
            WindowsConfidence   = r.WindowsEngineConfidence,
            TesseractText       = r.TesseractText,
            TesseractConfidence = r.TesseractConfidence,
            EditDistance        = r.EditDistance,
            ImageSource         = r.ImageSource.ToString(),
        };

    // ── UPC/EAN Supplemental read / write ─────────────────────────────────────

    /// <summary>
    /// Opens a short-lived DMCC connection to read the current supplemental mode
    /// from firmware and update the radio button selection to match.
    /// Works independently of whether a session is running.
    /// </summary>
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
        OperatorId      = string.IsNullOrWhiteSpace(OperatorOverride)
                              ? (SelectedTemplate?.OperatorId ?? string.Empty)
                              : OperatorOverride.Trim(),
        JobName         = SelectedTemplate?.JobName ?? string.Empty,
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
        _sessionMgr?.Dispose();
        _sessionMgr = null;
        _pollCts?.Dispose();
        _pollCts = null;
    }
}
