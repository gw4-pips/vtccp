namespace VtccpApp.ViewModels;

using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using ConfigEngine;
using ConfigEngine.Models;
using DeviceInterface;
using DeviceInterface.Dmcc;
using DeviceInterface.Dmst;
using DeviceInterface.Reports;
using DeviceInterface.Rfid;
using DeviceInterface.Rfid.Gcp;
using DeviceInterface.Rfid.Models;
using DeviceInterface.Validation;
using DeviceInterface.Webscan;
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

    public enum ScanMode { Manual, AutoPoll, Push, Webscan }

    // ── Dependencies ──────────────────────────────────────────────────────────

    private readonly ConfigRepository  _repo;
    private readonly HistoryViewModel  _history;
    private readonly VerificationXmlMap _xmlMap = new();

    // ── Runtime state ─────────────────────────────────────────────────────────

    private DeviceSession?                              _deviceSession;
    private DeviceInterface.Dmst.HttpEventSubscriber?  _pushHttpSubscriber;
    private WebscanHtmlFileAdapter?                     _webscanHtmlAdapter;
    private EventHandler<VerificationRecord>?           _webscanRecordHandler;
    private EventHandler<string>?                       _webscanParseFailedHandler;
    private int                                         _webscanSessionGeneration;
    private readonly object                             _webscanStateLock = new();
    private readonly WebscanAcceptanceTracker           _webscanAcceptance = new();
    private readonly System.Threading.SemaphoreSlim    _pushTruCheckSettingsGate = new(1, 1);
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

    /// <summary>
    /// Persistent RFID reader — outlives individual sessions so the operator can
    /// connect once via the Session Launcher panel and run many sessions.
    /// Owned by the ViewModel (coordinator is created with ownsReader:false).
    /// </summary>
    private IEpcReader? _rfidReader;

    /// <summary>
    /// Per-session writer for the "RFID Scans" auxiliary worksheet.
    /// Created lazily on the first RFID result of a session (needs the session's
    /// live Excel adapter); reset in <see cref="CleanupAsync"/>.
    /// </summary>
    private RfidTabWriter? _rfidTabWriter;
    private bool        _isRfidConnected;
    private bool        _isRfidBusy;
    private string      _rfidStatusMessage = "Not connected.";
    private string?     _selectedRfidPort;

    // ── Scan result display ───────────────────────────────────────────────────
    private string  _verifierResultLine = string.Empty;
    private string? _rfidResultLine;
    private string? _webscanImportError;

    // ── Session output directory ───────────────────────────────────────────────
    // Captured at OnStartAsync so fire-and-forget report generators have the path
    // without relying on mutable session-manager state after the session closes.
    private string _sessionOutputDir = string.Empty;

    // Session identifier (start-time stamp) — suffix for per-scan VCCS PDF filenames.
    private string _sessionId = string.Empty;

    /// <summary>
    /// FileSystem watcher (DmstHtmlScraper) started only in Push mode when
    /// <see cref="ConfigEngine.Models.HybridReportMode.Replace"/> is active.
    /// Watches the CodeQuality folder, preserves the Webscan HTML files, and makes
    /// them available by literal Verified-value correlation so AcceptRecordInnerAsync
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
    public ObservableCollection<string>        AvailableRfidPorts { get; } = [];

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

            // Keep the operator's selected mode when changing device profiles.
            // Push is the one mode whose availability depends on the profile;
            // fall back to Manual rather than leaving the UI on an unusable mode.
            if (!IsRunning && _scanMode == ScanMode.Push && !IsPushAvailable)
                ActiveScanMode = ScanMode.Manual;
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
            OnPropertyChanged(nameof(IsWebscanMode));
            OnPropertyChanged(nameof(ShowTriggerButton));
            RelayCommand.Refresh();
        }
    }

    public bool IsManualMode     => _scanMode == ScanMode.Manual;
    public bool IsAutoPollMode   => _scanMode == ScanMode.AutoPoll;
    public bool IsPushMode       => _scanMode == ScanMode.Push;
    public bool IsWebscanMode    => _scanMode == ScanMode.Webscan;

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

    // ── RFID properties ───────────────────────────────────────────────────────

    /// <summary>COM port selected in the RFID panel, e.g. "COM4". Persisted to AppSettings.</summary>
    public string? SelectedRfidPort
    {
        get => _selectedRfidPort;
        set { Set(ref _selectedRfidPort, value); RelayCommand.Refresh(); }
    }

    public bool IsRfidConnected
    {
        get => _isRfidConnected;
        private set { Set(ref _isRfidConnected, value); RelayCommand.Refresh(); }
    }

    public string RfidStatusMessage
    {
        get => _rfidStatusMessage;
        private set => Set(ref _rfidStatusMessage, value);
    }

    // ── Scan result display properties ───────────────────────────────────────

    /// <summary>Verifier result line shown in the result box (symbology, grade, OCR).</summary>
    public string VerifierResultLine
    {
        get => _verifierResultLine;
        private set { Set(ref _verifierResultLine, value); OnPropertyChanged(nameof(HasScanResult)); }
    }

    /// <summary>RFID result line shown below the verifier line; null when no RFID ran.</summary>
    public string? RfidResultLine
    {
        get => _rfidResultLine;
        private set { Set(ref _rfidResultLine, value); OnPropertyChanged(nameof(HasRfidResult)); }
    }

    public string? WebscanImportError
    {
        get => _webscanImportError;
        private set { Set(ref _webscanImportError, value); OnPropertyChanged(nameof(HasWebscanImportError)); }
    }

    public bool HasWebscanImportError => !string.IsNullOrWhiteSpace(_webscanImportError);

    /// <summary>True when a scan result is ready to display.</summary>
    public bool HasScanResult => !string.IsNullOrEmpty(_verifierResultLine);

    /// <summary>True when an RFID result is available to display below the verifier line.</summary>
    public bool HasRfidResult => !string.IsNullOrEmpty(_rfidResultLine);

    /// <summary>Waiting hint shown in the bottom-right of the result box while the session is running.</summary>
    public string WaitingMessage => _recordCount == 0
        ? "Waiting for first verification scan…"
        : "Waiting for next scan…";

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
    public RelayCommand SetWebscanCommand       { get; }
    public RelayCommand ReadSupplementalCommand  { get; }
    public RelayCommand WriteSupplementalCommand { get; }
    public RelayCommand OpenLiveFeedCommand      { get; }
    public RelayCommand OpenStitchingCommand     { get; }
    public RelayCommand RefreshRfidPortsCommand  { get; }
    public RelayCommand FindRfidPortCommand      { get; }
    public RelayCommand ConnectRfidCommand       { get; }
    public RelayCommand DisconnectRfidCommand    { get; }

    public SessionViewModel(ConfigRepository repo, HistoryViewModel history)
    {
        _repo    = repo;
        _history = history;

        StartCommand   = new RelayCommand(async () => await OnStartAsync(),
            () => !IsRunning && SelectedTemplate is not null &&
                  (_scanMode == ScanMode.Webscan || SelectedDevice is not null));
        StopCommand    = new RelayCommand(async () => await OnStopAsync(),
            () => IsRunning);
        TriggerCommand = new RelayCommand(async () => await OnTriggerAsync(),
            () => IsRunning && (_scanMode == ScanMode.Manual || _scanMode == ScanMode.Push));

        SetManualCommand   = new RelayCommand(() => SelectScanMode(ScanMode.Manual),   () => !IsRunning);
        SetAutoPollCommand = new RelayCommand(() => SelectScanMode(ScanMode.AutoPoll), () => !IsRunning);
        SetPushCommand     = new RelayCommand(() => SelectScanMode(ScanMode.Push),     () => !IsRunning && IsPushAvailable);
        SetWebscanCommand  = new RelayCommand(() => SelectScanMode(ScanMode.Webscan),  () => !IsRunning);

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

        RefreshRfidPortsCommand = new RelayCommand(
            RefreshRfidPorts,
            () => !_isRfidBusy);

        FindRfidPortCommand = new RelayCommand(
            FindRfidPort,
            () => !_isRfidBusy && !IsRfidConnected);

        ConnectRfidCommand = new RelayCommand(
            async () => await ConnectRfidAsync(),
            () => !IsRfidConnected && !_isRfidBusy &&
                  !string.IsNullOrWhiteSpace(SelectedRfidPort));

        DisconnectRfidCommand = new RelayCommand(
            async () => await DisconnectRfidAsync(),
            () => IsRfidConnected && !_isRfidBusy);

        Reload();
        RefreshRfidPorts();
        AutoFindRfidPortOnStartup();
    }

    // ── Reload from repository ────────────────────────────────────────────────

    public void Reload(bool settingsLoaded = false)
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

        if (settingsLoaded)
        {
            RestoreLastScanMode();
        }
    }

    private void SelectScanMode(ScanMode mode)
    {
        ActiveScanMode = mode;

        string modeName = mode.ToString();
        if (_repo.Settings.LastScanMode != modeName)
        {
            _repo.Settings.LastScanMode = modeName;
            _ = _repo.SaveSettingsAsync();
        }
    }

    /// <summary>
    /// Restores the last operator-selected scan mode after the repository has
    /// loaded. Missing or invalid values retain the original first-run default.
    /// </summary>
    private void RestoreLastScanMode()
    {
        if (Enum.TryParse<ScanMode>(
                _repo.Settings.LastScanMode,
                ignoreCase: true,
                out var savedMode))
        {
            ActiveScanMode = savedMode == ScanMode.Push && !IsPushAvailable
                ? ScanMode.Manual
                : savedMode;
            return;
        }

        // Preserve the pre-persistence startup behavior for existing installs.
        ActiveScanMode = IsPushAvailable ? ScanMode.Push : ScanMode.Manual;
    }

    // ── Start / Stop ──────────────────────────────────────────────────────────

    private async Task OnStartAsync()
    {
        var selectedTemplate = SelectedTemplate;
        var selectedDevice = SelectedDevice;
        if (selectedTemplate is null ||
            (_scanMode != ScanMode.Webscan && selectedDevice is null))
            return;

        string outputDir = !string.IsNullOrWhiteSpace(selectedTemplate.OutputDirectory)
            ? selectedTemplate.OutputDirectory
            : _repo.Settings.DefaultOutputDirectory;
        _sessionOutputDir = outputDir;   // captured for hybrid report generation
        _sessionId        = DateTime.Now.ToString("yyyyMMdd-HHmmss");

        SessionState state = selectedTemplate.ToSessionState(outputDir);
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
            if (_scanMode == ScanMode.Webscan)
            {
                // Webscan TruCheck is USB-connected. The independent HTML
                // adapter is started after the Excel session opens below so a
                // newly-created report can never race a closed record writer.
                StatusMessage = "Preparing Webscan HTML import…";
            }
            else if (_scanMode == ScanMode.Push)
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
                var cfg = selectedDevice!.ToDeviceConfig();
                var ctx = new VerificationRecord
                {
                    Symbology       = string.Empty,
                    DeviceSerial    = string.Empty,
                    DeviceName      = selectedDevice.Name,
                    FirmwareVersion = string.Empty,
                    OperatorId      = state.OperatorId  ?? string.Empty,
                    JobName         = state.JobName      ?? string.Empty,
                    BatchNumber     = state.BatchNumber  ?? string.Empty,
                    CompanyName     = state.CompanyName  ?? string.Empty,
                    LogoPath        = state.LogoPath,
                };
                // This Push subscriber is the DataMan HTTP path. Webscan
                // TruChecks are USB-connected and use a separate result path.
                _pushHttpSubscriber = new DeviceInterface.Dmst.HttpEventSubscriber(
                    cfg.Host,
                    cfg.Port,
                    _xmlMap,
                    ctx,
                    OnPushRecord,
                    Environment.GetEnvironmentVariable("VTCCP_HTTP_CAPTURE_DIR"));
                await _pushHttpSubscriber.StartAsync(_pollCts.Token);

                // ── HTML provenance watcher in Push mode ─────────────────────
                // In Push mode, DeviceSession (and its built-in DmstHtmlScraper) is
                // not used.  A VCCS PDF requires a real local HTML artifact, so start
                // the same watcher whenever either a strict PDF or Replace-mode hybrid
                // report needs that artifact. Correlation is by the verifier's literal
                // Verified: value only — never the filename or a timestamp tolerance.
                bool isReplaceHybrid =
                    _repo.Settings.GenerateHybridReport &&
                    _repo.Settings.HybridReportMode == ConfigEngine.Models.HybridReportMode.Replace;
                bool needsFilesystemHtml = _repo.Settings.GenerateVccsReport || isReplaceHybrid;
                if (needsFilesystemHtml)
                {
                    var watchPath = DeviceInterface.Dmst.DmstHtmlScraper.ConfiguredReportDirectory;
                    _htmlWatcher = new DeviceInterface.Dmst.DmstHtmlScraper(watchPath);
                    _htmlWatcher.Start();
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-PROVENANCE] HTML watcher started: '{watchPath}'");
                }
            }
            else
            {
                // Manual / AutoPoll: open DMCC connection (requires DMST to be closed).
                StatusMessage = "Connecting to device…";
                var cfg = selectedDevice!.ToDeviceConfig();
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

            if (_scanMode == ScanMode.Webscan)
            {
                int generation = System.Threading.Interlocked.Increment(
                    ref _webscanSessionGeneration);
                _webscanHtmlAdapter = new WebscanHtmlFileAdapter();
                _webscanAcceptance.BeginSession(generation);
                _webscanRecordHandler = (_, record) => OnWebscanRecord(generation, record);
                _webscanParseFailedHandler = (_, message) =>
                    OnWebscanParseFailed(generation, message);
                _webscanHtmlAdapter.RecordParsed += _webscanRecordHandler;
                _webscanHtmlAdapter.ParseFailed += _webscanParseFailedHandler;
                WebscanImportError = null;
                _webscanHtmlAdapter.Start();
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-WEBSCAN] HTML watcher started: '{_webscanHtmlAdapter.WatchDirectory}'");
            }

            // ── RFID coordinator (optional) ───────────────────────────────────
            // Runs when the reader is connected via the Session Launcher RFID panel.
            // If a COM port is selected but not yet connected, auto-connect now.
            // Connection failure is non-fatal — session continues without RFID.
            if (!IsRfidConnected && !string.IsNullOrWhiteSpace(SelectedRfidPort))
                await ConnectRfidAsync();

            if (IsRfidConnected && _rfidReader is not null)
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

                    // Reader lifetime is owned by the ViewModel (Connect/Disconnect
                    // panel) — ownsReader:false keeps the reader connected across
                    // sessions when the coordinator is disposed at session close.
                    var validator = new RfidValidator(_gcpValidator);
                    _rfidCoordinator = new RfidScanCoordinator(
                        _rfidReader, validator, rfidSettings, ownsReader: false);
                    System.Diagnostics.Debug.WriteLine($"[RFID] Coordinator started on {SelectedRfidPort}.");
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
                ScanMode.Push     => $"Push (DMST) — port {selectedDevice?.DmstListenPort ?? 0}",
                ScanMode.Webscan  => $"Webscan HTML (USB) — {WebscanHtmlFileAdapter.ConfiguredReportDirectory}",
                _                 => "Manual Trigger",
            };
            string deviceLabel = selectedDevice?.Name ?? "Webscan TruCheck";
            StatusMessage = $"Session active — {deviceLabel} / {selectedTemplate.Name}  [{modeLabel}]{triggerNote}";

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
        await QuiesceWebscanAsync();

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
            VerifierResultLine = string.Empty;
            RfidResultLine     = null;
            StatusMessage = rescuePath switch
            {
                null => $"Session closed. {RecordCount} record(s) written.",
                ""   => $"⚠ Session closed — file locked by Excel and rescue save also failed. {RecordCount} record(s) may be lost.",
                _    => $"⚠ File was open in Excel — {RecordCount} record(s) saved to rescue copy: {rescuePath}",
            };
        }
    }

    // ── RFID connect / disconnect ─────────────────────────────────────────────

    /// <summary>
    /// Runs silently at startup: tries the VID/PID registry lookup and updates the
    /// RFID status note so the operator sees "ASR-P35U found on COMx" or
    /// "Last used: COMx — reader not detected" without touching anything.
    /// Must be called after <see cref="RefreshRfidPorts"/> so the port list is
    /// already populated.
    /// </summary>
    private void AutoFindRfidPortOnStartup()
    {
#if ASREADER_SDK
        var found = EpcReaderFactory.FindAsReaderPort();
#else
        string? found = null;
#endif
        if (found is not null && AvailableRfidPorts.Contains(found))
        {
            // Device is plugged in — pre-select it and show a clear note.
            SelectedRfidPort  = found;
            RfidStatusMessage = $"ASR-P35U found on {found}.";
        }
        else if (!string.IsNullOrWhiteSpace(SelectedRfidPort))
        {
            // Device not detected but we have a previously-used port from AppSettings.
            RfidStatusMessage = $"Last used: {SelectedRfidPort} — reader not detected.";
        }
        // When neither condition holds the default "Not connected." message stands.
    }

    /// <summary>
    /// Re-enumerates the machine's COM ports into <see cref="AvailableRfidPorts"/>,
    /// preserving the current selection when the port is still present.
    /// Pre-selects the port persisted in AppSettings on first population.
    /// </summary>
    private void FindRfidPort()
    {
        RefreshRfidPorts();
#if ASREADER_SDK
        var found = EpcReaderFactory.FindAsReaderPort();
#else
        string? found = null;
#endif
        if (found is not null && AvailableRfidPorts.Contains(found))
        {
            SelectedRfidPort   = found;
            RfidStatusMessage  = $"Found ASR-P35U on {found}.";
        }
        else
        {
            RfidStatusMessage = "ASR-P35U not found — check USB connection.";
        }
    }

    private void RefreshRfidPorts()
    {
        string? current = SelectedRfidPort ?? _repo.Settings.RfidComPort;

        AvailableRfidPorts.Clear();
        // SerialPort.GetPortNames() directly — AsReaderP35UEpcReader/EpcReaderFactory
        // are excluded from compilation when the SDK DLL is absent, but the port
        // picker must still work (it shows what WOULD be used once the DLL lands).
        foreach (var p in System.IO.Ports.SerialPort.GetPortNames()
                     .OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            AvailableRfidPorts.Add(p);

        SelectedRfidPort = current is not null && AvailableRfidPorts.Contains(current)
            ? current
            : AvailableRfidPorts.FirstOrDefault();

        if (AvailableRfidPorts.Count == 0 && !IsRfidConnected)
            RfidStatusMessage = "No COM ports found — plug in the ASR-P35U and refresh.";
    }

    /// <summary>
    /// Connects the ASR-P35U on <see cref="SelectedRfidPort"/> and persists the
    /// port to AppSettings so future sessions auto-connect. Safe to call when
    /// already connected (no-op). Failure is non-fatal: status message is set
    /// and the app continues without RFID.
    /// </summary>
    private async Task ConnectRfidAsync()
    {
        if (_isRfidBusy || IsRfidConnected) return;
        if (string.IsNullOrWhiteSpace(SelectedRfidPort))
        {
            RfidStatusMessage = "Select a COM port first.";
            return;
        }

        _isRfidBusy = true;
        RelayCommand.Refresh();
#if ASREADER_SDK
        IEpcReader? reader = null;
        try
        {
            RfidStatusMessage = $"Connecting to {SelectedRfidPort}…";
            reader = EpcReaderFactory.CreateAsReaderP35U();
            await reader.ConnectAsync(SelectedRfidPort);
            _rfidReader     = reader;
            IsRfidConnected = true;
            RfidStatusMessage = $"Connected on {SelectedRfidPort}.";

            // Persist so the port pre-selects next launch and auto-connect works.
            if (_repo.Settings.RfidComPort != SelectedRfidPort)
            {
                _repo.Settings.RfidComPort = SelectedRfidPort;
                _ = _repo.SaveSettingsAsync();
            }
        }
        catch (Exception ex)
        {
            RfidStatusMessage = $"Connect failed: {ex.Message}";
            if (reader is not null)
                try { await reader.DisposeAsync(); } catch { /* best effort */ }
            _rfidReader     = null;
            IsRfidConnected = false;
        }
#else
        RfidStatusMessage =
            "AsReader SDK DLL not present in this build — RFID unavailable.";
        await Task.CompletedTask;
#endif
        _isRfidBusy = false;
        RelayCommand.Refresh();
    }

    /// <summary>
    /// Disconnects the reader. If a session is running with an active RFID
    /// coordinator, the coordinator is disposed first so no scan window opens
    /// against a dead reader; the barcode session itself keeps running.
    /// </summary>
    private async Task DisconnectRfidAsync()
    {
        if (_isRfidBusy || !IsRfidConnected) return;
        _isRfidBusy = true;
        RelayCommand.Refresh();
        try
        {
            if (_rfidCoordinator is not null)
            {
                await _rfidCoordinator.DisposeAsync();   // ownsReader:false — reader survives
                _rfidCoordinator = null;
            }
            if (_rfidReader is not null)
            {
                try { await _rfidReader.DisposeAsync(); } catch { /* best effort */ }
                _rfidReader = null;
            }
            IsRfidConnected   = false;
            RfidStatusMessage = "Disconnected.";
        }
        finally
        {
            _isRfidBusy = false;
            RelayCommand.Refresh();
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

        // ── TC Live cancel (DISABLED — no effect found) ───────────────────────────
        // Problem: DMST TC panel "Go Live" (monitor mode) blocks software triggers.
        //
        // Attempted: SET MONITOR-MODE.ENABLE OFF on port 23 — sent successfully but
        // produced no discernible effect on live mode.
        //
        // Attempted earlier: HTTP GET /monitormode?enable=false on port 44444 — port
        // 44444 requires the DataMan SDK handshake and rejects bare HTTP connections.
        //
        // Possible next approach: Wireshark the SDK connection during a Go Live →
        // trigger sequence to determine whether DMST sends a proprietary SDK command
        // (not raw DMCC text) that VTCCP would need to replicate, or whether the
        // device simply needs operator action to exit live mode before triggering.
        //
        // DmccCommand.SetMonitorModeOff / SetMonitorModeOn constants are available
        // if a working invocation path is found later.

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
                else if (RecordCount == 0 && _pendingAccept == 0)
                    // Only show waiting state before the first scan and only when no
                    // accept is in-flight — prevents race during the RFID scan window.
                    Application.Current.Dispatcher.Invoke(() => StatusMessage = "Waiting for first verification scan…");
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

    // ── Push-mode per-result TruCheck settings ─────────────────────────────────

    private async Task<VerificationRecord> AttachPushTruCheckSettingsAsync(
        VerificationRecord record,
        System.Threading.CancellationToken ct)
    {
        if (SelectedDevice is null)
            return record;

        // Push mode intentionally has no persistent SDK connection so DMST can
        // remain open. Take a short raw-DMCC snapshot after each pushed result
        // instead. The gate prevents overlapping HTTP callbacks from issuing
        // concurrent commands through the reader's raw port.
        await _pushTruCheckSettingsGate.WaitAsync(ct);
        try
        {
            var cfg = SelectedDevice.ToDeviceConfig();
            using var tcp = new System.Net.Sockets.TcpClient();
            using var connectCts =
                System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct);
            connectCts.CancelAfter(cfg.ConnectTimeoutMs);
            await tcp.ConnectAsync(cfg.Host, DmccCommand.RawDmccPort, connectCts.Token);

            var stream = tcp.GetStream();

            // A raw port-23 session starts silent. Enable extended replies before
            // the GET commands; the SET itself has no reply in silent mode.
            await WriteRawDmccCommandAsync(
                stream, DmccCommand.SetDmccResponseExtended, ct);

            DmccResponse applicationStandard = await SendRawDmccCommandAsync(
                stream, cfg, DmccCommand.GetApplicationStandard, ct);
            DmccResponse dataFormatCheck = await SendRawDmccCommandAsync(
                stream, cfg, DmccCommand.GetCustomDataParsingStandard, ct);
            DmccResponse apertureSetting = await SendRawDmccCommandAsync(
                stream, cfg, DmccCommand.GetAperture, ct);

            var enriched = TruCheckSettingsSnapshot.Apply(
                record, applicationStandard, dataFormatCheck, apertureSetting);
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-DMCC] Push post-result settings: " +
                $"AppStd='{enriched.ApplicationStandardSetting ?? "unavailable"}' " +
                $"DFC='{enriched.DataFormatCheckSetting ?? "unavailable"}' " +
                $"Aperture='{enriched.ApertureSettingMode ?? "unavailable"}'");
            return enriched;
        }
        catch (Exception ex) when (ex is not OperationCanceledException || !ct.IsCancellationRequested)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-DMCC] Push post-result settings non-fatal: " +
                $"{ex.GetType().Name}: {ex.Message}");
            return record;
        }
        finally
        {
            _pushTruCheckSettingsGate.Release();
        }
    }

    private static async Task WriteRawDmccCommandAsync(
        System.Net.Sockets.NetworkStream stream,
        string command,
        System.Threading.CancellationToken ct)
    {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(
            $"{DmccCommand.WireHeader}{command}\r\n");
        await stream.WriteAsync(bytes, ct);
        await stream.FlushAsync(ct);
    }

    private static async Task<DmccResponse> SendRawDmccCommandAsync(
        System.Net.Sockets.NetworkStream stream,
        DeviceConfig cfg,
        string command,
        System.Threading.CancellationToken ct)
    {
        await WriteRawDmccCommandAsync(stream, command, ct);

        var response = new System.Text.StringBuilder();
        byte[] buffer = new byte[512];
        using var overallCts =
            System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct);
        overallCts.CancelAfter(cfg.ResponseTimeoutMs);

        while (true)
        {
            using var idleCts =
                System.Threading.CancellationTokenSource.CreateLinkedTokenSource(overallCts.Token);
            idleCts.CancelAfter(cfg.IdleGapMs);
            try
            {
                int count = await stream.ReadAsync(buffer, idleCts.Token);
                if (count == 0) break;
                response.Append(System.Text.Encoding.ASCII.GetString(buffer, 0, count));
            }
            catch (OperationCanceledException) when (!overallCts.IsCancellationRequested)
            {
                // The device has been idle long enough for the complete extended
                // response, including an optional GET body, to have arrived.
                break;
            }
        }

        string raw = response.ToString();
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-DMCC] Push {command}: " +
            $"'{raw.Replace("\r", "\\r").Replace("\n", "\\n")}'");
        return DmccResponse.Parse(raw);
    }

    // ── Push (DMST) mode callback ─────────────────────────────────────────────

    // Called on thread-pool by DmstListener after each parsed push result.
    // The push script (v1.5) sends the complete quality XML in one TCP push, so no
    // secondary DMCC GET SYMBOL.RESULT fetch is needed or possible on this firmware.
    private void OnPushRecord(VerificationRecord pushRecord)
    {
        System.Threading.Interlocked.Increment(ref _pendingAccept);
        _ = Task.Run(async () =>
        {
            try
            {
                VerificationRecord record = pushRecord;
                if (_scanMode == ScanMode.Push)
                    record = await AttachPushTruCheckSettingsAsync(
                        record, _pollCts?.Token ?? default);
                var watcher = _htmlWatcher;
                if (watcher is not null)
                {
                    var (merged, sourcePath) = await watcher.TryMergeAsync(
                        record, _pollCts?.Token ?? default);
                    record = sourcePath is not null
                        ? merged with { WebscanSourcePath = sourcePath }
                        : merged;
                }

                await Application.Current.Dispatcher.InvokeAsync(
                    () => AcceptRecordInnerAsync(record)).Task.Unwrap();
            }
            catch { /* non-fatal write error */ }
            finally
            {
                System.Threading.Interlocked.Decrement(ref _pendingAccept);
            }
        });
    }

    // ── Webscan USB HTML callback ────────────────────────────────────────────

    /// <summary>
    /// Called by the Webscan file adapter when a complete local HTML export
    /// arrives. This deliberately bypasses all DataMan HTTP/DMCC enrichment.
    /// </summary>
    private void OnWebscanRecord(int generation, VerificationRecord record)
    {
        _webscanAcceptance.TryAdmit(generation, () => IsRunning, async () =>
        {
            System.Threading.Interlocked.Increment(ref _pendingAccept);
            try
            {
                await Application.Current.Dispatcher.InvokeAsync(
                    () => _webscanAcceptance.IsCurrent(generation) && IsRunning
                        ? AcceptWebscanRecordAsync(record)
                        : Task.CompletedTask).Task.Unwrap();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-WEBSCAN] record acceptance failed: {ex.Message}");
            }
            finally
            {
                System.Threading.Interlocked.Decrement(ref _pendingAccept);
            }
        });
    }

    private async Task AcceptWebscanRecordAsync(VerificationRecord record)
    {
        WebscanImportError = null;
        await AcceptRecordInnerAsync(record);
    }

    private void OnWebscanParseFailed(int generation, string message)
    {
        if (generation != System.Threading.Volatile.Read(ref _webscanSessionGeneration) ||
            !IsRunning)
            return;

        _ = Application.Current.Dispatcher.InvokeAsync(() =>
        {
            if (generation == System.Threading.Volatile.Read(ref _webscanSessionGeneration) &&
                IsRunning)
            {
                WebscanImportError = $"Webscan import failed: {message}";
                StatusMessage = WebscanImportError;
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
    /// the provenance-correlation and write span).
    /// </summary>
    private async Task AcceptRecordInnerAsync(VerificationRecord record)
    {
        if (_sessionMgr is null) return;

        // Reset result display so the new scan starts clean.
        VerifierResultLine = string.Empty;
        RfidResultLine     = null;

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

        // Show the verifier result immediately — before the RFID window opens so
        // the operator sees the grade the moment OCR finishes.  Record number is an
        // estimate (_recordCount + 1); it is confirmed after the Excel write below.
        {
            string preGrade = record.OverallGrade?.LetterGradeString is { Length: > 0 } pg ? pg : "?";
            string preNum   = record.OverallGrade?.NumericGrade is { } pn ? $" ({pn:F1})" : string.Empty;
            string preOcr   = record.OcrResult?.Tier is { Length: > 0 } pt ? $"  |  OCR: {pt}" : string.Empty;
            VerifierResultLine = $"Record {_recordCount + 1}: {record.Symbology} — {preGrade}{preNum}{preOcr}";
        }

        // ── GS1 parser results ────────────────────────────────────────────────
        // RFID capture is independent of the barcode grade. Keep the VeriWedge
        // parser result available for the dual parser report block whenever the
        // decoded data is GS1-applicable. Whether the RFID row says Validation
        // or Cross-Validation is decided separately from the presence of this
        // comparison panel.
        TruCheckValidationAssessment truCheck =
            RfidValidator.AssessTruCheckValidation(record);
        string? gs1Input = record.HtmlDecodedData ?? record.DecodedData;
        DigitalLinkValidationResult? veriWedgeValidation =
            VccsDigitalLinkValidationService.Validate(gs1Input);
        if (veriWedgeValidation.Status == DigitalLinkValidationStatus.NotApplicable &&
            VccsDigitalLinkValidationService.BuildLinearElementString(
                record.HtmlSymbology ?? record.Symbology,
                gs1Input) is { } linearElementString)
        {
            veriWedgeValidation =
                VccsDigitalLinkValidationService.ValidateElementString(linearElementString);
        }
        if (veriWedgeValidation.Status == DigitalLinkValidationStatus.NotApplicable &&
            VccsDigitalLinkValidationService.LooksLikeGs1ElementString(gs1Input))
        {
            veriWedgeValidation =
                VccsDigitalLinkValidationService.ValidateElementString(gs1Input);
        }
        bool veriWedgeParserUsed = veriWedgeValidation.Status is
            DigitalLinkValidationStatus.Valid or
            DigitalLinkValidationStatus.Invalid or
            DigitalLinkValidationStatus.Unavailable;
        record = record with
        {
            TruCheckValidationUsable = truCheck.Usable,
            TruCheckValidationFailed = truCheck.Failed,
            VeriWedgeValidationUsed = veriWedgeParserUsed,
            VccsDigitalLinkValidation = veriWedgeValidation,
        };

        // A Webscan composite has two native barcode reports but one RFID
        // inventory window. Compare normalized GTINs without changing either
        // native decoded value.
        bool isComposite = !string.IsNullOrWhiteSpace(record.LinearSymbology);
        string? twoDGtin14 = isComposite
            ? RfidValidator.ExtractAi01(record.HtmlDecodedData ?? record.DecodedData)
            : null;
        string? linearGtin14 = isComposite
            ? RfidValidator.NormalizeLinearGtin14(
                record.LinearSymbology,
                record.LinearDecodedData)
            : null;
        string? barcodeAgreement = null;
        string? barcodeAgreementDetail = null;
        if (isComposite)
        {
            if (twoDGtin14 is null || linearGtin14 is null)
            {
                barcodeAgreement = "Incomplete";
                barcodeAgreementDetail =
                    $"2D GTIN-14: {twoDGtin14 ?? "missing"}; " +
                    $"linear GTIN-14: {linearGtin14 ?? "missing"}";
            }
            else if (twoDGtin14 == linearGtin14)
            {
                barcodeAgreement = "Pass";
                barcodeAgreementDetail = $"GTIN-14: {twoDGtin14}";
            }
            else
            {
                barcodeAgreement = "Fail";
                barcodeAgreementDetail =
                    $"2D GTIN-14: {twoDGtin14}; linear GTIN-14: {linearGtin14}";
            }
            record = record with
            {
                LinearGtin14 = linearGtin14,
                BarcodeSymbolAgreement = barcodeAgreement,
                BarcodeSymbolAgreementDetail = barcodeAgreementDetail,
            };
        }

        // ── RFID cross-validation (awaited before Excel write) ────────────────
        // Awaiting here means the RFID result lands in the same Excel row as the
        // barcode grade — no row-update pass needed later.
        // Non-fatal: an RFID error never blocks barcode record acceptance.
        // Reader participation is a hardware connection fact, not a coordinator
        // lifetime fact. The coordinator can be absent if startup failed or this
        // session has not attached one yet, while the launcher reader remains
        // connected and available.
        bool rfidReaderConnected = _rfidReader?.IsConnected == true;
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

        // Preserve scanner participation even when the scan window returns no
        // result. Report presentation must not infer it from RfidStatus alone.
        record = record with { RfidReaderConnected = rfidReaderConnected };

        if (rfidResult is not null)
        {
            // Resolve the GCP prefix table date from the loaded validator (available when the
            // GcpLengthTable was successfully loaded at session start).  Format as yyyy-MM-dd
            // for the PDF provenance annotation, e.g. "From GCP prefix table as of 2026-05-03".
            string? gcpTableDate = _gcpValidator?.DataDate?.ToString("yyyy-MM-dd");

            bool? rfidLinearMatches = isComposite && rfidResult.RfidGtin14 is not null &&
                linearGtin14 is not null
                    ? rfidResult.RfidGtin14 == linearGtin14
                    : null;
            string? rfidScope = isComposite && rfidResult.RfidGtin14 is not null
                ? (rfidResult.Gtin14Matches, rfidLinearMatches) switch
                {
                    (true, true) => "Both",
                    (true, false) => "2D only",
                    (false, true) => "Linear only",
                    _ => "Neither",
                }
                : null;
            string? rfidDetail = rfidResult.MismatchDetail;
            if (isComposite && rfidResult.RfidGtin14 is not null &&
                rfidLinearMatches == false)
            {
                string linearMismatch =
                    $"GTIN14:RFID={rfidResult.RfidGtin14},Linear={linearGtin14 ?? "missing"}";
                rfidDetail = string.IsNullOrWhiteSpace(rfidDetail)
                    ? linearMismatch
                    : rfidDetail + ";" + linearMismatch;
            }
            bool compositePass = !isComposite ||
                barcodeAgreement == "Pass" &&
                rfidScope == "Both" &&
                rfidResult.Status == RfidValidationStatus.Pass;
            record = record with
            {
                RfidStatus         = isComposite && !compositePass
                    && rfidResult.Status == RfidValidationStatus.Pass
                        ? RfidValidationStatus.Fail.ToString()
                        : rfidResult.Status.ToString(),
                RfidEpcHex         = rfidResult.SelectedRead?.EpcHex,
                RfidEpcTagUri      = BuildEpcTagUri(rfidResult.ParsedEpc),
                RfidGtin14         = rfidResult.RfidGtin14,
                RfidSerial         = rfidResult.RfidSerial,
                RfidTid            = rfidResult.SelectedRead?.Tid,
                RfidTagLockStatus  = rfidResult.SelectedRead?.LockStatus,
                RfidMismatchDetail = rfidDetail,
                RfidScanWindowMs   = rfidResult.ScanWindowMs,
                RfidGcpValid       = rfidResult.GcpValid,
                RfidGcpStatus      = rfidResult.GcpStatus.ToString(),
                RfidGcpLength      = GcpValidator.GetEncodedGcpLength(rfidResult.ParsedEpc),
                RfidGcpRegisteredLength = rfidResult.GcpRegisteredLength,
                RfidGcpTableDate   = gcpTableDate,
                RfidLinearGtin14Matches = rfidLinearMatches,
                RfidMatchScope = rfidScope,
                CompositeOverallStatus = isComposite
                    ? (compositePass ? "Pass" : "Fail")
                    : null,
            };

            // A two-symbol Webscan import has an independent barcode-to-barcode
            // comparison. Keep the RFID scan as one window, but make either
            // comparison part of the composite outcome.
            if (record.IsWebscanComposite && record.LinearTwoDMatch is false)
            {
                record = record with
                {
                    RfidStatus = "Fail",
                    RfidMismatchDetail =
                        string.IsNullOrWhiteSpace(rfidResult.MismatchDetail)
                            ? record.LinearTwoDComparisonDetail
                            : $"{record.LinearTwoDComparisonDetail}; {rfidResult.MismatchDetail}",
                };
            }

            // Update the RFID result line in the display.
            RfidResultLine = rfidResult.Status switch
            {
                RfidValidationStatus.Pass                 => $"RFID ✓  EPC: {rfidResult.SelectedRead?.EpcHex ?? "—"}",
                RfidValidationStatus.Fail                 => $"RFID ✗  Mismatch — {rfidResult.MismatchDetail ?? "see report"}",
                RfidValidationStatus.NoTag                => "RFID ·  No tag detected",
                RfidValidationStatus.ParseError           => "RFID ⚠  Parse error",
                RfidValidationStatus.MultipleTagsDetected => "RFID ⚠  Multiple tags detected",
                _                                         => null,
            };
        }
        else if (isComposite)
        {
            record = record with
            {
                CompositeOverallStatus = "Fail",
                RfidMatchScope = "Neither",
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

        // ── RFID tab row (auxiliary "RFID Scans" worksheet) ───────────────────
        // Written after AddRecord so LastSummaryRow points at the barcode row
        // just appended.  Non-fatal: a tab-write failure never drops the record.
        if (rfidResult is not null && _sessionMgr.Adapter is { } rfidAdapter)
        {
            try
            {
                _rfidTabWriter ??= new RfidTabWriter(rfidAdapter);
                _rfidTabWriter.EnsureSheet();
                _rfidTabWriter.AppendResult(rfidResult, _sessionMgr.LastSummaryRow ?? 0);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[RFID] Tab write failed: {ex.Message}");
            }
            finally
            {
                // Contract: restore the main sheet as the active write target.
                try { rfidAdapter.EnsureSheet(_sessionMgr.MainSheetName); } catch { }
            }

            // Persist the tab row now if the file is writable; a lock here is
            // fine — the row is in the in-memory workbook and lands on the next
            // successful save or at CloseSession.
            if (savedToDisk)
                try { rfidAdapter.Save(); } catch { /* locked — flushed later */ }
        }

        _history.AddRecord(record);
        _recordCount++; OnPropertyChanged(nameof(RecordCount)); OnPropertyChanged(nameof(WaitingMessage));

        // ── Hybrid HTML report (fire-and-forget) ──────────────────────────────
        // Generates a self-contained report combining barcode grades + RFID data.
        // Runs on the thread-pool; failures are silently swallowed so they never
        // interfere with the scan loop.
        //
        // Alongside mode (default):
        //   Report lands in the session output dir (or HybridReportOutputDirectory).
        //
        // Replace mode:
        //   The original Webscan HTML is deleted only after Replace mode was explicitly
        //   selected, then the hybrid report is written to the same source path.
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
                        await HybridReportGenerator.SaveToPathAsync(record, Path.Combine(dir, baseName + ".html"));

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

            // v23-faithful HTML → PDF pipeline (WebView2 primary, wkhtmltopdf
            // silent fallback).  Never throws — failures are Debug-logged only.
            _ = DeviceInterface.Reports.VccsPdfRenderer.GenerateReportAsync(
                record, pdfDir, _sessionId, _pollCts?.Token ?? default);
        }

        // Confirm the verifier result line with the actual record number and save status.
        // RfidResultLine was already set above when RFID ran.
        {
            string grade     = record.OverallGrade?.LetterGradeString is { Length: > 0 } g ? g : "?";
            string num       = record.OverallGrade?.NumericGrade is { } n ? $" ({n:F1})" : string.Empty;
            string ocrSuffix = record.OcrResult?.Tier is { Length: > 0 } t ? $"  |  OCR: {t}" : string.Empty;
            string pdfSuffix = _repo.Settings.GenerateVccsReport &&
                               !DeviceInterface.Reports.VccsHtmlReportGenerator.HasCorrelatedFilesystemHtml(record)
                ? "  |  PDF: NOT GENERATED — no correlated DMST HTML"
                : string.Empty;
            string prefix    = savedToDisk ? string.Empty : "⚠ ";
            VerifierResultLine = $"{prefix}Record {RecordCount}: {record.Symbology} — {grade}{num}{ocrSuffix}{pdfSuffix}";
        }
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
#if COGNEX_SDK
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
#else
            SupplementalStatus =
                "Cognex DataMan SDK DLL not present in this build — device settings unavailable.";
            await Task.CompletedTask;
#endif
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
#if COGNEX_SDK
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
#else
            SupplementalStatus =
                "Cognex DataMan SDK DLL not present in this build — device settings unavailable.";
            await Task.CompletedTask;
#endif
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
        if (IsRunning)
        {
            _pollCts?.Cancel();
            await CleanupAsync();
            IsRunning = false;
        }

        // Reader outlives sessions — release it explicitly on app exit.
        if (_rfidReader is not null)
        {
            try
            {
#if ASREADER_SDK
                // AsReader SDK 1.3.0 races an internal receive callback against
                // DisConnect during process exit. Use its shutdown-only path;
                // the ordinary Disconnect button still uses DisposeAsync.
                if (_rfidReader is AsReaderP35UEpcReader asReader)
                    await asReader.ShutdownForApplicationExitAsync();
                else
#endif
                    await _rfidReader.DisposeAsync();
            }
            catch { /* best effort */ }
            _rfidReader     = null;
            IsRfidConnected = false;
        }
    }

    private async Task CleanupAsync()
    {
        // Cleanup also serves application-exit and startup-failure paths, which
        // bypass the Stop button's normal drain.
        await QuiesceWebscanAsync();

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
            await _rfidCoordinator.DisposeAsync();   // ownsReader:false — reader stays connected
            _rfidCoordinator = null;
        }
        _rfidTabWriter = null;   // bound to the closed session's adapter
        _gcpValidator  = null;
        _sessionMgr?.Dispose();
        _sessionMgr = null;
        _pollCts?.Dispose();
        _pollCts = null;
    }

    /// <summary>
    /// Invalidates Webscan callbacks and waits for both file imports and record
    /// acceptance before any session-owned writer can be disposed. This is
    /// shared by Stop, application exit, and startup-failure cleanup.
    /// </summary>
    private async Task QuiesceWebscanAsync()
    {
        WebscanHtmlFileAdapter? adapter;
        EventHandler<VerificationRecord>? recordHandler;
        EventHandler<string>? parseFailedHandler;
        Task[] acceptanceTasks;

        lock (_webscanStateLock)
        {
            _webscanSessionGeneration++;
            adapter = _webscanHtmlAdapter;
            recordHandler = _webscanRecordHandler;
            parseFailedHandler = _webscanParseFailedHandler;
            _webscanHtmlAdapter = null;
            _webscanRecordHandler = null;
            _webscanParseFailedHandler = null;
        }
        acceptanceTasks = _webscanAcceptance.InvalidateAndCapture();

        if (adapter is not null)
        {
            if (recordHandler is not null)
                adapter.RecordParsed -= recordHandler;
            if (parseFailedHandler is not null)
                adapter.ParseFailed -= parseFailedHandler;
            await adapter.StopAsync();
            adapter.Dispose();
        }

        try
        {
            await Task.WhenAll(acceptanceTasks);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-WEBSCAN] acceptance drain failed: {ex.Message}");
        }
    }

}
