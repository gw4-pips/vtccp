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

    private DeviceSession?                      _deviceSession;
    private SessionManager?                     _sessionMgr;
    private System.Threading.CancellationTokenSource? _pollCts;
    private bool                                _isRunning;
    private string                              _statusMessage = "Ready.";
    private int                                 _recordCount;

    // ── Selection ─────────────────────────────────────────────────────────────

    private DeviceProfile? _selectedDevice;
    private JobTemplate?   _selectedTemplate;
    private string         _operatorOverride = string.Empty;
    private ScanMode       _scanMode         = ScanMode.Manual;
    private int            _autoPollIntervalMs = 500;

    // ── Bindable collections ──────────────────────────────────────────────────

    public ObservableCollection<DeviceProfile> AvailableDevices   { get; } = [];
    public ObservableCollection<JobTemplate>   AvailableTemplates { get; } = [];

    // ── Bindable properties ───────────────────────────────────────────────────

    public DeviceProfile? SelectedDevice
    {
        get => _selectedDevice;
        set { Set(ref _selectedDevice, value); OnPropertyChanged(nameof(IsPushAvailable)); RelayCommand.Refresh(); }
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
            RelayCommand.Refresh();
        }
    }

    public bool IsManualMode   => _scanMode == ScanMode.Manual;
    public bool IsAutoPollMode => _scanMode == ScanMode.AutoPoll;
    public bool IsPushMode     => _scanMode == ScanMode.Push;

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

    public int RecordCount
    {
        get => _recordCount;
        private set => Set(ref _recordCount, value);
    }

    // ── Commands ──────────────────────────────────────────────────────────────

    public RelayCommand StartCommand        { get; }
    public RelayCommand StopCommand         { get; }
    public RelayCommand TriggerCommand      { get; }
    public RelayCommand SetManualCommand    { get; }
    public RelayCommand SetAutoPollCommand  { get; }
    public RelayCommand SetPushCommand      { get; }

    public SessionViewModel(ConfigRepository repo, HistoryViewModel history)
    {
        _repo    = repo;
        _history = history;

        StartCommand   = new RelayCommand(async () => await OnStartAsync(),
            () => !IsRunning && SelectedDevice is not null && SelectedTemplate is not null);
        StopCommand    = new RelayCommand(async () => await OnStopAsync(),
            () => IsRunning);
        TriggerCommand = new RelayCommand(async () => await OnTriggerAsync(),
            () => IsRunning && _scanMode == ScanMode.Manual);

        SetManualCommand   = new RelayCommand(() => ActiveScanMode = ScanMode.Manual,   () => !IsRunning);
        SetAutoPollCommand = new RelayCommand(() => ActiveScanMode = ScanMode.AutoPoll, () => !IsRunning);
        SetPushCommand     = new RelayCommand(() => ActiveScanMode = ScanMode.Push,     () => !IsRunning && IsPushAvailable);

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
    }

    // ── Start / Stop ──────────────────────────────────────────────────────────

    private async Task OnStartAsync()
    {
        if (SelectedDevice is null || SelectedTemplate is null) return;

        StatusMessage = "Connecting to device…";
        try
        {
            var cfg = SelectedDevice.ToDeviceConfig();
            _deviceSession = new DeviceSession(cfg, _xmlMap);
            await _deviceSession.ConnectAsync();

            string outputDir = !string.IsNullOrWhiteSpace(SelectedTemplate.OutputDirectory)
                ? SelectedTemplate.OutputDirectory
                : _repo.Settings.DefaultOutputDirectory;

            SessionState state = SelectedTemplate.ToSessionState(outputDir);
            if (!string.IsNullOrWhiteSpace(OperatorOverride))
                state.OperatorId = OperatorOverride.Trim();

            _sessionMgr = new SessionManager(TruCheckCompatibleSchema.Build());
            await Task.Run(() => _sessionMgr.StartSession(state));

            _history.SetSessionContext(state.JobName, state.OperatorId);
            _history.ClearHistory();
            RecordCount   = 0;
            IsRunning     = true;

            string modeLabel = _scanMode switch
            {
                ScanMode.AutoPoll => $"Auto-Poll ({_autoPollIntervalMs} ms)",
                ScanMode.Push     => "Push (DMST)",
                _                 => "Manual Trigger",
            };
            StatusMessage = $"Session active — {SelectedDevice.Name} / {SelectedTemplate.Name}  [{modeLabel}]";

            // Launch background scan mode if needed.
            _pollCts = new System.Threading.CancellationTokenSource();
            if (_scanMode == ScanMode.AutoPoll)
                _ = RunAutoPollLoopAsync(_pollCts.Token);
            else if (_scanMode == ScanMode.Push)
                await StartPushModeAsync(state, _pollCts.Token);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Connect failed: {ex.Message}";
            await CleanupAsync();
        }
    }

    private async Task OnStopAsync()
    {
        StatusMessage = "Closing session…";
        _pollCts?.Cancel();   // stops AutoPoll loop and Push listener
        try
        {
            if (_sessionMgr is not null)
                await Task.Run(() => _sessionMgr.CloseSession());
        }
        finally
        {
            await CleanupAsync();
            IsRunning     = false;
            StatusMessage = $"Session closed. {RecordCount} record(s) written.";
        }
    }

    // ── Manual trigger ────────────────────────────────────────────────────────

    private async Task OnTriggerAsync()
    {
        if (_deviceSession is null || _sessionMgr is null) return;
        try
        {
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

    // ── Push (DMST) mode ──────────────────────────────────────────────────────

    private async Task StartPushModeAsync(SessionState state, System.Threading.CancellationToken ct)
    {
        if (_deviceSession is null) return;

        var ctx = BuildContext();
        ctx = ctx with
        {
            OperatorId  = state.OperatorId,
            JobName     = state.JobName,
            BatchNumber = state.BatchNumber,
            CompanyName = state.CompanyName,
        };

        _deviceSession.ResultReceived += OnPushResultReceived;
        await _deviceSession.StartPushListenerAsync(ctx, ct);
    }

    private void OnPushResultReceived(object? sender, VerificationRecord record)
    {
        Application.Current.Dispatcher.InvokeAsync(async () =>
        {
            try   { await AcceptRecordAsync(record); }
            catch { /* non-fatal push delivery error */ }
        });
    }

    // ── Shared record acceptance ──────────────────────────────────────────────

    private async Task AcceptRecordAsync(VerificationRecord record)
    {
        if (_sessionMgr is null) return;
        await Task.Run(() => _sessionMgr.AddRecord(record));
        _history.AddRecord(record);
        RecordCount++;
        string grade = record.OverallGrade?.LetterGradeString is { Length: > 0 } g ? g : "?";
        string num   = record.OverallGrade?.NumericGrade is { } n ? $" ({n:F1})" : string.Empty;
        StatusMessage = $"Record {RecordCount}: {record.Symbology} — {grade}{num}";
    }

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

    private async Task CleanupAsync()
    {
        if (_deviceSession is not null)
        {
            _deviceSession.ResultReceived -= OnPushResultReceived;
            await _deviceSession.DisposeAsync();
            _deviceSession = null;
        }
        _sessionMgr?.Dispose();
        _sessionMgr = null;
        _pollCts?.Dispose();
        _pollCts = null;
    }
}
