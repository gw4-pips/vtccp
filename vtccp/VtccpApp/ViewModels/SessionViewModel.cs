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
    private DeviceInterface.Dmst.DmstListener?  _dmstListener;
    private SessionManager?                     _sessionMgr;
    private System.Threading.CancellationTokenSource? _pollCts;
    private bool                                _isRunning;
    private string                              _statusMessage = "Ready.";
    private int                                 _recordCount;

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

    public DeviceProfile? SelectedDevice
    {
        get => _selectedDevice;
        set
        {
            Set(ref _selectedDevice, value);
            OnPropertyChanged(nameof(IsPushAvailable));
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
            () => IsRunning && (_scanMode == ScanMode.Manual || _scanMode == ScanMode.Push));

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

        string outputDir = !string.IsNullOrWhiteSpace(SelectedTemplate.OutputDirectory)
            ? SelectedTemplate.OutputDirectory
            : _repo.Settings.DefaultOutputDirectory;

        SessionState state = SelectedTemplate.ToSessionState(outputDir);
        if (!string.IsNullOrWhiteSpace(OperatorOverride))
            state.OperatorId = OperatorOverride.Trim();

        _sessionMgr = new SessionManager(TruCheckCompatibleSchema.Build());
        _pollCts    = new System.Threading.CancellationTokenSource();

        try
        {
            if (_scanMode == ScanMode.Push)
            {
                // Push mode: DMST stays open on port 23 — we only listen for pushed XML.
                // No DMCC connection is opened so DataMan Setup Tool can remain active
                // for live view and positioning.
                int listenPort = SelectedDevice.DmstListenPort;
                StatusMessage = $"Starting push listener on port {listenPort}…";

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

                _dmstListener = new DeviceInterface.Dmst.DmstListener(
                    listenPort, _xmlMap, ctx, OnPushRecord);
                await _dmstListener.StartAsync(_pollCts.Token);
            }
            else
            {
                // Manual / AutoPoll: open DMCC connection (requires DMST to be closed).
                StatusMessage = "Connecting to device…";
                var cfg = SelectedDevice.ToDeviceConfig();
                _deviceSession = new DeviceSession(cfg, _xmlMap);
                await _deviceSession.ConnectAsync();
            }

            await Task.Run(() => _sessionMgr.StartSession(state));
            _history.SetSessionContext(state.JobName, state.OperatorId);
            _history.ClearHistory();
            _recordCount = 0; OnPropertyChanged(nameof(RecordCount));
            IsRunning    = true;

            string modeLabel = _scanMode switch
            {
                ScanMode.AutoPoll => $"Auto-Poll ({_autoPollIntervalMs} ms)",
                ScanMode.Push     => $"Push (DMST) — port {SelectedDevice.DmstListenPort}",
                _                 => "Manual Trigger",
            };
            StatusMessage = $"Session active — {SelectedDevice.Name} / {SelectedTemplate.Name}  [{modeLabel}]";

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
        if (_dmstListener is not null)
        {
            await _dmstListener.StopAsync();
            _dmstListener = null;
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

        // Short-timeout config — keep IdleGapMs from the profile so end-of-response
        // detection works correctly with this device's firmware.
        var cfg = SelectedDevice.ToDeviceConfig();
        cfg.ConnectTimeoutMs  = 3_000;
        cfg.ResponseTimeoutMs = 3_000;

        await using var client = new DeviceInterface.Dmcc.DmccClient(cfg);
        await client.ConnectAsync();

        var resp = await client.SendAsync(DeviceInterface.Dmcc.DmccCommand.Trigger);

        StatusMessage = resp.StatusCode switch
        {
            DeviceInterface.Dmcc.DmccStatus.Ok =>
                "Trigger sent — waiting for push result…",

            DeviceInterface.Dmcc.DmccStatus.NoRead =>
                "Trigger fired — no symbol in field of view.",

            DeviceInterface.Dmcc.DmccStatus.Busy =>
                "Device busy — trigger rejected. Wait a moment and retry.",

            _ => string.IsNullOrWhiteSpace(resp.Body)
                    ? $"Trigger: device returned code {resp.StatusCode}."
                    : $"Trigger: device returned code {resp.StatusCode} — {resp.Body}",
        };
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
    private void OnPushRecord(VerificationRecord record)
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
        System.Threading.Interlocked.Increment(ref _pendingAccept);
        try
        {
            await Task.Run(() => _sessionMgr.AddRecord(record));
            _history.AddRecord(record);
            _recordCount++; OnPropertyChanged(nameof(RecordCount));
            string grade = record.OverallGrade?.LetterGradeString is { Length: > 0 } g ? g : "?";
            string num   = record.OverallGrade?.NumericGrade is { } n ? $" ({n:F1})" : string.Empty;
            StatusMessage = $"Record {RecordCount}: {record.Symbology} — {grade}{num}";
        }
        finally
        {
            System.Threading.Interlocked.Decrement(ref _pendingAccept);
        }
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
            await _deviceSession.DisposeAsync();
            _deviceSession = null;
        }
        if (_dmstListener is not null)
        {
            await _dmstListener.StopAsync();
            _dmstListener = null;
        }
        _sessionMgr?.Dispose();
        _sessionMgr = null;
        _pollCts?.Dispose();
        _pollCts = null;
    }
}
