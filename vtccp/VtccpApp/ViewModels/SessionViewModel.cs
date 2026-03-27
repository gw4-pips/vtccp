namespace VtccpApp.ViewModels;

using System.Collections.ObjectModel;
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
/// Lets the operator pick a device profile + job template, enter an
/// optional operator override, and start/stop a live scan session.
/// Each result is also forwarded to <see cref="HistoryViewModel"/> for the
/// Results History page.
/// </summary>
public sealed class SessionViewModel : ViewModelBase
{
    private readonly ConfigRepository _repo;
    private readonly HistoryViewModel _history;
    private readonly VerificationXmlMap _xmlMap = new();

    // ── Runtime session state ─────────────────────────────────────────────────

    private DeviceSession?  _deviceSession;
    private SessionManager? _sessionMgr;
    private bool            _isRunning;
    private string          _statusMessage = "Ready.";
    private int             _recordCount;

    // ── Selection ─────────────────────────────────────────────────────────────

    private DeviceProfile? _selectedDevice;
    private JobTemplate?   _selectedTemplate;
    private string         _operatorOverride = string.Empty;

    // ── Bindable properties ───────────────────────────────────────────────────

    public ObservableCollection<DeviceProfile> AvailableDevices   { get; } = [];
    public ObservableCollection<JobTemplate>   AvailableTemplates { get; } = [];

    public DeviceProfile? SelectedDevice
    {
        get => _selectedDevice;
        set { Set(ref _selectedDevice, value); RelayCommand.Refresh(); }
    }

    public JobTemplate? SelectedTemplate
    {
        get => _selectedTemplate;
        set { Set(ref _selectedTemplate, value); RelayCommand.Refresh(); }
    }

    /// <summary>Operator ID entered at session start; overrides the template's value if non-empty.</summary>
    public string OperatorOverride
    {
        get => _operatorOverride;
        set => Set(ref _operatorOverride, value);
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

    public RelayCommand StartCommand   { get; }
    public RelayCommand StopCommand    { get; }
    public RelayCommand TriggerCommand { get; }

    public SessionViewModel(ConfigRepository repo, HistoryViewModel history)
    {
        _repo    = repo;
        _history = history;

        StartCommand   = new RelayCommand(async () => await OnStartAsync(),
            () => !IsRunning && SelectedDevice is not null && SelectedTemplate is not null);
        StopCommand    = new RelayCommand(async () => await OnStopAsync(),
            () => IsRunning);
        TriggerCommand = new RelayCommand(async () => await OnTriggerAsync(),
            () => IsRunning);

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

    // ── Command handlers ──────────────────────────────────────────────────────

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
                state = state with { OperatorId = OperatorOverride.Trim() };

            _sessionMgr = new SessionManager(TruCheckCompatibleSchema.Build());
            await Task.Run(() => _sessionMgr.StartSession(state));

            _history.ClearHistory();
            RecordCount   = 0;
            IsRunning     = true;
            StatusMessage = $"Session active — {SelectedDevice.Name} / {SelectedTemplate.Name}";
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

    private async Task OnTriggerAsync()
    {
        if (_deviceSession is null || _sessionMgr is null) return;
        try
        {
            var ctx = new VerificationRecord
            {
                Symbology       = string.Empty,
                DeviceSerial    = _deviceSession.DeviceInfo.Serial    ?? string.Empty,
                DeviceName      = _deviceSession.DeviceInfo.Name      ?? string.Empty,
                FirmwareVersion = _deviceSession.DeviceInfo.FirmwareVersion ?? string.Empty,
                OperatorId      = string.IsNullOrWhiteSpace(OperatorOverride)
                                      ? (SelectedTemplate?.OperatorId ?? string.Empty)
                                      : OperatorOverride.Trim(),
                JobName         = SelectedTemplate?.JobName ?? string.Empty,
            };

            VerificationRecord? record = await _deviceSession.TriggerAndGetResultAsync(ctx);

            if (record is not null)
            {
                await Task.Run(() => _sessionMgr.AddRecord(record));
                _history.AddRecord(record);
                RecordCount++;
                string grade = record.OverallGrade?.LetterGradeString is { Length: > 0 } g ? g : "?";
                StatusMessage = $"Record {RecordCount}: {record.Symbology} — {grade}";
            }
            else
            {
                StatusMessage = "No read.";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Trigger error: {ex.Message}";
        }
    }

    private async Task CleanupAsync()
    {
        if (_deviceSession is not null)
        {
            await _deviceSession.DisposeAsync();
            _deviceSession = null;
        }
        _sessionMgr?.Dispose();
        _sessionMgr = null;
    }
}
