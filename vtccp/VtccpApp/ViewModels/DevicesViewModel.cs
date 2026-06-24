namespace VtccpApp.ViewModels;

using System.Collections.ObjectModel;
using System.Windows;
using ConfigEngine;
using ConfigEngine.Models;
using DeviceInterface;
using VtccpApp.Commands;

/// <summary>
/// Manages the device profile list for the Devices page.
/// Exposes an <see cref="ObservableCollection{T}"/> for the list view and
/// commands for Add / Edit / Duplicate / Delete / Set-Default / Test-Connection /
/// Scan-Network / Import-Discovered / Dismiss-Discovery.
/// </summary>
public sealed class DevicesViewModel : ViewModelBase
{
    private readonly ConfigRepository _repo;

    private DeviceProfileViewModel? _selected;
    private DeviceProfileViewModel? _editing;
    private bool                    _isEditing;
    private bool                    _isNew;
    private bool                    _isTesting;
    private bool                    _isScanning;
    private string                  _statusMessage = string.Empty;

    // ── Bindable collections and properties ───────────────────────────────────

    public ObservableCollection<DeviceProfileViewModel>    Devices          { get; } = [];
    public ObservableCollection<DiscoveredDeviceViewModel> DiscoveredDevices { get; } = [];

    public DeviceProfileViewModel? Selected
    {
        get => _selected;
        set { Set(ref _selected, value); RelayCommand.Refresh(); }
    }

    /// <summary>The profile currently open in the edit form (clone of Selected or new).</summary>
    public DeviceProfileViewModel? Editing
    {
        get => _editing;
        private set => Set(ref _editing, value);
    }

    /// <summary>True when the edit/add panel is open.</summary>
    public bool IsEditing
    {
        get => _isEditing;
        private set
        {
            if (Set(ref _isEditing, value))
                RelayCommand.Refresh();
        }
    }

    /// <summary>True while a network scan is running.</summary>
    public bool IsScanning
    {
        get => _isScanning;
        private set
        {
            Set(ref _isScanning, value);
            RelayCommand.Refresh();
        }
    }

    /// <summary>"Add Device Profile" or "Edit Device Profile" depending on context.</summary>
    public string EditingTitle => _isNew ? "Add Device Profile" : "Edit Device Profile";

    /// <summary>True when the discovery results panel should be visible.</summary>
    public bool HasDiscoveredDevices => DiscoveredDevices.Count > 0;

    public string StatusMessage
    {
        get => _statusMessage;
        private set => Set(ref _statusMessage, value);
    }

    // ── Commands ──────────────────────────────────────────────────────────────

    public RelayCommand AddCommand             { get; }
    public RelayCommand EditCommand            { get; }
    public RelayCommand DuplicateCommand       { get; }
    public RelayCommand DeleteCommand          { get; }
    public RelayCommand DefaultCommand         { get; }
    public RelayCommand SaveCommand            { get; }
    public RelayCommand CancelCommand          { get; }
    public RelayCommand TestConnectCommand     { get; }
    public RelayCommand ScanNetworkCommand     { get; }
    public RelayCommand DismissDiscoveryCommand { get; }

    public RelayCommand<DiscoveredDeviceViewModel> ImportDiscoveredCommand { get; }

    private readonly Action? _onListChanged;

    public DevicesViewModel(ConfigRepository repo, Action? onListChanged = null)
    {
        _repo          = repo;
        _onListChanged = onListChanged;

        AddCommand       = new RelayCommand(OnAdd);
        EditCommand      = new RelayCommand(OnEdit,      () => Selected is not null && !IsEditing);
        DuplicateCommand = new RelayCommand(OnDuplicate, () => Selected is not null && !IsEditing);
        DeleteCommand    = new RelayCommand(OnDelete,    () => Selected is not null && !IsEditing);
        DefaultCommand   = new RelayCommand(OnSetDefault, () => Selected is not null && !IsEditing);
        SaveCommand      = new RelayCommand(OnSave,       () => IsEditing);
        CancelCommand    = new RelayCommand(OnCancel,     () => IsEditing);

        TestConnectCommand = new RelayCommand(
            async () => await OnTestConnectAsync(),
            () => Selected is not null && !IsEditing && !_isTesting);

        ScanNetworkCommand = new RelayCommand(
            async () => await OnScanNetworkAsync(),
            () => !IsScanning && !IsEditing);

        DismissDiscoveryCommand = new RelayCommand(OnDismissDiscovery);

        ImportDiscoveredCommand = new RelayCommand<DiscoveredDeviceViewModel>(OnImportDiscovered);

        Reload();
    }

    // ── Reload from repo ──────────────────────────────────────────────────────

    public void Reload()
    {
        Devices.Clear();
        foreach (var d in _repo.Devices)
            Devices.Add(new DeviceProfileViewModel(d));
        StatusMessage = Devices.Count == 0
            ? "No device profiles yet. Click + Add or 🔍 Scan to get started."
            : "Select a device to edit or test, or click + Add to create a new profile.";
    }

    // ── Command handlers — list ───────────────────────────────────────────────

    private void OnAdd()
    {
        _isNew    = true;
        Editing   = new DeviceProfileViewModel();
        IsEditing = true;
        OnPropertyChanged(nameof(EditingTitle));
        StatusMessage = string.Empty;
    }

    private void OnEdit()
    {
        if (Selected is null)
        {
            StatusMessage = "Select a device in the list first, then click Edit.";
            return;
        }
        _isNew    = false;
        Editing   = new DeviceProfileViewModel(Selected.ToModel());
        IsEditing = true;
        OnPropertyChanged(nameof(EditingTitle));
        StatusMessage = string.Empty;
    }

    private void OnDuplicate()
    {
        if (Selected is null)
        {
            StatusMessage = "Select a device in the list first, then click Duplicate.";
            return;
        }
        var clone = new DeviceProfileViewModel(Selected.ToModel())
        {
            Id   = Guid.NewGuid().ToString(),
            Name = Selected.Name + " (copy)",
        };
        _isNew    = true;
        Editing   = clone;
        IsEditing = true;
        OnPropertyChanged(nameof(EditingTitle));
        StatusMessage = "Editing duplicate — update the name and host, then click Save.";
    }

    private void OnDelete()
    {
        if (Selected is null)
        {
            StatusMessage = "Select a device in the list first, then click Delete.";
            return;
        }
        if (MessageBox.Show($"Delete device profile '{Selected.Name}'?",
                "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
            != MessageBoxResult.Yes) return;

        _repo.RemoveDevice(Selected.Id);
        _ = _repo.SaveAsync();
        Reload();
        _onListChanged?.Invoke();
        StatusMessage = "Device profile deleted.";
    }

    private void OnSetDefault()
    {
        if (Selected is null) return;
        foreach (var d in _repo.Devices) d.IsDefault = false;
        var target = _repo.FindDevice(Selected.Id);
        if (target is not null) target.IsDefault = true;
        _ = _repo.SaveAsync();
        Reload();
        _onListChanged?.Invoke();
        StatusMessage = $"'{Selected.Name}' is now the default device.";
    }

    // ── Command handlers — edit panel ─────────────────────────────────────────

    private void OnSave()
    {
        if (Editing is null) return;

        string name = Editing.Name.Trim();
        if (string.IsNullOrEmpty(name))              { StatusMessage = "Name is required.";       return; }
        if (string.IsNullOrEmpty(Editing.Host.Trim())) { StatusMessage = "Host is required.";     return; }
        if (Editing.Port is < 1 or > 65535)          { StatusMessage = "Port must be 1–65535.";   return; }

        Editing.Name = name;
        DeviceProfile model = Editing.ToModel();

        bool updated = _repo.UpdateDevice(model);
        if (!updated) _repo.AddDevice(model);
        _ = _repo.SaveAsync();

        Reload();
        _onListChanged?.Invoke();
        IsEditing     = false;
        StatusMessage = updated ? "Device profile updated." : "Device profile added.";
    }

    private void OnCancel()
    {
        IsEditing = false;
        Editing   = null;
        StatusMessage = string.Empty;
    }

    // ── Command handlers — Test Connection ───────────────────────────────────

    private async Task OnTestConnectAsync()
    {
        if (Selected is null) return;

        _isTesting = true;
        RelayCommand.Refresh();
        StatusMessage = $"Testing connection to {Selected.Host}:{Selected.Port}…";

        var sw = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(3));
            var cfg     = Selected.ToModel().ToDeviceConfig();
            var session = new DeviceSession(cfg);
            await using (session)
            {
                await session.ConnectAsync(cts.Token);
                sw.Stop();
                var info = session.DeviceInfo;
                StatusMessage =
                    $"Connected  ·  {sw.ElapsedMilliseconds} ms  |  " +
                    $"{info.Type ?? "?"}  ·  FW {info.FirmwareVersion ?? "?"}  ·  S/N {info.Serial ?? "?"}";
            }
        }
        catch (OperationCanceledException)
        {
            StatusMessage = $"Connection timed out after 3 s  ({Selected.Host}:{Selected.Port})";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Connection failed: {ex.Message}";
        }
        finally
        {
            _isTesting = false;
            RelayCommand.Refresh();
        }
    }

    // ── Command handlers — Network Discovery ─────────────────────────────────

    private async Task OnScanNetworkAsync()
    {
        IsScanning = true;
        DiscoveredDevices.Clear();
        OnPropertyChanged(nameof(HasDiscoveredDevices));
        StatusMessage = "Scanning network for DataMan devices (3 s)…";

        try
        {
            var found = await NetworkDiscoverer.DiscoverAsync(listenMs: 3_000);

            foreach (var d in found)
                DiscoveredDevices.Add(new DiscoveredDeviceViewModel(d));

            OnPropertyChanged(nameof(HasDiscoveredDevices));

            StatusMessage = DiscoveredDevices.Count > 0
                ? $"Found {DiscoveredDevices.Count} device(s). Click ⊕ Import to add a profile."
                : "No DataMan devices found on this subnet.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Network scan failed: {ex.Message}";
        }
        finally
        {
            IsScanning = false;
        }
    }

    private void OnImportDiscovered(DiscoveredDeviceViewModel? device)
    {
        if (device is null) return;
        _isNew  = true;
        Editing = new DeviceProfileViewModel
        {
            Name            = device.Name,
            Host            = device.Host,
            Port            = device.Port,
            DeviceType      = device.DeviceType,
        };
        IsEditing = true;
        OnPropertyChanged(nameof(EditingTitle));
        StatusMessage = $"Importing {device.Name} — review the fields and click Save.";
    }

    private void OnDismissDiscovery()
    {
        DiscoveredDevices.Clear();
        OnPropertyChanged(nameof(HasDiscoveredDevices));
    }
}
