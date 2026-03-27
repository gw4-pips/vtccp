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
/// commands for Add / Edit / Delete / Set-Default / Test-Connection.
/// </summary>
public sealed class DevicesViewModel : ViewModelBase
{
    private readonly ConfigRepository _repo;

    private DeviceProfileViewModel? _selected;
    private DeviceProfileViewModel? _editing;
    private bool                    _isEditing;
    private bool                    _isTesting;
    private string                  _statusMessage = string.Empty;

    // ── Bindable collections and properties ───────────────────────────────────

    public ObservableCollection<DeviceProfileViewModel> Devices { get; } = [];

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

    /// <summary>True when the edit panel is open.</summary>
    public bool IsEditing
    {
        get => _isEditing;
        private set => Set(ref _isEditing, value);
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => Set(ref _statusMessage, value);
    }

    // ── Commands ──────────────────────────────────────────────────────────────

    public RelayCommand AddCommand          { get; }
    public RelayCommand EditCommand         { get; }
    public RelayCommand DeleteCommand       { get; }
    public RelayCommand DefaultCommand      { get; }
    public RelayCommand SaveCommand         { get; }
    public RelayCommand CancelCommand       { get; }
    public RelayCommand BrowseLogoCommand   { get; }
    public RelayCommand TestConnectCommand  { get; }

    public DevicesViewModel(ConfigRepository repo)
    {
        _repo = repo;

        AddCommand        = new RelayCommand(OnAdd);
        EditCommand       = new RelayCommand(OnEdit,        () => Selected is not null && !IsEditing);
        DeleteCommand     = new RelayCommand(OnDelete,      () => Selected is not null && !IsEditing);
        DefaultCommand    = new RelayCommand(OnSetDefault,  () => Selected is not null && !IsEditing);
        SaveCommand       = new RelayCommand(OnSave,        () => IsEditing);
        CancelCommand     = new RelayCommand(OnCancel,      () => IsEditing);
        BrowseLogoCommand = new RelayCommand(OnBrowseLogo,  () => IsEditing);
        TestConnectCommand = new RelayCommand(
            async () => await OnTestConnectAsync(),
            () => Selected is not null && !IsEditing && !_isTesting);

        Reload();
    }

    // ── Reload from repo ──────────────────────────────────────────────────────

    public void Reload()
    {
        Devices.Clear();
        foreach (var d in _repo.Devices)
            Devices.Add(new DeviceProfileViewModel(d));
        StatusMessage = string.Empty;
    }

    // ── Command handlers ──────────────────────────────────────────────────────

    private void OnAdd()
    {
        Editing   = new DeviceProfileViewModel();
        IsEditing = true;
    }

    private void OnEdit()
    {
        if (Selected is null) return;
        Editing   = new DeviceProfileViewModel(Selected.ToModel());
        IsEditing = true;
    }

    private void OnDelete()
    {
        if (Selected is null) return;
        if (MessageBox.Show($"Delete device profile '{Selected.Name}'?",
                "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
            != MessageBoxResult.Yes) return;

        _repo.RemoveDevice(Selected.Id);
        Reload();
        StatusMessage = "Device profile deleted.";
    }

    private void OnSetDefault()
    {
        if (Selected is null) return;
        foreach (var d in _repo.Devices) d.IsDefault = false;
        var target = _repo.FindDevice(Selected.Id);
        if (target is not null) target.IsDefault = true;
        Reload();
        StatusMessage = $"'{Selected.Name}' is now the default device.";
    }

    private void OnSave()
    {
        if (Editing is null) return;

        string name = Editing.Name.Trim();
        if (string.IsNullOrEmpty(name)) { StatusMessage = "Name is required."; return; }
        if (string.IsNullOrEmpty(Editing.Host.Trim())) { StatusMessage = "Host is required."; return; }
        if (Editing.Port is < 1 or > 65535) { StatusMessage = "Port must be 1–65535."; return; }

        Editing.Name = name;
        DeviceProfile model = Editing.ToModel();

        bool updated = _repo.UpdateDevice(model);
        if (!updated) _repo.AddDevice(model);

        Reload();
        IsEditing     = false;
        StatusMessage = updated ? "Device profile updated." : "Device profile added.";
    }

    private void OnCancel()
    {
        IsEditing = false;
        Editing   = null;
    }

    private void OnBrowseLogo()
    {
        if (Editing is null) return;
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title  = "Select Logo Image",
            Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All files|*.*",
        };
        if (dlg.ShowDialog() == true) Editing.Notes = Editing.Notes;
    }

    /// <summary>
    /// Connects to the selected device, queries its type and firmware, and
    /// reports the result (including round-trip latency) in the status bar.
    /// Uses a 3-second timeout so the UI stays responsive.
    /// </summary>
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
}
