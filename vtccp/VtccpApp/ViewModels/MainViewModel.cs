namespace VtccpApp.ViewModels;

using ConfigEngine;
using VtccpApp.Commands;
using VtccpApp.Services;

/// <summary>
/// Root view-model for the application shell (MainWindow).
/// Owns the navigation state and top-level <see cref="ConfigRepository"/>.
/// </summary>
public sealed class MainViewModel : ViewModelBase
{
    // ── Repository (shared across child view-models) ───────────────────────────
    public ConfigRepository Repository { get; } = new();

    // ── Child view-models ─────────────────────────────────────────────────────
    public DevicesViewModel   DevicesVM   { get; }
    public TemplatesViewModel TemplatesVM { get; }
    public SessionViewModel   SessionVM   { get; }
    public HistoryViewModel   HistoryVM   { get; }
    public SettingsViewModel  SettingsVM  { get; }

    // ── Navigation ────────────────────────────────────────────────────────────

    private ViewModelBase? _currentPage;
    private string         _currentPageKey = string.Empty;

    public ViewModelBase? CurrentPage
    {
        get => _currentPage;
        private set => Set(ref _currentPage, value);
    }

    public string CurrentPageKey
    {
        get => _currentPageKey;
        private set => Set(ref _currentPageKey, value);
    }

    public RelayCommand NavDevicesCommand   { get; }
    public RelayCommand NavTemplatesCommand { get; }
    public RelayCommand NavSessionCommand   { get; }
    public RelayCommand NavHistoryCommand   { get; }
    public RelayCommand NavSettingsCommand  { get; }

    // ── Title bar ─────────────────────────────────────────────────────────────

    public string AppVersion { get; } = AppVersionDisplay.Current;
    public string AppTitle => $"VTCCP — VCCS DMV TruCheck Command Pilot {AppVersion}";
    public string SidebarVersion => $"{AppVersion}  |  VCCS";

    // ── Init ──────────────────────────────────────────────────────────────────

    public MainViewModel()
    {
        HistoryVM   = new HistoryViewModel();
        SessionVM   = new SessionViewModel(Repository, HistoryVM);
        DevicesVM   = new DevicesViewModel(Repository,  onListChanged: SessionVM.Reload);
        TemplatesVM = new TemplatesViewModel(Repository, onListChanged: SessionVM.Reload);
        SettingsVM  = new SettingsViewModel(Repository);

        NavDevicesCommand   = new RelayCommand(() => Navigate("Devices"));
        NavTemplatesCommand = new RelayCommand(() => Navigate("Templates"));
        NavSessionCommand   = new RelayCommand(() => Navigate("Session"));
        NavHistoryCommand   = new RelayCommand(() => Navigate("History"));
        NavSettingsCommand  = new RelayCommand(() => Navigate("Settings"));

        Navigate("Session");   // default page

        InstallGcpUpdateCommand = new RelayCommand(() => _ = InstallGcpUpdateAsync(), () => !_gcpInstalling);
        DismissGcpToastCommand  = new RelayCommand(() => IsGcpToastVisible = false);

        _ = InitAsync();
    }

    private async Task InitAsync()
    {
        await LoadConfigAsync();
        _ = CheckGcpUpdateAsync();   // background, non-blocking, best-effort
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    private void Navigate(string key)
    {
        CurrentPageKey = key;
        CurrentPage = key switch
        {
            "Devices"   => DevicesVM,
            "Templates" => TemplatesVM,
            "Session"   => SessionVM,
            "History"   => HistoryVM,
            "Settings"  => SettingsVM,
            _           => SessionVM,
        };
    }

    // ── Config persistence ────────────────────────────────────────────────────

    public async Task LoadConfigAsync()
    {
        try
        {
            await Repository.LoadAsync();
            DevicesVM.Reload();
            TemplatesVM.Reload();
            SessionVM.Reload();
            SettingsVM.Reload();
        }
        catch { /* first run — defaults in effect */ }
    }

    public async Task SaveConfigAsync()
    {
        try   { await Repository.SaveAsync(); }
        catch { /* non-fatal */ }
    }

    // ── GCP prefix table update toast ─────────────────────────────────────────
    // Startup check against the Azure update service (Settings → Data Sources).
    // Silent no-op when the service URL / device token are not configured or
    // the workstation is offline.

    private bool   _isGcpToastVisible;
    private string _gcpToastMessage = string.Empty;
    private bool   _gcpInstalling;

    public RelayCommand InstallGcpUpdateCommand { get; }
    public RelayCommand DismissGcpToastCommand  { get; }

    public bool IsGcpToastVisible
    {
        get => _isGcpToastVisible;
        private set => Set(ref _isGcpToastVisible, value);
    }

    public string GcpToastMessage
    {
        get => _gcpToastMessage;
        private set => Set(ref _gcpToastMessage, value);
    }

    private async Task CheckGcpUpdateAsync()
    {
        try
        {
            var service = Services.GcpUpdateServiceFactory.Create(Repository.Settings);
            if (service is null) return;

            var check = await service.CheckNowAsync();
            if (check is not { UpdateAvailable: true }) return;

            GcpToastMessage =
                $"GCP prefix table update available ({check.ServerDate:yyyy-MM-dd}). Install now?";
            IsGcpToastVisible = true;
        }
        catch { /* best-effort — never disturb startup */ }
    }

    private async Task InstallGcpUpdateAsync()
    {
        if (_gcpInstalling) return;
        _gcpInstalling = true;
        RelayCommand.Refresh();
        try
        {
            var service = Services.GcpUpdateServiceFactory.Create(Repository.Settings);
            if (service is null) return;

            GcpToastMessage = "Downloading GCP prefix table…";
            var installedDate = await service.DownloadAndInstallAsync();

            Repository.Settings.GcpDataPath =
                Services.GcpUpdateServiceFactory.ResolveLocalXmlPath(Repository.Settings);
            Repository.Settings.GcpLastModified = installedDate?.ToString("O");
            await Repository.SaveSettingsAsync();
            SettingsVM.Reload();

            GcpToastMessage =
                $"GCP prefix table {installedDate:yyyy-MM-dd} installed. New sessions use it automatically.";
        }
        catch (Exception ex)
        {
            GcpToastMessage = $"GCP table update failed: {ex.Message}";
        }
        finally
        {
            _gcpInstalling = false;
            RelayCommand.Refresh();
        }
    }
}
