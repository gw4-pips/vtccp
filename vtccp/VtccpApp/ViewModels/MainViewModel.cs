namespace VtccpApp.ViewModels;

using ConfigEngine;
using VtccpApp.Commands;

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

    // ── Title bar ─────────────────────────────────────────────────────────────

    public string AppTitle { get; } = "VTCCP — VCCS DMV TruCheck Command Pilot";

    // ── Init ──────────────────────────────────────────────────────────────────

    public MainViewModel()
    {
        DevicesVM   = new DevicesViewModel(Repository);
        TemplatesVM = new TemplatesViewModel(Repository);
        SessionVM   = new SessionViewModel(Repository);

        NavDevicesCommand   = new RelayCommand(() => Navigate("Devices"));
        NavTemplatesCommand = new RelayCommand(() => Navigate("Templates"));
        NavSessionCommand   = new RelayCommand(() => Navigate("Session"));

        Navigate("Session");   // default page

        _ = LoadConfigAsync();
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
        }
        catch { /* first run — defaults in effect */ }
    }

    public async Task SaveConfigAsync()
    {
        try   { await Repository.SaveAsync(); }
        catch { /* non-fatal */ }
    }
}
