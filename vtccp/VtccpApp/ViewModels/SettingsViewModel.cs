namespace VtccpApp.ViewModels;

using ConfigEngine;
using ConfigEngine.Models;
using VtccpApp.Commands;
using VtccpApp.Services;

/// <summary>
/// View-model for the application Settings page.
///
/// Wraps <see cref="AppSettings"/> properties with observable bindings and
/// persists changes to <c>appsettings.json</c> via <see cref="ConfigRepository"/>
/// after every toggle so operators never lose a setting change.
/// </summary>
public sealed class SettingsViewModel : ViewModelBase
{
    private readonly ConfigRepository _repo;

    public SettingsViewModel(ConfigRepository repo)
    {
        _repo = repo;
        CheckGcpUpdateCommand   = new RelayCommand(() => _ = CheckGcpUpdateAsync(),   () => !_gcpBusy);
        InstallGcpUpdateCommand = new RelayCommand(() => _ = InstallGcpUpdateAsync(), () => !_gcpBusy && IsGcpUpdateAvailable);
    }

    // ── Data Sources — GCP prefix table ────────────────────────────────────────

    private bool   _gcpBusy;
    private bool   _isGcpUpdateAvailable;
    private string _gcpCheckStatus = string.Empty;

    public RelayCommand CheckGcpUpdateCommand   { get; }
    public RelayCommand InstallGcpUpdateCommand { get; }

    /// <summary>Human-readable date of the in-use GCP prefix table, or "not installed".</summary>
    public string CurrentGcpTableDate
    {
        get
        {
            string? iso = _repo.Settings.GcpLastModified;
            if (!string.IsNullOrWhiteSpace(iso) && DateTimeOffset.TryParse(iso, out var d))
                return d.ToString("yyyy-MM-dd");
            return "not installed";
        }
    }

    /// <summary>Path of the in-use GCP prefix table file (informational).</summary>
    public string CurrentGcpTablePath =>
        _repo.Settings.GcpDataPath ?? "\u2014 (bundled seed copy or none)";

    /// <summary>Base URL of the Azure Function update service. Empty = auto-update disabled.</summary>
    public string GcpUpdateServiceUrl
    {
        get => _repo.Settings.GcpUpdateServiceUrl ?? string.Empty;
        set
        {
            string? trimmed = string.IsNullOrWhiteSpace(value) ? null : value.Trim();
            if (_repo.Settings.GcpUpdateServiceUrl == trimmed) return;
            _repo.Settings.GcpUpdateServiceUrl = trimmed;
            OnPropertyChanged();
            _ = SaveAsync();
        }
    }

    /// <summary>Pre-shared device token identifying this workstation to the update service.</summary>
    public string GcpDeviceToken
    {
        get => _repo.Settings.GcpDeviceToken ?? string.Empty;
        set
        {
            string? trimmed = string.IsNullOrWhiteSpace(value) ? null : value.Trim();
            if (_repo.Settings.GcpDeviceToken == trimmed) return;
            _repo.Settings.GcpDeviceToken = trimmed;
            OnPropertyChanged();
            _ = SaveAsync();
        }
    }

    /// <summary>Status line under the "Check for update" button.</summary>
    public string GcpCheckStatus
    {
        get => _gcpCheckStatus;
        private set => Set(ref _gcpCheckStatus, value);
    }

    public bool IsGcpUpdateAvailable
    {
        get => _isGcpUpdateAvailable;
        private set { Set(ref _isGcpUpdateAvailable, value); RelayCommand.Refresh(); }
    }

    private async Task CheckGcpUpdateAsync()
    {
        _gcpBusy = true; RelayCommand.Refresh();
        try
        {
            var service = GcpUpdateServiceFactory.Create(_repo.Settings);
            if (service is null)
            {
                GcpCheckStatus = "Enter the update service URL and device token first.";
                return;
            }

            GcpCheckStatus = "Checking\u2026";
            var check = await service.CheckNowAsync();
            if (check is null)
            {
                GcpCheckStatus = "Update service unreachable (offline, wrong URL, or invalid token).";
                IsGcpUpdateAvailable = false;
            }
            else if (check.UpdateAvailable)
            {
                string localStr = check.LocalDate is { } ld ? ld.ToString("yyyy-MM-dd") : "none";
                GcpCheckStatus = $"Update available: {check.ServerDate:yyyy-MM-dd} (installed: {localStr}).";
                IsGcpUpdateAvailable = true;
            }
            else
            {
                GcpCheckStatus = $"Up to date ({check.LocalDate:yyyy-MM-dd}).";
                IsGcpUpdateAvailable = false;
            }
        }
        finally { _gcpBusy = false; RelayCommand.Refresh(); }
    }

    private async Task InstallGcpUpdateAsync()
    {
        _gcpBusy = true; RelayCommand.Refresh();
        try
        {
            var service = GcpUpdateServiceFactory.Create(_repo.Settings);
            if (service is null) return;

            GcpCheckStatus = "Downloading\u2026";
            var installedDate = await service.DownloadAndInstallAsync();

            _repo.Settings.GcpDataPath = GcpUpdateServiceFactory.ResolveLocalXmlPath(_repo.Settings);
            _repo.Settings.GcpLastModified = installedDate?.ToString("O");
            await SaveAsync();

            IsGcpUpdateAvailable = false;
            GcpCheckStatus = $"Installed table {installedDate:yyyy-MM-dd}. New sessions use it automatically.";
            OnPropertyChanged(nameof(CurrentGcpTableDate));
            OnPropertyChanged(nameof(CurrentGcpTablePath));
        }
        catch (Exception ex)
        {
            GcpCheckStatus = $"Update failed: {ex.Message}";
        }
        finally { _gcpBusy = false; RelayCommand.Refresh(); }
    }

    // ── Hybrid report ──────────────────────────────────────────────────────────

    /// <summary>Whether the hybrid HTML report generator is enabled for every scan.</summary>
    public bool GenerateHybridReport
    {
        get => _repo.Settings.GenerateHybridReport;
        set
        {
            if (_repo.Settings.GenerateHybridReport == value) return;
            _repo.Settings.GenerateHybridReport = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsHybridConfigVisible));
            _ = SaveAsync();
        }
    }

    /// <summary>Controls visibility of the hybrid mode sub-options.</summary>
    public bool IsHybridConfigVisible => _repo.Settings.GenerateHybridReport;

    /// <summary>
    /// True when the hybrid report mode is <see cref="HybridReportMode.Alongside"/>.
    /// Setting this to true switches the mode.
    /// </summary>
    public bool IsAlongsideMode
    {
        get => _repo.Settings.HybridReportMode == HybridReportMode.Alongside;
        set
        {
            if (!value || _repo.Settings.HybridReportMode == HybridReportMode.Alongside) return;
            _repo.Settings.HybridReportMode = HybridReportMode.Alongside;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsReplaceMode));
            _ = SaveAsync();
        }
    }

    /// <summary>
    /// True when the hybrid report mode is <see cref="HybridReportMode.Replace"/>.
    /// Setting this to true switches the mode.
    /// </summary>
    public bool IsReplaceMode
    {
        get => _repo.Settings.HybridReportMode == HybridReportMode.Replace;
        set
        {
            if (!value || _repo.Settings.HybridReportMode == HybridReportMode.Replace) return;
            _repo.Settings.HybridReportMode = HybridReportMode.Replace;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsAlongsideMode));
            _ = SaveAsync();
        }
    }

    /// <summary>
    /// Optional custom output directory for hybrid reports (Alongside mode only).
    /// Empty string = use the session output directory (same folder as the Excel workbook).
    /// </summary>
    public string HybridReportOutputDirectory
    {
        get => _repo.Settings.HybridReportOutputDirectory ?? string.Empty;
        set
        {
            string? trimmed = string.IsNullOrWhiteSpace(value) ? null : value.Trim();
            if (_repo.Settings.HybridReportOutputDirectory == trimmed) return;
            _repo.Settings.HybridReportOutputDirectory = trimmed;
            OnPropertyChanged();
            _ = SaveAsync();
        }
    }

    // ── Reload ─────────────────────────────────────────────────────────────────

    /// <summary>
    /// Refreshes all bound properties after an external settings reload.
    /// Call from <c>MainViewModel.LoadConfigAsync</c> after the repository is loaded.
    /// </summary>
    public void Reload()
    {
        OnPropertyChanged(nameof(GenerateHybridReport));
        OnPropertyChanged(nameof(IsHybridConfigVisible));
        OnPropertyChanged(nameof(IsAlongsideMode));
        OnPropertyChanged(nameof(IsReplaceMode));
        OnPropertyChanged(nameof(HybridReportOutputDirectory));
        OnPropertyChanged(nameof(CurrentGcpTableDate));
        OnPropertyChanged(nameof(CurrentGcpTablePath));
        OnPropertyChanged(nameof(GcpUpdateServiceUrl));
        OnPropertyChanged(nameof(GcpDeviceToken));
    }

    // ── Persistence ────────────────────────────────────────────────────────────

    private async Task SaveAsync()
    {
        try   { await _repo.SaveSettingsAsync(); }
        catch { /* non-fatal — in-memory state is correct; disk write failed */ }
    }
}
