namespace VtccpApp.ViewModels;

using ConfigEngine;
using ConfigEngine.Models;

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
    }

    // ── Persistence ────────────────────────────────────────────────────────────

    private async Task SaveAsync()
    {
        try   { await _repo.SaveSettingsAsync(); }
        catch { /* non-fatal — in-memory state is correct; disk write failed */ }
    }
}
