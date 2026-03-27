namespace VtccpApp.ViewModels;

using System.Collections.ObjectModel;
using System.Windows;
using ConfigEngine;
using ConfigEngine.Models;
using VtccpApp.Commands;

/// <summary>
/// Manages the job template list for the Templates page.
/// </summary>
public sealed class TemplatesViewModel : ViewModelBase
{
    private readonly ConfigRepository _repo;

    private JobTemplateViewModel? _selected;
    private JobTemplateViewModel? _editing;
    private bool                  _isEditing;
    private string                _statusMessage = string.Empty;

    // ── Bindable properties ───────────────────────────────────────────────────

    public ObservableCollection<JobTemplateViewModel> Templates { get; } = [];

    public JobTemplateViewModel? Selected
    {
        get => _selected;
        set { Set(ref _selected, value); RelayCommand.Refresh(); }
    }

    public JobTemplateViewModel? Editing
    {
        get => _editing;
        private set => Set(ref _editing, value);
    }

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

    public RelayCommand AddCommand        { get; }
    public RelayCommand EditCommand       { get; }
    public RelayCommand DeleteCommand     { get; }
    public RelayCommand DefaultCommand    { get; }
    public RelayCommand SaveCommand       { get; }
    public RelayCommand CancelCommand     { get; }
    public RelayCommand BrowseOutputCommand { get; }
    public RelayCommand BrowseLogoCommand   { get; }

    public TemplatesViewModel(ConfigRepository repo)
    {
        _repo = repo;

        AddCommand          = new RelayCommand(OnAdd);
        EditCommand         = new RelayCommand(OnEdit,    () => Selected is not null && !IsEditing);
        DeleteCommand       = new RelayCommand(OnDelete,  () => Selected is not null && !IsEditing);
        DefaultCommand      = new RelayCommand(OnDefault, () => Selected is not null && !IsEditing);
        SaveCommand         = new RelayCommand(OnSave,    () => IsEditing);
        CancelCommand       = new RelayCommand(OnCancel,  () => IsEditing);
        BrowseOutputCommand = new RelayCommand(OnBrowseOutput, () => IsEditing);
        BrowseLogoCommand   = new RelayCommand(OnBrowseLogo,   () => IsEditing);

        Reload();
    }

    // ── Reload ────────────────────────────────────────────────────────────────

    public void Reload()
    {
        Templates.Clear();
        foreach (var t in _repo.Templates)
            Templates.Add(new JobTemplateViewModel(t));
        StatusMessage = string.Empty;
    }

    // ── Command handlers ──────────────────────────────────────────────────────

    private void OnAdd()
    {
        Editing   = new JobTemplateViewModel();
        IsEditing = true;
    }

    private void OnEdit()
    {
        if (Selected is null) return;
        Editing   = new JobTemplateViewModel(Selected.ToModel());
        IsEditing = true;
    }

    private void OnDelete()
    {
        if (Selected is null) return;
        if (MessageBox.Show($"Delete job template '{Selected.Name}'?",
                "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
            != MessageBoxResult.Yes) return;

        _repo.RemoveTemplate(Selected.Id);
        Reload();
        StatusMessage = "Job template deleted.";
    }

    private void OnDefault()
    {
        if (Selected is null) return;
        foreach (var t in _repo.Templates) t.IsDefault = false;
        var target = _repo.FindTemplate(Selected.Id);
        if (target is not null) target.IsDefault = true;
        Reload();
        StatusMessage = $"'{Selected.Name}' is now the default template.";
    }

    private void OnSave()
    {
        if (Editing is null) return;
        if (string.IsNullOrWhiteSpace(Editing.Name)) { StatusMessage = "Name is required."; return; }

        Editing.Name = Editing.Name.Trim();
        JobTemplate model = Editing.ToModel();

        bool updated = _repo.UpdateTemplate(model);
        if (!updated) _repo.AddTemplate(model);

        Reload();
        IsEditing     = false;
        StatusMessage = updated ? "Job template updated." : "Job template added.";
    }

    private void OnCancel()
    {
        IsEditing = false;
        Editing   = null;
    }

    private void OnBrowseOutput()
    {
        if (Editing is null) return;
        using var dlg = new System.Windows.Forms.FolderBrowserDialog
        {
            Description         = "Select Output Directory",
            UseDescriptionForTitle = true,
            SelectedPath        = Editing.OutputDirectory ?? string.Empty,
        };
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            Editing.OutputDirectory = dlg.SelectedPath;
    }

    private void OnBrowseLogo()
    {
        if (Editing is null) return;
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title  = "Select Logo Image",
            Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All files|*.*",
        };
        if (dlg.ShowDialog() == true) Editing.LogoPath = dlg.FileName;
    }
}
