namespace VtccpApp.ViewModels;

using ConfigEngine.Models;
using ExcelEngine.Models;

/// <summary>
/// Editable view-model wrapper around a <see cref="JobTemplate"/>.
/// </summary>
public sealed class JobTemplateViewModel : ViewModelBase
{
    private string           _id               = Guid.NewGuid().ToString();
    private string           _name             = "New Job Template";
    private string?          _jobName;
    private string?          _operatorId;
    private BatchMode        _batchMode        = BatchMode.Manual;
    private OutputFormat     _outputFormat     = OutputFormat.Xlsx;
    private RollIncrementMode _rollMode        = RollIncrementMode.Manual;
    private int              _rollStartValue   = 1;
    private string?          _outputDirectory;
    private string?          _logoPath;
    private bool             _isDefault;
    private string           _notes            = string.Empty;

    public string            Id              { get => _id;             set => Set(ref _id,             value); }
    public string            Name            { get => _name;           set => Set(ref _name,           value); }
    public string?           JobName         { get => _jobName;        set => Set(ref _jobName,        value); }
    public string?           OperatorId      { get => _operatorId;     set => Set(ref _operatorId,     value); }
    public BatchMode         BatchMode       { get => _batchMode;      set => Set(ref _batchMode,      value); }
    public OutputFormat      OutputFormat    { get => _outputFormat;   set => Set(ref _outputFormat,   value); }
    public RollIncrementMode RollMode        { get => _rollMode;       set => Set(ref _rollMode,       value); }
    public int               RollStartValue  { get => _rollStartValue; set => Set(ref _rollStartValue, value); }
    public string?           OutputDirectory { get => _outputDirectory;set => Set(ref _outputDirectory,value); }
    public string?           LogoPath        { get => _logoPath;       set => Set(ref _logoPath,       value); }
    public bool              IsDefault       { get => _isDefault;      set => Set(ref _isDefault,      value); }
    public string            Notes           { get => _notes;          set => Set(ref _notes,          value); }

    // ── Enum options for ComboBoxes ───────────────────────────────────────────

    public static IReadOnlyList<BatchMode>        BatchModeOptions   { get; } =
        Enum.GetValues<BatchMode>();
    public static IReadOnlyList<OutputFormat>     OutputFormatOptions { get; } =
        Enum.GetValues<OutputFormat>();
    public static IReadOnlyList<RollIncrementMode> RollModeOptions   { get; } =
        Enum.GetValues<RollIncrementMode>();

    public JobTemplateViewModel() { }

    public JobTemplateViewModel(JobTemplate t) => LoadFrom(t);

    public void LoadFrom(JobTemplate t)
    {
        Id              = t.Id;
        Name            = t.Name;
        JobName         = t.JobName;
        OperatorId      = t.OperatorId;
        BatchMode       = t.BatchMode;
        OutputFormat    = t.OutputFormat;
        RollMode        = t.RollIncrementMode;
        RollStartValue  = t.RollStartValue;
        OutputDirectory = t.OutputDirectory;
        LogoPath        = t.LogoPath;
        IsDefault       = t.IsDefault;
        Notes           = t.Notes ?? string.Empty;
    }

    public JobTemplate ToModel() => new()
    {
        Id                = Id,
        Name              = Name,
        JobName           = string.IsNullOrWhiteSpace(JobName) ? null : JobName,
        OperatorId        = string.IsNullOrWhiteSpace(OperatorId) ? null : OperatorId,
        BatchMode         = BatchMode,
        OutputFormat      = OutputFormat,
        RollIncrementMode = RollMode,
        RollStartValue    = RollStartValue,
        OutputDirectory   = string.IsNullOrWhiteSpace(OutputDirectory) ? null : OutputDirectory,
        LogoPath          = string.IsNullOrWhiteSpace(LogoPath) ? null : LogoPath,
        IsDefault         = IsDefault,
        Notes             = string.IsNullOrWhiteSpace(Notes) ? null : Notes,
    };

    public override string ToString() => Name;
}
