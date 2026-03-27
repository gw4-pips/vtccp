namespace ConfigEngine.Models;

using ExcelEngine.Models;

/// <summary>
/// Named, serializable job template. Stores all operator-configurable fields
/// that would otherwise be set in <see cref="SessionState"/> before opening
/// a session.
/// </summary>
public sealed class JobTemplate
{
    // ── Identity ──────────────────────────────────────────────────────────────

    /// <summary>Stable unique identifier (GUID string). Set on creation.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>Operator-facing template name, e.g. "Incoming Inspection – GS1 DM".</summary>
    public string Name { get; set; } = "New Job Template";

    /// <summary>Optional free-text notes / description.</summary>
    public string? Notes { get; set; }

    /// <summary>Whether this template is the application-wide default.</summary>
    public bool IsDefault { get; set; }

    // ── Job identity ──────────────────────────────────────────────────────────

    /// <summary>
    /// Default job name written to the Excel output header.
    /// Maps to <see cref="SessionState.JobName"/>. May be overridden by the operator at runtime.
    /// </summary>
    public string? JobName { get; set; }

    /// <summary>Default operator ID. Maps to <see cref="SessionState.OperatorId"/>.</summary>
    public string? OperatorId { get; set; }

    // ── Batch / format ────────────────────────────────────────────────────────

    /// <summary>
    /// How the Batch/Lot column is populated.
    /// <see cref="BatchMode.Manual"/>: caller supplies the value per scan.
    /// <see cref="BatchMode.AutoFromGS1"/>: extracted automatically from AI(10) in the barcode.
    /// Maps to <see cref="SessionState.BatchMode"/>.
    /// </summary>
    public BatchMode BatchMode { get; set; } = BatchMode.Manual;

    /// <summary>
    /// Excel output format.
    /// Maps to <see cref="SessionState.OutputFormat"/>.
    /// </summary>
    public OutputFormat OutputFormat { get; set; } = OutputFormat.Xlsx;

    // ── Roll configuration ────────────────────────────────────────────────────

    /// <summary>Roll number increment mode.</summary>
    public RollIncrementMode RollIncrementMode { get; set; } = RollIncrementMode.Manual;

    /// <summary>Starting roll value for <see cref="RollIncrementMode.AutoIncrement"/> mode.</summary>
    public int RollStartValue { get; set; } = 1;

    // ── Output ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Output directory override. If null, falls back to <see cref="AppSettings.DefaultOutputDirectory"/>.
    /// Maps to <see cref="SessionState.OutputDirectory"/>.
    /// </summary>
    public string? OutputDirectory { get; set; }

    /// <summary>
    /// Path to the operator/company logo image file.
    /// Stored in <see cref="SessionState.LogoPath"/> and used by the WPF shell for header branding.
    /// </summary>
    public string? LogoPath { get; set; }

    // ── Derived helpers ───────────────────────────────────────────────────────

    /// <summary>
    /// Converts this template to a <see cref="SessionState"/> ready for
    /// <see cref="SessionManager.StartSession"/>.
    /// </summary>
    /// <param name="outputDirectoryFallback">
    /// Directory to use when <see cref="OutputDirectory"/> is null or empty.
    /// </param>
    public SessionState ToSessionState(string? outputDirectoryFallback = null) => new()
    {
        JobName           = JobName ?? string.Empty,
        OperatorId        = OperatorId ?? string.Empty,
        BatchMode         = BatchMode,
        OutputFormat      = OutputFormat,
        RollIncrementMode = RollIncrementMode,
        RollStartValue    = RollStartValue,
        OutputDirectory   = !string.IsNullOrWhiteSpace(OutputDirectory)
                                ? OutputDirectory
                                : outputDirectoryFallback,
        LogoPath          = LogoPath,
    };

    public override string ToString() => Name;
}
