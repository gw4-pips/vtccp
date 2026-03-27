namespace VtccpApp.Models;

using ExcelEngine.Models;

/// <summary>
/// Flat, display-friendly projection of a <see cref="VerificationRecord"/>.
/// Contains only the columns shown in the Results History DataGrid.
/// Immutable once created; created via <see cref="From"/>.
/// </summary>
public sealed class ScanResultRow
{
    // ── Grid display columns ──────────────────────────────────────────────────

    /// <summary>1-based sequence number within the current session.</summary>
    public int    RowNumber  { get; init; }

    /// <summary>Time-of-scan formatted as HH:mm:ss.</summary>
    public string Time       { get; init; } = string.Empty;

    /// <summary>Symbology short name (e.g. "Data Matrix ECC 200").</summary>
    public string Symbology  { get; init; } = string.Empty;

    /// <summary>
    /// Overall grade letter (A / B / C / D / F) or "—" if not graded.
    /// Used for badge colour and the grade filter combo.
    /// </summary>
    public string Grade      { get; init; } = "—";

    /// <summary>
    /// Overall numeric grade (0.0–4.0) formatted to one decimal place ("4.0",
    /// "3.5", …) or "—" when not measured. This is the primary display value;
    /// the letter grade is secondary.
    /// </summary>
    public string NumericGradeDisplay { get; init; } = "—";

    /// <summary>Overall numeric grade raw value, or null.</summary>
    public decimal? NumericGrade { get; init; }

    /// <summary>Pass / Fail / N/A</summary>
    public string PassFail   { get; init; } = "N/A";

    /// <summary>True when the overall result is a pass.</summary>
    public bool IsPass       { get; init; }

    /// <summary>UEC percentage for 2-D symbols (null for 1-D).</summary>
    public decimal? UecPercent { get; init; }

    /// <summary>
    /// Decoded barcode data truncated to 60 characters for the grid column.
    /// Use <see cref="Source"/>.<see cref="VerificationRecord.DecodedData"/>
    /// for the full string.
    /// </summary>
    public string DecodedDataPreview { get; init; } = string.Empty;

    /// <summary>True when the decoded data was truncated for grid display.</summary>
    public bool IsTruncated { get; init; }

    /// <summary>Operator ID at time of scan.</summary>
    public string OperatorId { get; init; } = string.Empty;

    /// <summary>Job name at time of scan.</summary>
    public string JobName    { get; init; } = string.Empty;

    // ── Full source record ────────────────────────────────────────────────────

    /// <summary>
    /// The complete <see cref="VerificationRecord"/> — used by the detail
    /// strip to show the full decoded string and all grading parameters.
    /// </summary>
    public VerificationRecord Source { get; init; } = null!;

    // ── Convenience passthrough for the detail strip ──────────────────────────

    /// <summary>Full decoded barcode data (un-truncated).</summary>
    public string FullDecodedData => Source.DecodedData ?? string.Empty;

    // ── Factory ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Projects a <see cref="VerificationRecord"/> into a <see cref="ScanResultRow"/>
    /// for display in the Results History grid.
    /// </summary>
    public static ScanResultRow From(VerificationRecord r, int rowNumber)
    {
        string? raw     = r.DecodedData;
        bool truncated  = raw is { Length: > 60 };
        string preview  = truncated ? raw![..60] + "…" : raw ?? string.Empty;

        string letter  = r.OverallGrade?.LetterGradeString is { Length: > 0 } g ? g : "—";
        string numeric = r.OverallGrade?.NumericGrade is { } n
            ? n.ToString("F1")
            : "—";

        return new ScanResultRow
        {
            RowNumber          = rowNumber,
            Time               = r.VerificationDateTime.ToString("HH:mm:ss"),
            Symbology          = r.Symbology,
            Grade              = letter,
            NumericGrade       = r.OverallGrade?.NumericGrade,
            NumericGradeDisplay = numeric,
            PassFail           = r.OverallGrade?.PassFail switch
            {
                OverallPassFail.Pass          => "Pass",
                OverallPassFail.Fail          => "Fail",
                OverallPassFail.NotApplicable => "N/A",
                _                             => "N/A",
            },
            IsPass             = r.OverallGrade?.PassFail == OverallPassFail.Pass,
            UecPercent         = r.UEC_Percent,
            DecodedDataPreview = preview,
            IsTruncated        = truncated,
            OperatorId         = r.OperatorId  ?? string.Empty,
            JobName            = r.JobName     ?? string.Empty,
            Source             = r,
        };
    }
}
