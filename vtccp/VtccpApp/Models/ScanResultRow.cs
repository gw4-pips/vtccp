namespace VtccpApp.Models;

using ExcelEngine.Models;

/// <summary>
/// Flat, display-friendly projection of a <see cref="VerificationRecord"/>.
/// Contains only the columns shown in the Results History DataGrid.
/// Immutable once created; created via <see cref="From"/>.
/// </summary>
public sealed class ScanResultRow
{
    // ── Display columns ───────────────────────────────────────────────────────

    /// <summary>1-based sequence number within the current session.</summary>
    public int    RowNumber  { get; init; }

    /// <summary>Time-of-scan formatted as HH:mm:ss.</summary>
    public string Time       { get; init; } = string.Empty;

    /// <summary>Symbology short name (e.g. "Data Matrix ECC 200").</summary>
    public string Symbology  { get; init; } = string.Empty;

    /// <summary>Overall grade letter (A / B / C / D / F) or "—" if not graded.</summary>
    public string Grade      { get; init; } = "—";

    /// <summary>Overall numeric grade (0.0–4.0) or null.</summary>
    public decimal? NumericGrade { get; init; }

    /// <summary>Pass / Fail / N/A</summary>
    public string PassFail   { get; init; } = "N/A";

    /// <summary>True when the overall result is a pass.</summary>
    public bool IsPass       { get; init; }

    /// <summary>UEC percentage for 2-D symbols (null for 1-D).</summary>
    public decimal? UecPercent { get; init; }

    /// <summary>Decoded barcode data, truncated to 60 characters for display.</summary>
    public string DecodedData { get; init; } = string.Empty;

    /// <summary>Operator ID at time of scan.</summary>
    public string OperatorId { get; init; } = string.Empty;

    /// <summary>Job name at time of scan.</summary>
    public string JobName    { get; init; } = string.Empty;

    // ── Source lookup (kept for detail flyout, not shown in main grid) ────────

    /// <summary>Full source record (for a future detail panel).</summary>
    public VerificationRecord Source { get; init; } = null!;

    // ── Factory ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Projects a <see cref="VerificationRecord"/> into a <see cref="ScanResultRow"/>
    /// for display in the Results History grid.
    /// </summary>
    public static ScanResultRow From(VerificationRecord r, int rowNumber) => new()
    {
        RowNumber    = rowNumber,
        Time         = r.VerificationDateTime.ToString("HH:mm:ss"),
        Symbology    = r.Symbology,
        Grade        = r.OverallGrade?.LetterGradeString is { Length: > 0 } g ? g : "—",
        NumericGrade = r.OverallGrade?.NumericGrade,
        PassFail     = r.OverallGrade?.PassFail switch
        {
            OverallPassFail.Pass          => "Pass",
            OverallPassFail.Fail          => "Fail",
            OverallPassFail.NotApplicable => "N/A",
            _                             => "N/A",
        },
        IsPass       = r.OverallGrade?.PassFail == OverallPassFail.Pass,
        UecPercent   = r.UEC_Percent,
        DecodedData  = r.DecodedData is { Length: > 60 }
                           ? r.DecodedData[..60] + "…"
                           : r.DecodedData ?? string.Empty,
        OperatorId   = r.OperatorId  ?? string.Empty,
        JobName      = r.JobName     ?? string.Empty,
        Source       = r,
    };
}
