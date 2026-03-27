namespace VtccpApp.Models;

/// <summary>
/// Pure, cross-platform filter predicate over <see cref="ScanResultRow"/> records.
/// All string comparisons are case-insensitive.
/// </summary>
public sealed class HistoryFilter
{
    // ── Filter values (empty / "All" = no restriction) ────────────────────────

    /// <summary>
    /// Grade letter to match ("A", "B", "C", "D", "F") or empty/"All" for all.
    /// </summary>
    public string GradeFilter { get; set; } = string.Empty;

    /// <summary>"Pass", "Fail", or empty/"All" for all.</summary>
    public string PassFailFilter { get; set; } = string.Empty;

    /// <summary>Substring match on Symbology, or empty/"All" for all.</summary>
    public string SymbologyFilter { get; set; } = string.Empty;

    // ── Predicate ─────────────────────────────────────────────────────────────

    public bool Matches(ScanResultRow row)
    {
        if (!string.IsNullOrEmpty(GradeFilter) &&
            !GradeFilter.Equals("All", StringComparison.OrdinalIgnoreCase) &&
            !row.Grade.Equals(GradeFilter, StringComparison.OrdinalIgnoreCase))
            return false;

        if (!string.IsNullOrEmpty(PassFailFilter) &&
            !PassFailFilter.Equals("All", StringComparison.OrdinalIgnoreCase) &&
            !row.PassFail.Equals(PassFailFilter, StringComparison.OrdinalIgnoreCase))
            return false;

        if (!string.IsNullOrEmpty(SymbologyFilter) &&
            !SymbologyFilter.Equals("All", StringComparison.OrdinalIgnoreCase) &&
            row.Symbology.IndexOf(SymbologyFilter, StringComparison.OrdinalIgnoreCase) < 0)
            return false;

        return true;
    }

    /// <summary>Returns true when no filter restrictions are active.</summary>
    public bool IsEmpty =>
        (string.IsNullOrEmpty(GradeFilter)     || GradeFilter     == "All") &&
        (string.IsNullOrEmpty(PassFailFilter)  || PassFailFilter   == "All") &&
        (string.IsNullOrEmpty(SymbologyFilter) || SymbologyFilter  == "All");
}
