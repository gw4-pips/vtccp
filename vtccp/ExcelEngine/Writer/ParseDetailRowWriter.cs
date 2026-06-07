namespace ExcelEngine.Writer;

using ExcelEngine.Adapters;
using ExcelEngine.Models;
using ExcelEngine.Schema;

/// <summary>
/// Writes the Level-1 parse-detail child row immediately below each parent
/// verification record row when a GS1 or ISO 15434 / MIL-STD-130 data format
/// check is present.
///
/// Layout:
///   Col 1  — "↳" sentinel (amber background signals "non-standard row")
///   Col 2  — compact HRI string, e.g.:
///             ]d2 | GS1 | Header | GTIN: 0123456789012 | BatchLot: A1234 | USE BY: 261231
///
/// Format rules (brevity-first):
///   • Lead with AIM ID (Symbology ID) when available, e.g. "]d2"
///   • Abbreviate Standard: "GS1 Application Data Format" → "GS1", "MIL-STD-130" → "MIL-130"
///   • Skip "AI:*" rows — they carry only the AI code number, redundant with the data row
///   • Skip "Chk Digit" rows — internal GS1 check-digit decomposition, not operator-relevant
///   • "GS1 Header" → emit as "Header" (standard abbreviation already in the prefix segment)
///   • All other rows → "{Name}: {Data}"
///   • Elements separated by " | "
///
/// Row properties:
///   OutlineLevel = 1  (Level-1 child of Level-0 parent)
///   Hidden       = false (visible by default)
///   Background   = pale amber (#FFF2CC)
///
/// For COM live mode, ExcelWriter exposes LastParseDetailRow so SessionManager can
/// call IExcelAdapter.ScheduleRowHide() to auto-collapse the row after a configurable
/// delay (default: 20 s).  File-based adapters ignore the schedule call.
/// </summary>
public sealed class ParseDetailRowWriter
{
    private const uint   AmberFill = 0xFFFFF2CC;
    private const string Sentinel  = "↳";

    private readonly IExcelAdapter _adapter;
    private readonly int           _colCount;

    public ParseDetailRowWriter(IExcelAdapter adapter, ColumnSchema schema)
    {
        _adapter  = adapter;
        _colCount = schema.Columns.Count;
    }

    /// <summary>
    /// Write a single parse-detail child row at <paramref name="row"/> (1-based).
    /// <paramref name="aimId"/> is the AIM symbology identifier from the parent record
    /// (e.g. "]d2", "]Q1", "]E0") — prepended as the first pipe-delimited segment.
    /// Pass null when not available.
    /// Caller advances _nextDataRow by 1 after this call.
    /// </summary>
    public void WriteParseDetailRow(int row, DataFormatCheckResult dfc, string? aimId = null)
    {
        _adapter.WriteString(row, 1, Sentinel);

        string hri = BuildHri(dfc, aimId);
        if (!string.IsNullOrEmpty(hri))
            _adapter.WriteString(row, 2, hri);

        _adapter.SetRowBackground(row, _colCount, AmberFill);
        _adapter.SetRowOutlineLevel(row, 1);
        // Starts visible; COM auto-collapse timer fires via SessionManager → ScheduleRowHide.
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Build the compact pipe-delimited HRI string.
    ///
    /// Example output:
    ///   ]d2 | GS1 | Header | GTIN: 0123456789012 | BatchLot: A1234
    /// </summary>
    private static string BuildHri(DataFormatCheckResult dfc, string? aimId)
    {
        var sb = new System.Text.StringBuilder();

        // Segment 1 — AIM / Symbology ID
        if (!string.IsNullOrWhiteSpace(aimId))
            sb.Append(aimId);

        // Segment 2 — abbreviated standard name
        string? shortStd = AbbreviateStandard(dfc.Standard);
        if (shortStd is not null)
        {
            if (sb.Length > 0) sb.Append(" | ");
            sb.Append(shortStd);
        }

        // Remaining segments — filtered DFC rows
        foreach (var r in dfc.Rows)
        {
            // Skip AI-code rows (e.g. Name="AI:GTIN", Data="01") — redundant with data row
            if (r.Name.StartsWith("AI:", StringComparison.OrdinalIgnoreCase))
                continue;

            // Skip check-digit decomposition
            if (r.Name.Equals("Chk Digit", StringComparison.OrdinalIgnoreCase))
                continue;

            if (sb.Length > 0) sb.Append(" | ");

            // "GS1 Header" → emit as bare "Header" token (no data value —
            // the <F1> encoding token is an artifact, not operator-relevant content)
            if (r.Name.Equals("GS1 Header", StringComparison.OrdinalIgnoreCase))
            {
                sb.Append("Header");
                continue;
            }

            // All other rows → "{Name}: {Data}"
            sb.Append(r.Name);
            if (!string.IsNullOrWhiteSpace(r.Data))
            {
                sb.Append(": ");
                sb.Append(r.Data);
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Map a verbose standard name to a short display token.
    /// Returns null when Standard is null or empty.
    /// </summary>
    private static string? AbbreviateStandard(string? standard)
    {
        if (string.IsNullOrWhiteSpace(standard)) return null;

        return standard switch
        {
            var s when s.StartsWith("GS1", StringComparison.OrdinalIgnoreCase)         => "GS1",
            var s when s.Contains("MIL-STD-130", StringComparison.OrdinalIgnoreCase)   => "MIL-130",
            var s when s.Contains("15434", StringComparison.OrdinalIgnoreCase)          => "ISO-15434",
            var s when s.Contains("Custom", StringComparison.OrdinalIgnoreCase)         => "Custom",
            _                                                                            => standard,
        };
    }
}
