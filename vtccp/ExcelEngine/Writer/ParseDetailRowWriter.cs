namespace ExcelEngine.Writer;

using ExcelEngine.Adapters;
using ExcelEngine.Models;
using ExcelEngine.Schema;

/// <summary>
/// Writes the Level-1 parse-detail child row that appears immediately below each
/// parent verification record row when a GS1 or ISO 15434 / MIL-STD-130 data format
/// check is present.
///
/// Layout:
///   Col 1  (Date position) — "↳" sentinel label (italic, amber background)
///   Col 2  (Time position) — formatted HRI string, e.g.:
///                "GTIN: 00355513710213 | BATCH/LOT: A1234 | USE BY OR EXPIRY: 261231"
///
/// Row properties:
///   OutlineLevel = 1  (child of Level-0 parent)
///   Hidden       = false (open / visible by default)
///   Background   = pale amber (#FFF2CC)
///
/// For COM live mode, ExcelWriter exposes LastParseDetailRow so SessionManager can
/// call IExcelAdapter.ScheduleRowHide() to auto-collapse the row after a configurable
/// delay (default: 20 s).  File-based adapters ignore the schedule call.
/// </summary>
public sealed class ParseDetailRowWriter
{
    private const uint AmberFill    = 0xFFFFF2CC;  // pale amber / Office light-yellow
    private const string Sentinel   = "↳";

    private readonly IExcelAdapter _adapter;
    private readonly int           _colCount;

    public ParseDetailRowWriter(IExcelAdapter adapter, ColumnSchema schema)
    {
        _adapter  = adapter;
        _colCount = schema.Columns.Count;
    }

    /// <summary>
    /// Write a single parse-detail child row at <paramref name="row"/> (1-based).
    /// The caller is responsible for advancing _nextDataRow by 1 after this call.
    /// </summary>
    public void WriteParseDetailRow(int row, DataFormatCheckResult dfc)
    {
        _adapter.WriteString(row, 1, Sentinel);

        string hri = BuildHri(dfc);
        if (!string.IsNullOrEmpty(hri))
            _adapter.WriteString(row, 2, hri);

        _adapter.SetRowBackground(row, _colCount, AmberFill);
        _adapter.SetRowOutlineLevel(row, 1);
        // Row starts visible; COM timer hides it via ScheduleRowHide from SessionManager.
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Build a pipe-separated human-readable interpretation (HRI) string from the
    /// DFC row table.  Each entry is formatted as "{Name}: {Data}".
    /// A standard prefix (e.g. "[GS1]") is prepended when Standard is set and rows exist.
    /// Returns an empty string when both Standard and Rows are absent.
    /// </summary>
    private static string BuildHri(DataFormatCheckResult dfc)
    {
        if (dfc.Rows.Count == 0)
            return dfc.Standard ?? string.Empty;

        var parts = new System.Text.StringBuilder();
        for (int i = 0; i < dfc.Rows.Count; i++)
        {
            if (i > 0) parts.Append(" | ");
            var r = dfc.Rows[i];
            parts.Append(r.Name);
            if (!string.IsNullOrWhiteSpace(r.Data))
            {
                parts.Append(": ");
                parts.Append(r.Data);
            }
        }

        if (dfc.Standard is not null)
            return $"[{dfc.Standard}]  {parts}";

        return parts.ToString();
    }
}
