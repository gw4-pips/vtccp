namespace ExcelEngine.Writer;

using ExcelEngine.Adapters;
using ExcelEngine.Models;
using ExcelEngine.Schema;

/// <summary>
/// Writes the ISO 15416 per-scan sub-table for 1D barcode records.
///
/// Webscan layout: immediately below each 1D record's main data row, up to 10 scan rows
/// are written, each with the 11 per-scan parameters plus the per-scan grade.
/// Columns used are the 1D summary columns (Avg_Edge, Avg_RlRd, Avg_SC, …) —
/// the same schema columns that hold the averages in the main row.
/// A "Scan N" label is written in the first column (Date column position) of each sub-row.
///
/// Column anchoring (1-based, from the WebscanCompatible schema):
///   Date column      = 1  → used for "Scan N" label
///   Avg_Edge         = col of "Edge"  in schema
///   Avg_RlRd         = col of "Rl/Rd"
///   Avg_SC           = col of "SC"
///   Avg_MinEC        = col of "MinEC"
///   Avg_MOD          = col of "MOD"
///   Avg_Defect       = col of "DEF"
///   Avg_DCOD         = col of "DCOD"
///   Avg_DEC          = col of "DEC"
///   Avg_MinQZ        = col of "MinQZ"
///   SymbolAnsiGrade  = col of "Symbol ANSI Grade" → per-scan grade written here
///
/// The sub-rows are offset rows in the same Main sheet, not additional data rows in the
/// schema sense — they are written directly and _nextDataRow in ExcelWriter is advanced
/// past them automatically after calling WriteScans().
/// </summary>
public sealed class PerScanTableWriter
{
    private const string ScanLabelFormat = "Scan {0}";

    private readonly IExcelAdapter _adapter;
    private readonly ColumnSchema _schema;

    // Cached column indices (1-based) resolved once from the schema
    private int _colScanLabel;    // always 1 (Date column)
    private int _colEdge;
    private int _colRlRd;
    private int _colSC;
    private int _colMinEC;
    private int _colMOD;
    private int _colDefect;
    private int _colDCOD;
    private int _colDEC;
    private int _colMinQZ;
    private int _colPerScanGrade;
    private bool _resolved;

    public PerScanTableWriter(IExcelAdapter adapter, ColumnSchema schema)
    {
        _adapter = adapter;
        _schema = schema;
        _colScanLabel = 1;
    }

    /// <summary>Maximum scan sub-rows written per 1D record (matches Webscan TruCheck limit).</summary>
    public const int MaxScansPerRecord = 10;

    /// <summary>
    /// Write per-scan sub-rows starting at <paramref name="startRow"/>.
    /// Writes at most <see cref="MaxScansPerRecord"/> rows.
    /// Returns the number of rows written (0 if ScanResults is empty).
    /// </summary>
    public int WriteScans(IReadOnlyList<ScanResult1D> scans, int startRow)
    {
        if (scans.Count == 0) return 0;

        EnsureColumnsResolved();

        // ISO 15416 defines up to 10 scan lines; cap to match Webscan TruCheck display.
        int count = Math.Min(scans.Count, MaxScansPerRecord);

        int rowOffset = 0;
        for (int i = 0; i < count; i++)
        {
            var scan = scans[i];
            int row = startRow + rowOffset;
            int scanNum = scan.ScanNumber > 0 ? scan.ScanNumber : rowOffset + 1;

            // Scan label in the Date column
            _adapter.WriteString(row, _colScanLabel, string.Format(ScanLabelFormat, scanNum));

            // Per-scan parameter values
            if (scan.Edge.HasValue)       _adapter.WriteNumber(row, _colEdge,   (double)scan.Edge.Value);
            if (scan.Reflectance is not null) _adapter.WriteString(row, _colRlRd,   scan.Reflectance);
            if (scan.SC.HasValue)         _adapter.WriteNumber(row, _colSC,     (double)scan.SC.Value);
            if (scan.MinEC.HasValue)      _adapter.WriteNumber(row, _colMinEC,  (double)scan.MinEC.Value);
            if (scan.MOD.HasValue)        _adapter.WriteNumber(row, _colMOD,    (double)scan.MOD.Value);
            if (scan.Defect.HasValue)     _adapter.WriteNumber(row, _colDefect, (double)scan.Defect.Value);
            if (scan.DCOD is not null)    _adapter.WriteString(row, _colDCOD,   scan.DCOD);
            if (scan.DEC.HasValue)        _adapter.WriteNumber(row, _colDEC,    (double)scan.DEC.Value);
            if (scan.LQZ.HasValue && scan.RQZ.HasValue)
            {
                // MinQZ is the lesser of LQZ/RQZ (or HQZ when present) per the standard
                decimal minQZ = scan.HQZ.HasValue
                    ? Math.Min(scan.LQZ.Value, Math.Min(scan.RQZ.Value, scan.HQZ.Value))
                    : Math.Min(scan.LQZ.Value, scan.RQZ.Value);
                _adapter.WriteNumber(row, _colMinQZ, (double)minQZ);
            }
            else if (scan.LQZ.HasValue)   _adapter.WriteNumber(row, _colMinQZ, (double)scan.LQZ.Value);

            // Per-scan grade
            if (scan.PerScanGrade?.NumericGrade.HasValue == true)
                _adapter.WriteNumber(row, _colPerScanGrade, (double)scan.PerScanGrade.NumericGrade!.Value);

            rowOffset++;
        }

        return rowOffset;
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private void EnsureColumnsResolved()
    {
        if (_resolved) return;

        _colScanLabel    = 1;
        _colEdge         = RequireCol("Avg_Edge");
        _colRlRd         = RequireCol("Avg_RlRd");
        _colSC           = RequireCol("Avg_SC");
        _colMinEC        = RequireCol("Avg_MinEC");
        _colMOD          = RequireCol("Avg_MOD");
        _colDefect       = RequireCol("Avg_Defect");
        _colDCOD         = RequireCol("Avg_DCOD");
        _colDEC          = RequireCol("Avg_DEC");
        _colMinQZ        = RequireCol("Avg_MinQZ");
        _colPerScanGrade = RequireCol("SymbolAnsiGrade_Numeric");

        _resolved = true;
    }

    private int RequireCol(string fieldId)
    {
        var idx = _schema.GetColumnIndex(fieldId)
            ?? throw new InvalidOperationException(
                $"PerScanTableWriter: column '{fieldId}' not found in schema.");
        return idx;
    }
}
