namespace ExcelEngine.Writer;

using ExcelEngine.Adapters;
using ExcelEngine.Models;

/// <summary>
/// Writes modulation-array data to the "Modulation Values" worksheet.
///
/// Sheet layout per record block:
///   Row N+0:     Label row — datetime | symbology | grade | matrix-size | symbol-dims | data-snippet
///                Bold; VCCS-navy background.
///   Row N+1 ..   Grid rows — one cell per module, raw reflectance value written as a number.
///   N+GridRows:  Cell background: dark (isBlack=true) → near-black; light → near-white;
///                quiet-zone border (outermost row/col) → light grey.
///   N+GridRows+1: blank separator
///   N+GridRows+2: blank separator
///
/// Cell background ARGB:
///   Dark module   : 0xFF1A1A1A (near-black)
///   Light module  : 0xFFF5F5F5 (near-white)
///   QZ border     : 0xFFE0E0E0 (light grey — visually distinct from symbol data)
///   Label header  : 0xFF1E3A5F (VCCS navy)
///
/// IMPORTANT: The caller (ExcelWriter) must call adapter.EnsureSheet("Main") after
/// WriteRecord returns to restore Main as the active write target.
/// </summary>
public sealed class ModValuesSheetWriter
{
    public const string SheetName = "Modulation Values";

    private const uint ArgbDark   = 0xFF1A1A1A;
    private const uint ArgbLight  = 0xFFF5F5F5;
    private const uint ArgbQZ     = 0xFFE0E0E0;
    private const uint ArgbHeader = 0xFF1E3A5F;

    private readonly IExcelAdapter _adapter;
    private int  _nextRow;
    private bool _sheetEnsured;

    public ModValuesSheetWriter(IExcelAdapter adapter)
    {
        _adapter      = adapter;
        _nextRow      = 1;
        _sheetEnsured = false;
    }

    /// <summary>
    /// Write one record's modulation grid to the "Modulation Values" sheet.
    /// No-ops if <paramref name="record"/>.ModulationValues is null.
    /// Switches the adapter's active sheet to "Modulation Values" for the duration.
    /// The caller must call adapter.EnsureSheet("Main") after this returns.
    /// </summary>
    public void WriteRecord(VerificationRecord record)
    {
        var data = record.ModulationValues;
        if (data is null) return;

        EnsureSheetActive();

        // ── Label row ─────────────────────────────────────────────────────────
        string dtStr   = data.ScanDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        string grade   = record.OverallGrade?.ToString() ?? "-";
        string size    = record.MatrixSize ?? "-";
        string sym     = record.Symbology;
        string snippet = record.DecodedData is { Length: > 0 } d
            ? (d.Length > 20 ? d[..20] + "\u2026" : d)
            : "";
        string label = $"{dtStr}  |  {sym}  |  {grade}  |  {size}" +
                       $"  |  {data.SymbolRows}\u00d7{data.SymbolCols} symbol  |  {snippet}";

        int labelColCount = data.GridCols + 1;   // col1=row-idx, col2..GridCols+1=module cols
        _adapter.WriteString(_nextRow, 1, label);
        _adapter.SetCellBold(_nextRow, 1);
        _adapter.SetRowBackground(_nextRow, labelColCount, ArgbHeader);
        _nextRow++;

        // ── Grid rows ─────────────────────────────────────────────────────────
        for (int rowIdx = 0; rowIdx < data.GridRows; rowIdx++)
        {
            // Col 1: row index label (0-based, matching the firmware array index)
            _adapter.WriteNumber(_nextRow, 1, rowIdx, null);

            bool isQZRow = rowIdx == 0 || rowIdx == data.GridRows - 1;

            for (int colIdx = 0; colIdx < data.GridCols; colIdx++)
            {
                int flatIdx = rowIdx * data.GridCols + colIdx;
                if (flatIdx >= data.Raw.Count) break;

                int  rawVal  = data.Raw[flatIdx];
                bool isBlack = flatIdx < data.IsBlack.Count && data.IsBlack[flatIdx];
                bool isQZCol = colIdx == 0 || colIdx == data.GridCols - 1;
                bool isQZ    = isQZRow || isQZCol;

                int excelCol = colIdx + 2;   // offset: col 1 is row-index label
                _adapter.WriteNumber(_nextRow, excelCol, rawVal, null);
                _adapter.SetCellBackground(_nextRow, excelCol,
                    isQZ ? ArgbQZ : (isBlack ? ArgbDark : ArgbLight));
            }

            _nextRow++;
        }

        // ── Two-row blank separator between records ───────────────────────────
        _nextRow += 2;
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private void EnsureSheetActive()
    {
        int existingRows = _adapter.EnsureSheet(SheetName);
        if (!_sheetEnsured)
        {
            _nextRow      = existingRows > 0 ? existingRows + 2 : 1;
            _sheetEnsured = true;
        }
    }
}
