namespace ExcelEngine.Writer;

using ExcelEngine.Adapters;
using ExcelEngine.Models;

/// <summary>
/// Writes 1D element width data to the "Element Widths" worksheet.
///
/// Sheet layout per record block:
///   Row N+0: Record label header (e.g. "UPC-A | 2025-08-11 | Scan 1") — bold
///   Row N+1: Section header "Element Sizes" — bold, light-blue bg
///   Row N+2: Column header row (Element | SPACE | BAR | SPACE | ...) — bold, blue bg
///   Row N+3..M: Element size data rows
///   Row M+1: (blank separator)
///   Row M+2: Section header "Element Deviations" — bold, light-blue bg
///   Row M+3: Column header row (same columns) — bold, blue bg
///   Row M+4..P: Element deviation data rows
///   Row P+1..P+2: (blank separator between records)
///
/// Multiple records are appended sequentially in the same sheet.
///
/// IMPORTANT: The caller (ExcelWriter) must call adapter.EnsureSheet("Main") after
/// WriteRecord returns to restore the Main sheet as the active write target.
/// </summary>
public sealed class ElementWidthsWriter
{
    public const string SheetName = "Element Widths";

    // Background colours: section headers use light blue, column headers use mid-blue
    private const uint SectionHeaderBgArgb = 0xDCE6F1;   // light blue
    private const uint ColHeaderBgArgb     = 0x4472C4;   // same mid-blue as Main sheet

    private readonly IExcelAdapter _adapter;
    private int _nextRow;
    private bool _sheetEnsured;

    public ElementWidthsWriter(IExcelAdapter adapter)
    {
        _adapter = adapter;
        _nextRow = 1;
        _sheetEnsured = false;
    }

    /// <summary>
    /// Write one record's element width data block to the "Element Widths" sheet.
    /// Switches the adapter's active sheet to "Element Widths" for the duration.
    /// The caller must call adapter.EnsureSheet("Main") after this returns.
    /// </summary>
    public void WriteRecord(ElementWidthData data)
    {
        EnsureSheetActive();

        // ── Record label header ──────────────────────────────────────────────
        if (!string.IsNullOrWhiteSpace(data.RecordLabel))
        {
            _adapter.WriteString(_nextRow, 1, data.RecordLabel);
            _adapter.SetCellBold(_nextRow, 1);
            _nextRow++;
        }

        // ── Element Sizes block ──────────────────────────────────────────────
        WriteSectionBlock(data.ColumnHeaders, data.ElementSizes, "Element Sizes");

        // ── Blank separator between the two tables ───────────────────────────
        _nextRow++;

        // ── Element Deviations block ─────────────────────────────────────────
        WriteSectionBlock(data.ColumnHeaders, data.ElementDeviations, "Element Deviations");

        // ── Blank separator between records ──────────────────────────────────
        _nextRow += 2;
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private void EnsureSheetActive()
    {
        int existingRows = _adapter.EnsureSheet(SheetName);
        if (!_sheetEnsured)
        {
            _nextRow = existingRows > 0 ? existingRows + 2 : 1;
            _sheetEnsured = true;
        }
    }

    private void WriteSectionBlock(
        IReadOnlyList<string> columnHeaders,
        IReadOnlyList<ElementWidthRow> rows,
        string sectionTitle)
    {
        // Section title row
        _adapter.WriteString(_nextRow, 1, sectionTitle);
        _adapter.SetCellBold(_nextRow, 1);
        _adapter.SetCellBackground(_nextRow, 1, SectionHeaderBgArgb);
        _nextRow++;

        // Column header row: col 1 = "Element", then one col per header
        int totalCols = columnHeaders.Count + 1;
        _adapter.WriteString(_nextRow, 1, "Element");
        _adapter.SetCellBold(_nextRow, 1);
        _adapter.SetCellBackground(_nextRow, 1, ColHeaderBgArgb);
        for (int c = 0; c < columnHeaders.Count; c++)
        {
            int colNum = c + 2;
            _adapter.WriteString(_nextRow, colNum, columnHeaders[c]);
            _adapter.SetCellBold(_nextRow, colNum);
            _adapter.SetCellBackground(_nextRow, colNum, ColHeaderBgArgb);
        }
        _nextRow++;

        // Data rows
        foreach (var row in rows)
        {
            _adapter.WriteString(_nextRow, 1, row.ElementName);
            for (int c = 0; c < row.Values.Count; c++)
            {
                var val = row.Values[c];
                if (val.HasValue)
                    _adapter.WriteNumber(_nextRow, c + 2, (double)val.Value, null);
            }
            _nextRow++;
        }
    }
}
