namespace ExcelEngine.Writer;

using ExcelEngine.Adapters;
using ExcelEngine.Models;
using ExcelEngine.Schema;

/// <summary>
/// Core format-agnostic writer. Works with any IExcelAdapter.
/// Manages header rows, data rows, column widths, and basic formatting
/// to match the Webscan TruCheck "Main" worksheet layout.
///
/// Row layout:
///   Row 1: Title row — "VCCS DMV TruCheck Command Pilot — Job: {job} | ..."
///   Row 2: Column header labels (bold, blue background)
///   Row 3+: Data rows (one per VerificationRecord)
///
/// GS1 Data Format Check fields are written as dedicated columns
/// in the same data row (DFC_Standard, DFC_R1_Name/Data/Check ... DFC_R8_*).
/// </summary>
public sealed class ExcelWriter : IDisposable
{
    private const int TitleRow = 1;
    private const int HeaderRow = 2;
    private const int FirstDataRow = 3;

    private readonly IExcelAdapter _adapter;
    private readonly ColumnSchema _schema;
    private readonly SessionState _session;
    private readonly string _sheetName;
    private readonly ElementWidthsWriter _ewWriter;
    private readonly PerScanTableWriter _perScanWriter;

    private int _nextDataRow;
    private int _dataRowCount;
    private bool _headersWritten;

    // Header row background colour — Webscan uses mid-blue (RGB 68 114 196)
    private const uint HeaderBgArgb = 0x4472C4;

    /// <summary>Number of data rows written so far (not counting title/header rows).</summary>
    public int DataRowCount => _dataRowCount;

    public ExcelWriter(IExcelAdapter adapter, ColumnSchema schema, SessionState session, string sheetName = "Main")
    {
        _adapter = adapter;
        _schema = schema;
        _session = session;
        _sheetName = sheetName;
        _ewWriter = new ElementWidthsWriter(adapter);
        _perScanWriter = new PerScanTableWriter(adapter, schema);
    }

    /// <summary>
    /// Open or create the output file and initialise the writer.
    /// Must be called before AppendRecord.
    /// Throws <see cref="IOException"/> if the file is locked by another process (e.g. Excel).
    /// </summary>
    public void Open(string filePath)
    {
        var lockError = ExcelFileManager.CheckFileLocked(filePath);
        if (lockError is not null)
            throw new IOException(lockError);

        bool existed = _adapter.OpenOrCreate(filePath);
        int existingRows = _adapter.EnsureSheet(_sheetName);

        if (existed && existingRows >= FirstDataRow)
        {
            _headersWritten = true;
            _nextDataRow = existingRows + 1;
            _dataRowCount = existingRows - (FirstDataRow - 1);
        }
        else
        {
            _headersWritten = false;
            _nextDataRow = FirstDataRow;
            _dataRowCount = 0;
        }

        ApplyColumnWidths();
    }

    /// <summary>
    /// Write a single VerificationRecord as the next data row.
    /// Writes title + header rows first if this is a new file.
    /// Supports both 1D (ISO 15416) and 2D (ISO 15415 / Data Matrix) records in the same file.
    /// For 1D records with ElementWidths data, also writes to the "Element Widths" sheet.
    /// </summary>
    public void AppendRecord(VerificationRecord record)
    {
        CheckRowLimit();

        if (!_headersWritten)
        {
            WriteTitleRow();
            WriteHeaderRow();
            _headersWritten = true;
        }

        Dictionary<string, object?> values;
        if (record.SymbologyFamily == SymbologyFamily.Linear1D)
            values = ISO15416Mapper.Map(record, _schema);
        else
            values = DataMatrix2DMapper.Map(record, _schema);

        WriteDataRow(_nextDataRow, values);

        if (record.DataFormatCheck is not null)
            WriteDfcColumns(_nextDataRow, record.DataFormatCheck);

        // Advance past the summary row.
        _nextDataRow++;
        _dataRowCount++;

        // For 1D records: write per-scan sub-table rows immediately below the summary row.
        // Per-scan rows are auxiliary (not counted in _dataRowCount), but _nextDataRow and
        // the effective row-limit check must account for them to prevent .xls overflow.
        if (record.SymbologyFamily == SymbologyFamily.Linear1D && record.ScanResults.Count > 0)
        {
            int scanRowsWritten = _perScanWriter.WriteScans(record.ScanResults, _nextDataRow);
            _nextDataRow += scanRowsWritten;
            // Re-check limit after adding scan rows so .xls near-limit warning is accurate.
            CheckRowLimit();
        }

        // For 1D records with element width data, write to "Element Widths" sheet
        // then restore the Main sheet as the active adapter target.
        if (record.SymbologyFamily == SymbologyFamily.Linear1D && record.ElementWidths is not null)
        {
            _ewWriter.WriteRecord(record.ElementWidths);
            _adapter.EnsureSheet(_sheetName);  // restore Main as active sheet
        }
    }

    /// <summary>Save and close the underlying file.</summary>
    public void Save() => _adapter.Save();

    public void Dispose() => _adapter.Dispose();

    // ── Private helpers ───────────────────────────────────────────────────────

    private void WriteTitleRow()
    {
        var title = $"VCCS DMV TruCheck Command Pilot" +
                    $" | Job: {_session.JobName ?? "(no job)"}" +
                    $" | Operator: {_session.OperatorId ?? "-"}" +
                    $" | Roll: {_session.RollNumber}";
        _adapter.WriteString(TitleRow, 1, title);
        _adapter.SetRowBold(TitleRow, _schema.Columns.Count);
    }

    private void WriteHeaderRow()
    {
        for (int i = 0; i < _schema.Columns.Count; i++)
            _adapter.WriteString(HeaderRow, i + 1, _schema.Columns[i].DisplayName);

        _adapter.SetRowBold(HeaderRow, _schema.Columns.Count);
        _adapter.SetRowBackground(HeaderRow, _schema.Columns.Count, HeaderBgArgb);
    }

    private void WriteDataRow(int rowNum, Dictionary<string, object?> values)
    {
        for (int i = 0; i < _schema.Columns.Count; i++)
        {
            var col = _schema.Columns[i];
            if (!values.TryGetValue(col.FieldId, out var val) || val is null)
                continue;

            int colNum = i + 1;

            if (val is DateTime dt)
            {
                _adapter.WriteDateTime(rowNum, colNum, dt, col.NumberFormat);
            }
            else if (val is double d)
            {
                _adapter.WriteNumber(rowNum, colNum, d, col.NumberFormat);
            }
            else if (val is string s && !string.IsNullOrEmpty(s))
            {
                if (col.NumberFormat is not null && double.TryParse(s, out double parsed))
                    _adapter.WriteNumber(rowNum, colNum, parsed, col.NumberFormat);
                else
                    _adapter.WriteString(rowNum, colNum, s);
            }
        }
    }

    /// <summary>
    /// Write the GS1 Data Format Check block into the dedicated DFC columns of the data row.
    /// Columns DFC_Standard, DFC_R1_Name/Data/Check … DFC_R8_Name/Data/Check are part of the schema.
    /// </summary>
    private void WriteDfcColumns(int rowNum, DataFormatCheckResult dfc)
    {
        WriteSchemaCell(rowNum, "DFC_Standard",
            dfc.Standard is not null
                ? $"{dfc.Standard}: {dfc.Overall switch { OverallPassFail.Pass => "PASS", OverallPassFail.Fail => "FAIL", _ => "" }}"
                : null);

        for (int i = 0; i < dfc.Rows.Count && i < 8; i++)
        {
            int slot = i + 1;
            var row = dfc.Rows[i];
            WriteSchemaCell(rowNum, $"DFC_R{slot}_Name",  row.Name);
            WriteSchemaCell(rowNum, $"DFC_R{slot}_Data",  row.Data);
            WriteSchemaCell(rowNum, $"DFC_R{slot}_Check", row.Check);
        }
    }

    private void WriteSchemaCell(int rowNum, string fieldId, string? value)
    {
        if (value is null) return;
        var colIdx = _schema.GetColumnIndex(fieldId);
        if (colIdx.HasValue)
            _adapter.WriteString(rowNum, colIdx.Value, value);
    }

    private void ApplyColumnWidths()
    {
        for (int i = 0; i < _schema.Columns.Count; i++)
            _adapter.SetColumnWidth(i + 1, _schema.Columns[i].Width);
    }

    private void CheckRowLimit()
    {
        // Use _nextDataRow (actual physical next row) so per-scan auxiliary rows are counted.
        int physicalUsed = _nextDataRow - 1;  // rows consumed so far (0-based count)

        if (physicalUsed >= _adapter.MaxDataRows - 100)
        {
            throw new InvalidOperationException(
                $"Output file is approaching the {_adapter.MaxDataRows:N0}-row limit " +
                $"({physicalUsed:N0} physical rows written). Start a new job file.");
        }

        // XLS near-limit runtime warning (60,000 threshold)
        if (_adapter.MaxDataRows <= 65_536 && physicalUsed == 60_000)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP] XLS warning: 60,000 physical rows written. " +
                $"Maximum is {_adapter.MaxDataRows:N0}. Consider starting a new job file soon.");
        }
    }
}
