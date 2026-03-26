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
///   Row 1: Title row — "VCCS DMV TruCheck Command Pilot vX.X — Job: {job}" (merged conceptually)
///   Row 2: Column header labels
///   Row 3+: Data rows (one per VerificationRecord)
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

    private int _nextDataRow;
    private bool _headersWritten;

    // Background colour for header row (Webscan uses a mid-blue; approximate with RGB 68 114 196)
    private const uint HeaderBgArgb = 0x4472C4;

    public ExcelWriter(IExcelAdapter adapter, ColumnSchema schema, SessionState session, string sheetName = "Main")
    {
        _adapter = adapter;
        _schema = schema;
        _session = session;
        _sheetName = sheetName;
    }

    /// <summary>
    /// Open or create the output file and initialise the writer.
    /// Must be called before AppendRecord.
    /// </summary>
    public void Open(string filePath)
    {
        bool existed = _adapter.OpenOrCreate(filePath);
        int existingRows = _adapter.EnsureSheet(_sheetName);

        if (existed && existingRows >= FirstDataRow)
        {
            // Append mode — headers already written
            _headersWritten = true;
            _nextDataRow = existingRows + 1;
        }
        else
        {
            // New file or empty sheet
            _headersWritten = false;
            _nextDataRow = FirstDataRow;
        }

        // Always set column widths (idempotent for appending)
        ApplyColumnWidths();
    }

    /// <summary>
    /// Write a single VerificationRecord as the next data row.
    /// Writes header rows first if this is a new file.
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

        var values = record.SymbologyFamily == SymbologyFamily.Linear1D
            ? throw new NotSupportedException("Use the 1D mapper for ISO 15416 records (Task 3).")
            : DataMatrix2DMapper.Map(record, _schema);

        WriteDataRow(_nextDataRow, values);

        // GS1 Data Format Check is recorded as extra cells on the same row
        // by convention — the Webscan layout places the DFC data in dedicated columns
        // (not appended as extra rows), so nothing extra to write here.

        _nextDataRow++;
        _adapter.CurrentDataRowCount.ToString(); // force property read (no-op, just reference)
    }

    /// <summary>Save and close the underlying file.</summary>
    public void Save() => _adapter.Save();

    public void Dispose() => _adapter.Dispose();

    // ── Private helpers ───────────────────────────────────────────────────────

    private void WriteTitleRow()
    {
        var title = $"VCCS DMV TruCheck Command Pilot — Job: {_session.JobName ?? "(no job)"}" +
                    $" | Operator: {_session.OperatorId ?? "-"}" +
                    $" | Roll: {_session.RollNumber}";
        _adapter.WriteString(TitleRow, 1, title);
        _adapter.SetRowBold(TitleRow, _schema.Columns.Count);
    }

    private void WriteHeaderRow()
    {
        for (int i = 0; i < _schema.Columns.Count; i++)
        {
            _adapter.WriteString(HeaderRow, i + 1, _schema.Columns[i].DisplayName);
        }
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
                // Try to parse as number for numeric-looking string fields
                // (e.g. grade values that come back as strings from certain paths)
                if (col.NumberFormat is not null && double.TryParse(s, out double parsed))
                    _adapter.WriteNumber(rowNum, colNum, parsed, col.NumberFormat);
                else
                    _adapter.WriteString(rowNum, colNum, s);
            }
        }
    }

    private void ApplyColumnWidths()
    {
        for (int i = 0; i < _schema.Columns.Count; i++)
        {
            _adapter.SetColumnWidth(i + 1, _schema.Columns[i].Width);
        }
    }

    private void CheckRowLimit()
    {
        int dataRows = _nextDataRow - FirstDataRow;
        if (dataRows >= _adapter.MaxDataRows - 100)
        {
            throw new InvalidOperationException(
                $"Output file is approaching the {_adapter.MaxDataRows:N0}-row limit " +
                $"({dataRows:N0} data rows written). Start a new job file.");
        }
    }
}
