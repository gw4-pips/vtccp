namespace ExcelEngine.Adapters;

/// <summary>
/// Abstracts format-specific Excel I/O so the engine core (ExcelWriter) is format-agnostic.
/// Two implementations: XlsxAdapter (EPPlus / .xlsx) and XlsAdapter (NPOI HSSF / .xls).
/// </summary>
public interface IExcelAdapter : IDisposable
{
    /// <summary>
    /// Open an existing file for append, or create a new one.
    /// Returns true if the file already existed.
    /// </summary>
    bool OpenOrCreate(string filePath);

    /// <summary>
    /// Ensure the named worksheet exists and is active.
    /// Returns the zero-based row count already present (header rows included).
    /// </summary>
    int EnsureSheet(string sheetName);

    /// <summary>Write a string value to the given 1-based row/col. </summary>
    void WriteString(int row, int col, string? value);

    /// <summary>Write a numeric value (double) with an optional Excel number format string.</summary>
    void WriteNumber(int row, int col, double value, string? numberFormat = null);

    /// <summary>Write a DateTime value with an optional Excel number format string.</summary>
    void WriteDateTime(int row, int col, DateTime value, string? numberFormat = null);

    /// <summary>Apply bold formatting to every cell in the given row.</summary>
    void SetRowBold(int row, int colCount);

    /// <summary>Set the column width in character units (approximate).</summary>
    void SetColumnWidth(int col, double width);

    /// <summary>Apply background fill colour to a row (used for header rows).</summary>
    void SetRowBackground(int row, int colCount, uint argbColor);

    /// <summary>Set the row height in points (e.g. 30.0 for a double-height row).</summary>
    void SetRowHeight(int row, double heightPoints);

    /// <summary>Enable word-wrap on every cell in the given row.</summary>
    void SetRowWrapText(int row, int colCount);

    /// <summary>Apply bold formatting to a single cell.</summary>
    void SetCellBold(int row, int col);

    /// <summary>Apply background fill colour to a single cell.</summary>
    void SetCellBackground(int row, int col, uint argbColor);

    /// <summary>Save the workbook to the path provided in OpenOrCreate.</summary>
    void Save();

    /// <summary>Maximum data rows this format supports before the file must be rotated.</summary>
    int MaxDataRows { get; }
}
