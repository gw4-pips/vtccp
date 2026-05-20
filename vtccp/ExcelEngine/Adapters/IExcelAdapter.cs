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

    /// <summary>Remove background fill from every cell in the given row (no-fill / transparent).</summary>
    void ClearRowFill(int row, int colCount);

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

    /// <summary>Save the workbook to an explicit path (used for rescue copies when the primary path is locked).</summary>
    void SaveToPath(string path);

    /// <summary>
    /// Embed a JPEG image anchored at the top-left corner of the given 1-based row/col.
    /// The image is sized to fit within a ~220×220px bounding box (scan / barcode images).
    /// Callers should pre-set the target row height to ~160pt via SetRowHeight before calling.
    /// </summary>
    void WriteEmbeddedImage(int row, int col, byte[] jpegBytes);

    /// <summary>
    /// Embed a company logo in the title row.
    /// The image is placed as a floating overlay anchored to the top-right area of
    /// <paramref name="row"/> and sized to ~160 × 54 px (landscape banner).
    /// Supports PNG, JPG/JPEG, BMP and GIF — detected from <paramref name="fileExtension"/>.
    /// No-op if <paramref name="imageBytes"/> is null or empty.
    /// </summary>
    void WriteLogoImage(int row, byte[] imageBytes, string fileExtension);

    /// <summary>Maximum data rows this format supports before the file must be rotated.</summary>
    int MaxDataRows { get; }
}
