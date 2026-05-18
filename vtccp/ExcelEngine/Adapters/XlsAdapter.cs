namespace ExcelEngine.Adapters;

using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System.Diagnostics;

/// <summary>
/// IExcelAdapter implementation using NPOI HSSF for .xls (BIFF8) output.
/// XLS limits: 65,536 rows, 256 columns. Our ~120 columns fit.
/// A warning is logged when data rows exceed 60,000 (approaching the limit).
/// </summary>
public sealed class XlsAdapter : IExcelAdapter
{
    private HSSFWorkbook? _wb;
    private ISheet? _ws;
    private string _filePath = string.Empty;

    // Row-level style accumulator: maps 1-based row index → the shared ICellStyle for that row.
    // SetRowBold / ClearRowFill / SetRowWrapText all modify the SAME style object for a given row
    // and re-apply it, so the entire row shares one HSSF style entry instead of N (one per column).
    // This is critical for HSSF: the XLS format caps a workbook at ~4,000 unique cell styles;
    // applying a new style per cell per header rewrite exhausts that limit quickly.
    private readonly Dictionary<int, ICellStyle> _rowStyleMap = [];
    private readonly Dictionary<string, ICellStyle> _styleCache = [];
    private IFont? _boldFont;

    public int MaxDataRows => 65_536;

    public bool OpenOrCreate(string filePath)
    {
        _filePath = filePath;
        if (File.Exists(filePath))
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            _wb = new HSSFWorkbook(fs);
            return true;
        }
        _wb = new HSSFWorkbook();
        return false;
    }

    // Initialises workbook-scoped font/style objects that are reused for the workbook lifetime.
    // Called once after _wb is created or loaded. Safe to call multiple times (idempotent).
    private void EnsureWorkbookStyles()
    {
        if (_boldFont is not null) return;
        _boldFont = _wb!.CreateFont();
        _boldFont.IsBold = true;
    }

    public int EnsureSheet(string sheetName)
    {
        _ws = _wb!.GetSheet(sheetName) ?? _wb.CreateSheet(sheetName);
        int rowCount = _ws.LastRowNum + 1;
        if (rowCount == 1 && _ws.GetRow(0) == null)
            rowCount = 0;

        EnsureWorkbookStyles();

        if (rowCount > 60_000)
            Debug.WriteLine($"[VTCCP] XLS warning: {rowCount} rows in '{sheetName}'. Max is 65,536.");

        return rowCount;
    }

    public void WriteString(int row, int col, string? value)
    {
        if (value is null) return;
        var r = GetOrCreateRow(row - 1);
        var cell = r.GetCell(col - 1) ?? r.CreateCell(col - 1);
        cell.SetCellValue(value);
    }

    public void WriteNumber(int row, int col, double value, string? numberFormat = null)
    {
        var r = GetOrCreateRow(row - 1);
        var cell = r.GetCell(col - 1) ?? r.CreateCell(col - 1, CellType.Numeric);
        cell.SetCellValue(value);
        if (numberFormat is not null)
        {
            var style = GetOrCreateFormatStyle(numberFormat);
            cell.CellStyle = style;
        }
    }

    public void WriteDateTime(int row, int col, DateTime value, string? numberFormat = null)
    {
        var r = GetOrCreateRow(row - 1);
        var cell = r.GetCell(col - 1) ?? r.CreateCell(col - 1);
        var fmt = numberFormat ?? "yyyy-mm-dd";
        var style = GetOrCreateFormatStyle(fmt);
        cell.SetCellValue(value);
        cell.CellStyle = style;
    }

    public void SetRowBold(int row, int colCount)
    {
        var style = GetOrBuildRowStyle(row);
        style.SetFont(_boldFont!);
        ApplyStyleToRow(row, colCount, style);
    }

    public void SetColumnWidth(int col, double width)
    {
        // NPOI uses 1/256 of a character width
        _ws!.SetColumnWidth(col - 1, (int)(width * 256));
    }

    public void ClearRowFill(int row, int colCount)
    {
        var style = GetOrBuildRowStyle(row);
        style.FillPattern = FillPattern.NoFill;
        ApplyStyleToRow(row, colCount, style);
    }

    public void SetRowBackground(int row, int colCount, uint argbColor)
    {
        var r = GetOrCreateRow(row - 1);
        byte red = (byte)((argbColor >> 16) & 0xFF);
        byte green = (byte)((argbColor >> 8) & 0xFF);
        byte blue = (byte)(argbColor & 0xFF);
        var hssf = (HSSFWorkbook)_wb!;
        var palette = hssf.GetCustomPalette();

        short colorIndex = HSSFColor.Coral.Index;
        try
        {
            palette.SetColorAtIndex(colorIndex, red, green, blue);
        }
        catch
        {
            // Palette full — fall back to a built-in near-blue
            colorIndex = HSSFColor.CornflowerBlue.Index;
        }

        for (int c = 0; c < colCount; c++)
        {
            var cell = r.GetCell(c) ?? r.CreateCell(c);
            var style = _wb!.CreateCellStyle();
            style.FillForegroundColor = colorIndex;
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;
        }
    }

    public void SetRowHeight(int row, double heightPoints)
    {
        var r = GetOrCreateRow(row - 1);
        r.HeightInPoints = (float)heightPoints;
    }

    public void SetRowWrapText(int row, int colCount)
    {
        var style = GetOrBuildRowStyle(row);
        style.WrapText = true;
        ApplyStyleToRow(row, colCount, style);
    }

    public void SetCellBold(int row, int col)
    {
        // Uses a shared workbook-scoped bold style (cached in _rowStyleMap at the sentinel key -1).
        // Individual label cells (e.g. in ImagesSheetWriter) don't need per-cell cloning because
        // WriteString leaves them with the default style; bold is applied first, then background.
        if (!_rowStyleMap.TryGetValue(-1, out var boldStyle))
        {
            boldStyle = _wb!.CreateCellStyle();
            boldStyle.SetFont(_boldFont!);
            _rowStyleMap[-1] = boldStyle;
        }
        var r = GetOrCreateRow(row - 1);
        var cell = r.GetCell(col - 1) ?? r.CreateCell(col - 1);
        cell.CellStyle = boldStyle;
    }

    public void SetCellBackground(int row, int col, uint argbColor)
    {
        var r = GetOrCreateRow(row - 1);
        var cell = r.GetCell(col - 1) ?? r.CreateCell(col - 1);
        byte red = (byte)((argbColor >> 16) & 0xFF);
        byte green = (byte)((argbColor >> 8) & 0xFF);
        byte blue = (byte)(argbColor & 0xFF);
        var hssf = (HSSFWorkbook)_wb!;
        var palette = hssf.GetCustomPalette();
        short colorIndex = HSSFColor.LightBlue.Index;
        try
        {
            palette.SetColorAtIndex(colorIndex, red, green, blue);
        }
        catch
        {
            colorIndex = HSSFColor.LightBlue.Index;
        }
        // Clone existing cell style so bold/font settings from SetCellBold are preserved.
        var style = _wb!.CreateCellStyle();
        style.CloneStyleFrom(cell.CellStyle ?? _wb.CreateCellStyle());
        style.FillForegroundColor = colorIndex;
        style.FillPattern = FillPattern.SolidForeground;
        cell.CellStyle = style;
    }

    public void WriteEmbeddedImage(int row, int col, byte[] jpegBytes)
    {
        if (_wb is null || _ws is null) return;

        int picIdx = _wb.AddPicture(jpegBytes, PictureType.JPEG);

        // Get the existing drawing patriarch for this sheet, or create one.
        // NPOI throws if CreateDrawingPatriarch() is called when one already exists.
        var patriarch = (_ws.DrawingPatriarch as HSSFPatriarch)
                        ?? (HSSFPatriarch)_ws.CreateDrawingPatriarch();

        // Anchor: (dx1, dy1, dx2, dy2, col1, row1, col2, row2) — col/row are 0-based.
        // Span approximately 3 columns × 6 rows to hold a ~220×220px image.
        var anchor = new HSSFClientAnchor(0, 0, 1023, 255,
            col - 1, row - 1, col + 2, row + 5);
        anchor.AnchorType = AnchorType.MoveAndResize;
        patriarch.CreatePicture(anchor, picIdx);
    }

    public void Save()
    {
        using var fs = new FileStream(_filePath, FileMode.Create, FileAccess.Write);
        _wb!.Write(fs);
    }

    public void SaveToPath(string path)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        _wb!.Write(fs);
    }

    public void Dispose()
    {
        _wb?.Close();
    }

    // Returns the shared ICellStyle for the given 1-based row, creating it on first access.
    // All three row-level formatting operations (bold / wrap / clear-fill) mutate this same
    // object so they accumulate rather than overwrite — and so the row shares one style entry
    // across all its cells instead of N entries.
    private ICellStyle GetOrBuildRowStyle(int row)
    {
        if (_rowStyleMap.TryGetValue(row, out var s)) return s;
        s = _wb!.CreateCellStyle();
        _rowStyleMap[row] = s;
        return s;
    }

    // Assigns the given style to every cell in the 1-based row [col 0 .. colCount-1].
    private void ApplyStyleToRow(int row, int colCount, ICellStyle style)
    {
        var r = GetOrCreateRow(row - 1);
        for (int c = 0; c < colCount; c++)
        {
            var cell = r.GetCell(c) ?? r.CreateCell(c);
            cell.CellStyle = style;
        }
    }

    private IRow GetOrCreateRow(int zeroBasedRow)
        => _ws!.GetRow(zeroBasedRow) ?? _ws.CreateRow(zeroBasedRow);

    private ICellStyle GetOrCreateFormatStyle(string format)
    {
        if (_styleCache.TryGetValue(format, out var cached)) return cached;
        var style = _wb!.CreateCellStyle();
        var fmt = _wb.CreateDataFormat();
        style.DataFormat = fmt.GetFormat(format);
        _styleCache[format] = style;
        return style;
    }
}
