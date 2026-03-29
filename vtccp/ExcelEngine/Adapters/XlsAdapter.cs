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

    private readonly Dictionary<string, ICellStyle> _styleCache = [];
    private ICellStyle? _boldStyle;
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

    public int EnsureSheet(string sheetName)
    {
        _ws = _wb!.GetSheet(sheetName) ?? _wb.CreateSheet(sheetName);
        int rowCount = _ws.LastRowNum + 1;
        if (rowCount == 1 && _ws.GetRow(0) == null)
            rowCount = 0;

        _boldFont = _wb.CreateFont();
        _boldFont.IsBold = true;
        _boldStyle = _wb.CreateCellStyle();
        _boldStyle.SetFont(_boldFont);

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
        var r = GetOrCreateRow(row - 1);
        for (int c = 0; c < colCount; c++)
        {
            var cell = r.GetCell(c) ?? r.CreateCell(c);
            var style = _wb!.CreateCellStyle();
            style.CloneStyleFrom(cell.CellStyle ?? _wb.CreateCellStyle());
            style.SetFont(_boldFont!);
            cell.CellStyle = style;
        }
    }

    public void SetColumnWidth(int col, double width)
    {
        // NPOI uses 1/256 of a character width
        _ws!.SetColumnWidth(col - 1, (int)(width * 256));
    }

    public void ClearRowFill(int row, int colCount)
    {
        var r = GetOrCreateRow(row - 1);
        for (int c = 0; c < colCount; c++)
        {
            var cell = r.GetCell(c) ?? r.CreateCell(c);
            var style = _wb!.CreateCellStyle();
            style.FillPattern = FillPattern.NoFill;
            cell.CellStyle = style;
        }
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
        var r = GetOrCreateRow(row - 1);
        for (int c = 0; c < colCount; c++)
        {
            var cell = r.GetCell(c) ?? r.CreateCell(c);
            var style = _wb!.CreateCellStyle();
            style.CloneStyleFrom(cell.CellStyle ?? _wb.CreateCellStyle());
            style.WrapText = true;
            cell.CellStyle = style;
        }
    }

    public void SetCellBold(int row, int col)
    {
        var r = GetOrCreateRow(row - 1);
        var cell = r.GetCell(col - 1) ?? r.CreateCell(col - 1);
        var style = _wb!.CreateCellStyle();
        style.CloneStyleFrom(cell.CellStyle ?? _wb.CreateCellStyle());
        style.SetFont(_boldFont!);
        cell.CellStyle = style;
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
