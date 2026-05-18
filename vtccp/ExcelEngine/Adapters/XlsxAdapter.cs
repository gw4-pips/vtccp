namespace ExcelEngine.Adapters;

using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Drawing;

/// <summary>
/// IExcelAdapter implementation using EPPlus for .xlsx output.
/// </summary>
public sealed class XlsxAdapter : IExcelAdapter
{
    private ExcelPackage? _pkg;
    private ExcelWorksheet? _ws;
    private string _filePath = string.Empty;
    private int _imgSeq = 0;

    public int MaxDataRows => 1_000_000;

    static XlsxAdapter()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public bool OpenOrCreate(string filePath)
    {
        _filePath = filePath;
        var fi = new FileInfo(filePath);
        bool existed = fi.Exists;
        _pkg = existed ? new ExcelPackage(fi) : new ExcelPackage();
        return existed;
    }

    public int EnsureSheet(string sheetName)
    {
        _ws = _pkg!.Workbook.Worksheets[sheetName]
              ?? _pkg.Workbook.Worksheets.Add(sheetName);
        return Math.Max(0, _ws.Dimension?.Rows ?? 0);
    }

    public void WriteString(int row, int col, string? value)
    {
        if (value is not null)
            _ws!.Cells[row, col].Value = value;
    }

    public void WriteNumber(int row, int col, double value, string? numberFormat = null)
    {
        var cell = _ws!.Cells[row, col];
        cell.Value = value;
        if (numberFormat is not null)
            cell.Style.Numberformat.Format = numberFormat;
    }

    public void WriteDateTime(int row, int col, DateTime value, string? numberFormat = null)
    {
        var cell = _ws!.Cells[row, col];
        cell.Value = value;
        cell.Style.Numberformat.Format = numberFormat ?? "yyyy-mm-dd";
    }

    public void SetRowBold(int row, int colCount)
    {
        _ws!.Cells[row, 1, row, colCount].Style.Font.Bold = true;
    }

    public void SetColumnWidth(int col, double width)
    {
        _ws!.Column(col).Width = width;
    }

    public void ClearRowFill(int row, int colCount)
    {
        _ws!.Cells[row, 1, row, colCount].Style.Fill.PatternType = ExcelFillStyle.None;
    }

    public void SetRowBackground(int row, int colCount, uint argbColor)
    {
        var cells = _ws!.Cells[row, 1, row, colCount];
        cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
        var r = (byte)((argbColor >> 16) & 0xFF);
        var g = (byte)((argbColor >> 8) & 0xFF);
        var b = (byte)(argbColor & 0xFF);
        cells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, r, g, b));
    }

    public void SetRowHeight(int row, double heightPoints)
    {
        _ws!.Row(row).Height = heightPoints;
        _ws.Row(row).CustomHeight = true;
    }

    public void SetRowWrapText(int row, int colCount)
    {
        _ws!.Cells[row, 1, row, colCount].Style.WrapText = true;
    }

    public void SetCellBold(int row, int col)
    {
        _ws!.Cells[row, col].Style.Font.Bold = true;
    }

    public void SetCellBackground(int row, int col, uint argbColor)
    {
        var cell = _ws!.Cells[row, col];
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        var r = (byte)((argbColor >> 16) & 0xFF);
        var g = (byte)((argbColor >> 8) & 0xFF);
        var b = (byte)(argbColor & 0xFF);
        cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, r, g, b));
    }

    public void WriteEmbeddedImage(int row, int col, byte[] jpegBytes)
    {
        using var ms = new MemoryStream(jpegBytes);
        // Name must be unique within the workbook across all sheets.
        string name = $"img_r{row}_c{col}_{++_imgSeq}";
        var pic = _ws!.Drawings.AddPicture(name, ms, ePictureType.Jpg);
        // SetPosition takes 0-based row/col + pixel offsets from the cell edge.
        pic.SetPosition(row - 1, 2, col - 1, 2);
        // SetSize: width × height in pixels. Square to accommodate both QR and DM aspect ratios.
        pic.SetSize(220, 220);
    }

    public void Save()
    {
        _pkg!.SaveAs(new FileInfo(_filePath));
    }

    public void SaveToPath(string path)
    {
        _pkg!.SaveAs(new FileInfo(path));
    }

    public void Dispose()
    {
        _pkg?.Dispose();
    }
}
