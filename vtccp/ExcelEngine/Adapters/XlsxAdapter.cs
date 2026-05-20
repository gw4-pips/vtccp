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
        string name = $"img_r{row}_c{col}_{++_imgSeq}";
        // EPPlus 7.x: AddPicture(name, stream) auto-detects format from stream header.
        var pic = _ws!.Drawings.AddPicture(name, ms);
        pic.SetPosition(row - 1, 2, col - 1, 2);
        pic.SetSize(220, 220);
    }

    public void WriteLogoImage(int row, byte[] imageBytes, string fileExtension)
    {
        if (imageBytes is null || imageBytes.Length == 0 || _ws is null) return;

        using var ms = new MemoryStream(imageBytes);
        string name = $"logo_r{row}_{++_imgSeq}";
        // EPPlus 7.x: format auto-detected from stream header (PNG magic, JFIF SOI, etc.).
        var pic = _ws.Drawings.AddPicture(name, ms);

        // Anchor at column 8 (0-based index 8 = column I), 4 px from the cell edge.
        // Column I is visible without horizontal scrolling on typical monitors and
        // clears the title-text cells without overlapping them.
        pic.SetPosition(row - 1, 4, 8, 4);

        // 160 × 54 px: fits a landscape banner logo inside the tall title row (54pt ≈ 72px).
        pic.SetSize(160, 54);
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
