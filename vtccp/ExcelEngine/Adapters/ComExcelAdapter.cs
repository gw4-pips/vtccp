namespace ExcelEngine.Adapters;

using System.Runtime.InteropServices;
using System.Runtime.Versioning;

// Marshal.GetActiveObject was removed in .NET Core / .NET 5+.
// Replicate it via a direct P/Invoke to oleaut32.dll (Windows only).
// CLSID for "Excel.Application" — stable across all Office versions.
// ReSharper disable InconsistentNaming

/// <summary>
/// IExcelAdapter implementation that writes directly into a running Excel process
/// via COM automation, using late binding (<c>dynamic</c>) to avoid a compile-time
/// dependency on Microsoft.Office.Interop.Excel.
///
/// Purpose: when an operator has the output XLSX open in Excel during a scan session,
/// EPPlus's SaveAs() fails with an IOException (file locked).  This adapter bypasses
/// file I/O entirely — rows are written to Excel's in-memory workbook object, and
/// Save() calls Excel's own save mechanism, which never conflicts with itself.
///
/// Usage: call <see cref="TryAttach"/> at session start.  If it returns non-null,
/// pass the adapter to ExcelWriter instead of the default XlsxAdapter.  If it
/// returns null (Excel not running, file not found in open workbooks, or COM
/// unavailable), fall back to XlsxAdapter.
///
/// Threading: all COM calls must occur on an STA thread.  SessionManager.AddRecord
/// is always dispatched through the WPF Dispatcher (STA), so this is guaranteed.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class ComExcelAdapter : IExcelAdapter
{
    private readonly dynamic _workbook;
    private dynamic _worksheet;
    private string  _filePath;
    private int     _imgSeq;

    // MsoTriState integer constants — avoids a compile-time reference to
    // Microsoft.Office.Core, which is not available in the ExcelEngine project.
    private const int MsoFalse = 0;
    private const int MsoCTrue = -1;

    public int MaxDataRows => 1_048_576;

    private ComExcelAdapter(dynamic workbook, string filePath)
    {
        _workbook  = workbook;
        _worksheet = workbook.ActiveSheet;
        _filePath  = filePath;
    }

    // ── Factory ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Attempts to find a running Excel instance that has <paramref name="filePath"/>
    /// open as a workbook.  Returns a ready-to-use adapter, or <c>null</c> if:
    /// <list type="bullet">
    ///   <item>Excel is not running on this machine.</item>
    ///   <item>The file is not open in any Excel workbook.</item>
    ///   <item>COM automation is unavailable.</item>
    /// </list>
    /// </summary>
    [SupportedOSPlatform("windows")]
    public static ComExcelAdapter? TryAttach(string filePath)
    {
        try
        {
            dynamic? xlApp = GetRunningExcel();
            if (xlApp is null) return null;

            string target = Path.GetFullPath(filePath).ToLowerInvariant();
            foreach (dynamic wb in xlApp.Workbooks)
            {
                try
                {
                    string wbPath = Path.GetFullPath((string)wb.FullName).ToLowerInvariant();
                    if (wbPath == target)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[VTCCP-COM] Attached to open workbook: {wb.FullName}");
                        return new ComExcelAdapter(wb, filePath);
                    }
                }
                catch { /* workbook may not have a path (e.g. unsaved) — skip */ }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-COM] Excel is running but '{Path.GetFileName(filePath)}' is not open.");
            return null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-COM] TryAttach failed ({ex.GetType().Name}): {ex.Message}");
            return null;
        }
    }

    [SupportedOSPlatform("windows")]
    private static dynamic? GetRunningExcel()
    {
        // Excel.Application CLSID — stable across all Office versions.
        Guid xlsClsid = new Guid("00024500-0000-0000-C000-000000000046");
        int  hr       = NativeMethods.GetActiveObject(xlsClsid, IntPtr.Zero, out object obj);
        if (hr < 0) return null;   // S_FALSE or error → Excel is not running
        return obj;
    }

    [SupportedOSPlatform("windows")]
    private static class NativeMethods
    {
        [DllImport("oleaut32.dll")]
        internal static extern int GetActiveObject(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
    }

    // ── IExcelAdapter ─────────────────────────────────────────────────────────

    public bool OpenOrCreate(string filePath)
    {
        // The workbook is already open in Excel (found by TryAttach).
        // Record the pinned path for SaveToPath; always report existed=true.
        _filePath = filePath;
        return true;
    }

    public int EnsureSheet(string sheetName)
    {
        dynamic? found = null;
        foreach (dynamic ws in _workbook.Worksheets)
        {
            try
            {
                if (string.Equals((string)ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    found = ws;
                    break;
                }
            }
            catch { /* sheet in bad state — skip */ }
        }

        if (found is null)
        {
            dynamic sheets = _workbook.Worksheets;
            found = sheets.Add(After: sheets[sheets.Count]);
            found.Name = sheetName;
        }

        _worksheet = found;

        try
        {
            dynamic usedRange = _worksheet.UsedRange;
            return (int)usedRange.Rows.Count;
        }
        catch { return 0; }
    }

    public void WriteString(int row, int col, string? value)
    {
        if (value is null) return;
        _worksheet.Cells[row, col].Value = value;
    }

    public void WriteNumber(int row, int col, double value, string? numberFormat = null)
    {
        _worksheet.Cells[row, col].Value = value;
        if (numberFormat is not null)
            _worksheet.Cells[row, col].NumberFormat = numberFormat;
    }

    public void WriteDateTime(int row, int col, DateTime value, string? numberFormat = null)
    {
        // Excel stores DateTime as an OLE Automation date (double).
        _worksheet.Cells[row, col].Value = value.ToOADate();
        _worksheet.Cells[row, col].NumberFormat = numberFormat ?? "yyyy-mm-dd";
    }

    public void SetRowBold(int row, int colCount)
    {
        _worksheet.Range[_worksheet.Cells[row, 1], _worksheet.Cells[row, colCount]].Font.Bold = true;
    }

    public void SetColumnWidth(int col, double width)
    {
        _worksheet.Columns[col].ColumnWidth = width;
    }

    public void ClearRowFill(int row, int colCount)
    {
        // -4142 = xlColorIndexNone (no fill).
        _worksheet.Range[_worksheet.Cells[row, 1], _worksheet.Cells[row, colCount]]
            .Interior.ColorIndex = -4142;
    }

    public void SetRowBackground(int row, int colCount, uint argbColor)
    {
        _worksheet.Range[_worksheet.Cells[row, 1], _worksheet.Cells[row, colCount]]
            .Interior.Color = ArgbToOleColor(argbColor);
    }

    public void SetRowHeight(int row, double heightPoints)
    {
        _worksheet.Rows[row].RowHeight = heightPoints;
    }

    public void SetRowWrapText(int row, int colCount)
    {
        _worksheet.Range[_worksheet.Cells[row, 1], _worksheet.Cells[row, colCount]]
            .WrapText = true;
    }

    public void SetCellBold(int row, int col)
    {
        _worksheet.Cells[row, col].Font.Bold = true;
    }

    public void SetCellBackground(int row, int col, uint argbColor)
    {
        _worksheet.Cells[row, col].Interior.Color = ArgbToOleColor(argbColor);
    }

    public void Save()
    {
        // Save through Excel directly — no file-lock conflict with EPPlus SaveAs().
        _workbook.Save();
        System.Diagnostics.Debug.WriteLine("[VTCCP-COM] Workbook saved via COM.");
    }

    public void SaveToPath(string path)
    {
        // FileFormat 51 = xlOpenXMLWorkbook (.xlsx, no macro).
        _workbook.SaveAs(path, 51);
        System.Diagnostics.Debug.WriteLine($"[VTCCP-COM] Workbook saved via COM to: {path}");
    }

    public void WriteEmbeddedImage(int row, int col, byte[] jpegBytes)
    {
        // COM AddPicture requires the image to be on disk.  Write to a temp file,
        // insert it, then delete the temp file.
        string tmp = Path.Combine(Path.GetTempPath(), $"vtccp_img_{++_imgSeq}.jpg");
        try
        {
            File.WriteAllBytes(tmp, jpegBytes);

            dynamic cell = _worksheet.Cells[row, col];
            double  left = (double)cell.Left + 2;
            double  top  = (double)cell.Top  + 2;

            // AddPicture(filename, linkToFile, saveWithDocument, left, top, width, height)
            _worksheet.Shapes.AddPicture(tmp, MsoFalse, MsoCTrue, left, top, 220.0, 220.0);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-COM] WriteEmbeddedImage failed: {ex.Message}");
        }
        finally
        {
            try { File.Delete(tmp); } catch { /* ignore */ }
        }
    }

    public void WriteLogoImage(int row, byte[] imageBytes, string fileExtension)
    {
        if (imageBytes is null || imageBytes.Length == 0) return;

        string ext = fileExtension.TrimStart('.').ToLowerInvariant();
        string tmp = Path.Combine(Path.GetTempPath(), $"vtccp_logo_{++_imgSeq}.{ext}");
        try
        {
            File.WriteAllBytes(tmp, imageBytes);

            // Anchor at column I (index 9) — same as XlsxAdapter.
            dynamic cell = _worksheet.Cells[row, 9];
            double  left = (double)cell.Left + 4;
            double  top  = (double)cell.Top  + 4;

            _worksheet.Shapes.AddPicture(tmp, MsoFalse, MsoCTrue, left, top, 160.0, 54.0);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-COM] WriteLogoImage failed: {ex.Message}");
        }
        finally
        {
            try { File.Delete(tmp); } catch { /* ignore */ }
        }
    }

    public void Dispose()
    {
        // Do NOT call _workbook.Close() — the operator wants it to stay open.
        // Do NOT release the COM RCW — Excel is still running and owns the object.
        // The GC will release the RCW naturally when this adapter is collected.
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Converts an ARGB colour (uint, as used throughout IExcelAdapter) to an OLE BGR
    /// colour integer (as expected by Excel COM <c>Interior.Color</c>).
    /// ARGB: [31:24]=A, [23:16]=R, [15:8]=G, [7:0]=B.
    /// OLE:  integer whose bytes are B, G, R (R and B are swapped vs ARGB).
    /// </summary>
    private static int ArgbToOleColor(uint argb)
    {
        int r = (int)((argb >> 16) & 0xFF);
        int g = (int)((argb >> 8)  & 0xFF);
        int b = (int)(argb          & 0xFF);
        return (b << 16) | (g << 8) | r;
    }
}
