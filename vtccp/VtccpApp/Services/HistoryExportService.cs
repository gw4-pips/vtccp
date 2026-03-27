namespace VtccpApp.Services;

using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using VtccpApp.Models;

/// <summary>
/// Exports the current Results History to an XLSX workbook using EPPlus.
///
/// Output structure:
///   Sheet 1 "Summary"      — export header, session info, pass/fail stats, grade distribution.
///   Sheet 2 "Scan Records" — one row per scan with all key fields; frozen header; auto-filter.
///
/// Colour coding:
///   Pass rows → light green fill.  Fail rows → light red fill.
///   Grade A/B badges → green;  C/D → amber;  F → red.
/// </summary>
public static class HistoryExportService
{
    // ── EPPlus colour constants ────────────────────────────────────────────────

    private static readonly System.Drawing.Color ColGradeA    = System.Drawing.Color.FromArgb(0x27, 0xAE, 0x60);
    private static readonly System.Drawing.Color ColGradeB    = System.Drawing.Color.FromArgb(0x2E, 0xCC, 0x71);
    private static readonly System.Drawing.Color ColGradeC    = System.Drawing.Color.FromArgb(0xF3, 0x9C, 0x12);
    private static readonly System.Drawing.Color ColGradeD    = System.Drawing.Color.FromArgb(0xE6, 0x7E, 0x22);
    private static readonly System.Drawing.Color ColGradeF    = System.Drawing.Color.FromArgb(0xE7, 0x4C, 0x3C);
    private static readonly System.Drawing.Color ColPassRow   = System.Drawing.Color.FromArgb(0xEB, 0xF9, 0xEE);
    private static readonly System.Drawing.Color ColFailRow   = System.Drawing.Color.FromArgb(0xFD, 0xED, 0xED);
    private static readonly System.Drawing.Color ColHeader    = System.Drawing.Color.FromArgb(0x2C, 0x3E, 0x50);
    private static readonly System.Drawing.Color ColSubHeader = System.Drawing.Color.FromArgb(0x34, 0x49, 0x5E);
    private static readonly System.Drawing.Color ColLabel     = System.Drawing.Color.FromArgb(0xEC, 0xF0, 0xF1);

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>
    /// Generates a two-sheet XLSX workbook from the supplied history rows and
    /// saves it to <paramref name="filePath"/>.
    /// </summary>
    /// <param name="rows">Rows from <c>HistoryViewModel.AllRecords</c> (full session, unfiltered).</param>
    /// <param name="filePath">Absolute path for the output file.</param>
    /// <param name="jobName">Optional job name shown in the summary header.</param>
    /// <param name="operatorId">Optional operator ID shown in the summary header.</param>
    /// <param name="ct">Cancellation token.</param>
    public static async Task ExportAsync(
        IReadOnlyList<ScanResultRow> rows,
        string filePath,
        string? jobName     = null,
        string? operatorId  = null,
        CancellationToken ct = default)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var pkg = new ExcelPackage();

        BuildSummarySheet(pkg, rows, jobName, operatorId);
        BuildRecordsSheet(pkg, rows);

        await pkg.SaveAsAsync(new FileInfo(filePath), ct);
    }

    // ── Sheet 1: Summary ──────────────────────────────────────────────────────

    private static void BuildSummarySheet(
        ExcelPackage pkg, IReadOnlyList<ScanResultRow> rows,
        string? jobName, string? operatorId)
    {
        var ws = pkg.Workbook.Worksheets.Add("Summary");

        int pass = rows.Count(r => r.IsPass);
        int fail = rows.Count - pass;
        double passRate = rows.Count > 0 ? (pass * 100.0 / rows.Count) : 0.0;

        // ── Title block ──────────────────────────────────────────────────────
        SetCell(ws, 1, 1, "VTCCP — Session Export", bold: true, size: 16, fg: ColHeader);
        SetCell(ws, 2, 1, $"Generated: {DateTime.Now:yyyy-MM-dd  HH:mm:ss}", size: 10, italic: true);

        // ── Session info ─────────────────────────────────────────────────────
        ws.Row(4).Height = 6; // spacer
        SetCell(ws, 5, 1, "SESSION INFORMATION", bold: true, size: 11, fg: ColSubHeader);
        LabelValue(ws, 6,  1, "Job Name",    jobName     ?? "—");
        LabelValue(ws, 7,  1, "Operator",    operatorId  ?? "—");
        LabelValue(ws, 8,  1, "Records",     rows.Count.ToString());

        // ── Pass/Fail summary ────────────────────────────────────────────────
        ws.Row(10).Height = 6;
        SetCell(ws, 11, 1, "PASS / FAIL SUMMARY", bold: true, size: 11, fg: ColSubHeader);
        LabelValue(ws, 12, 1, "Pass",       pass.ToString(), valColor: ColGradeB);
        LabelValue(ws, 13, 1, "Fail",       fail.ToString(), valColor: ColGradeF);
        LabelValue(ws, 14, 1, "Pass Rate",  $"{passRate:F1}%",
            valColor: passRate >= 100.0 ? ColGradeA : passRate >= 70.0 ? ColGradeC : ColGradeF);

        // ── Grade distribution ───────────────────────────────────────────────
        ws.Row(16).Height = 6;
        SetCell(ws, 17, 1, "GRADE DISTRIBUTION", bold: true, size: 11, fg: ColSubHeader);
        int r = 18;
        foreach (string g in new[] { "A", "B", "C", "D", "F", "—" })
        {
            int cnt = rows.Count(row => row.Grade == g);
            if (rows.Count == 0 || cnt > 0)
            {
                LabelValue(ws, r, 1, $"Grade {g}", cnt.ToString(), valColor: GradeColour(g));
                r++;
            }
        }

        // ── Symbology breakdown ───────────────────────────────────────────────
        ws.Row(r).Height = 6; r++;
        SetCell(ws, r, 1, "SYMBOLOGY BREAKDOWN", bold: true, size: 11, fg: ColSubHeader); r++;
        foreach (var grp in rows.GroupBy(x => x.Symbology).OrderBy(g => g.Key))
        {
            LabelValue(ws, r, 1, grp.Key, $"{grp.Count()} record(s)");
            r++;
        }

        // ── Column widths ─────────────────────────────────────────────────────
        ws.Column(1).Width = 26;
        ws.Column(2).Width = 22;
        ws.View.ShowGridLines = false;
    }

    private static void LabelValue(
        ExcelWorksheet ws, int row, int col,
        string label, string value,
        System.Drawing.Color? valColor = null)
    {
        var lbl = ws.Cells[row, col];
        lbl.Value = label;
        lbl.Style.Font.Bold  = true;
        lbl.Style.Fill.PatternType = ExcelFillStyle.Solid;
        lbl.Style.Fill.BackgroundColor.SetColor(ColLabel);

        var val = ws.Cells[row, col + 1];
        val.Value = value;
        if (valColor.HasValue)
        {
            val.Style.Font.Color.SetColor(valColor.Value);
            val.Style.Font.Bold = true;
        }
    }

    // ── Sheet 2: Scan Records ─────────────────────────────────────────────────

    private static void BuildRecordsSheet(ExcelPackage pkg, IReadOnlyList<ScanResultRow> rows)
    {
        var ws = pkg.Workbook.Worksheets.Add("Scan Records");

        // ── Header ────────────────────────────────────────────────────────────
        string[] headers =
        [
            "#", "Time", "Symbology",
            "Numeric Grade", "Letter", "Pass/Fail",
            "UEC%", "Full Decoded Data", "Operator", "Job",
        ];

        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cells[1, c + 1];
            cell.Value = headers[c];
            cell.Style.Font.Bold  = true;
            cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(ColHeader);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        // ── Data rows ─────────────────────────────────────────────────────────
        for (int i = 0; i < rows.Count; i++)
        {
            ScanResultRow row = rows[i];
            int r = i + 2;

            ws.Cells[r, 1].Value  = row.RowNumber;
            ws.Cells[r, 2].Value  = row.Time;
            ws.Cells[r, 3].Value  = row.Symbology;
            ws.Cells[r, 4].Value  = row.NumericGradeDisplay;
            ws.Cells[r, 5].Value  = row.Grade;
            ws.Cells[r, 6].Value  = row.PassFail;
            ws.Cells[r, 7].Value  = row.UecPercent.HasValue ? (object)row.UecPercent.Value : string.Empty;
            ws.Cells[r, 8].Value  = row.FullDecodedData;
            ws.Cells[r, 9].Value  = row.OperatorId;
            ws.Cells[r, 10].Value = row.JobName;

            // Row fill — pass/fail colouring
            System.Drawing.Color rowFill = row.PassFail == "Pass" ? ColPassRow
                                         : row.PassFail == "Fail" ? ColFailRow
                                         : System.Drawing.Color.White;

            var rowRange = ws.Cells[r, 1, r, headers.Length];
            rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rowRange.Style.Fill.BackgroundColor.SetColor(rowFill);

            // Grade letter cell badge colour
            ws.Cells[r, 5].Style.Font.Color.SetColor(GradeColour(row.Grade));
            ws.Cells[r, 5].Style.Font.Bold = true;
            ws.Cells[r, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // UEC% one decimal
            if (row.UecPercent.HasValue)
                ws.Cells[r, 7].Style.Numberformat.Format = "0.0";
        }

        // ── Formatting ────────────────────────────────────────────────────────
        ws.Cells[1, 1, 1, headers.Length].AutoFilter = true;
        ws.View.FreezePanes(2, 1);

        ws.Column(1).Width  = 6;
        ws.Column(2).Width  = 10;
        ws.Column(3).Width  = 24;
        ws.Column(4).Width  = 14;
        ws.Column(5).Width  = 8;
        ws.Column(6).Width  = 10;
        ws.Column(7).Width  = 8;
        ws.Column(8).Width  = 56;
        ws.Column(9).Width  = 14;
        ws.Column(10).Width = 18;

        // Wrap long decoded data column
        ws.Column(8).Style.WrapText = true;
        ws.View.ShowGridLines = true;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static void SetCell(
        ExcelWorksheet ws, int row, int col, string value,
        bool bold = false, int size = 11, bool italic = false,
        System.Drawing.Color? fg = null)
    {
        var cell = ws.Cells[row, col];
        cell.Value = value;
        if (bold)   cell.Style.Font.Bold   = true;
        if (italic) cell.Style.Font.Italic = true;
        cell.Style.Font.Size = size;
        if (fg.HasValue) cell.Style.Font.Color.SetColor(fg.Value);
    }

    private static System.Drawing.Color GradeColour(string grade) => grade switch
    {
        "A" => ColGradeA,
        "B" => ColGradeB,
        "C" => ColGradeC,
        "D" => ColGradeD,
        "F" => ColGradeF,
        _   => System.Drawing.Color.Gray,
    };
}
