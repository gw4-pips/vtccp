using System.Text;
using ExcelEngine.Models;

namespace DeviceInterface.Reports;

/// <summary>
/// Generates a self-contained hybrid HTML verification report that combines
/// Webscan TruCheck™ barcode grading data with VCCS FlexWedge™ RFID
/// validation results.
///
/// Visual structure mirrors the Webscan TruCheck HTML output exactly
/// (same inline table CSS, same section headers, same column layout),
/// with three additions:
///   1. An RFID status badge in the report header.
///   2. A dedicated "RFID Validation — VCCS FlexWedge™" section inserted
///      between Verification Grades and the Symbol Image.
///   3. CSS @media print / @page rules for direct PDF export via the browser.
///
/// The output is fully self-contained: all images are embedded as base64
/// data URIs and all styles are in a single &lt;style&gt; block.
/// No external files (logo.jpg, reportImage1.jpg) are referenced.
/// </summary>
public static class HybridReportGenerator
{
    // ── Public API ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Builds the hybrid HTML string from the supplied verification record.
    /// </summary>
    public static string Generate(VerificationRecord r)
    {
        var sb = new StringBuilder(32_768);
        AppendHtmlHead(sb, r);
        AppendReportHeader(sb, r);
        sb.AppendLine("<body>");
        AppendPrintButton(sb);
        AppendReportSummary(sb, r);
        AppendVerificationGrades(sb, r);
        AppendRfidSection(sb, r);
        AppendImageSection(sb, r);
        AppendQualityParameters(sb, r);
        AppendFooter(sb);
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        return sb.ToString();
    }

    /// <summary>
    /// Generates the hybrid HTML and writes it to <paramref name="outputDir"/>.
    /// The filename is derived from <see cref="VerificationRecord.VerificationDateTime"/>.
    /// Supply <paramref name="filenameOverride"/> (without extension) to use a custom name.
    /// Does nothing if <paramref name="outputDir"/> is null or whitespace.
    /// </summary>
    public static async Task SaveAsync(
        VerificationRecord r,
        string             outputDir,
        string?            filenameOverride = null,
        CancellationToken  ct               = default)
    {
        if (string.IsNullOrWhiteSpace(outputDir)) return;

        string html     = Generate(r);
        string filename = (filenameOverride is { Length: > 0 } s ? s : null)
            ?? $"{r.VerificationDateTime:yyyy-MM-dd_HH-mm-ss}_hybrid_report.html";
        Directory.CreateDirectory(outputDir);
        string path = Path.Combine(outputDir, filename);
        await File.WriteAllTextAsync(path, html, System.Text.Encoding.UTF8, ct)
                  .ConfigureAwait(false);
    }

    // ── Top-level section builders ─────────────────────────────────────────────

    private static void AppendHtmlHead(StringBuilder sb, VerificationRecord r)
    {
        const string title =
            "Hybrid Webscan TruCheck\u2122 \u00b7 VCCS FlexWedge\u2122 RFID Verification Report";

        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("  <meta charset=\"utf-8\"/>");
        sb.AppendLine("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"/>");
        sb.AppendLine($"  <title>{H(title)}</title>");
        sb.AppendLine("  <style type=\"text/css\">");
        sb.AppendLine(BaseCss);
        sb.AppendLine("  </style>");
        sb.AppendLine("</head>");
    }

    private static void AppendReportHeader(StringBuilder sb, VerificationRecord r)
    {
        string dateTime = r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt");
        string device   = H(r.DeviceModel   ?? "Cognex DataMan");
        string fw       = H(r.FirmwareVersion ?? "\u2014");
        string serial   = H(r.DeviceSerial  ?? "\u2014");
        string rfidBadge = RfidStatusBadge(r);

        sb.AppendLine("<!-- Report Header -->");
        sb.AppendLine("<div>");
        sb.AppendLine("  <table width=\"100%\">");
        sb.AppendLine("    <col width=\"1.4in\"/>");
        sb.AppendLine("    <col/>");
        sb.AppendLine("    <tr>");

        // Left column: VCCS logo placeholder (text SVG-style box)
        sb.AppendLine("      <td style=\"vertical-align:middle;\">");
        sb.AppendLine("        <div style=\"margin:6pt;padding:6pt;border:2px solid black;text-align:center;font-weight:bold;\">");
        sb.AppendLine("          <div style=\"font-size:14pt;\">VCCS</div>");
        sb.AppendLine("          <div style=\"font-size:7pt;font-weight:normal;letter-spacing:1pt;\">FlexWedge\u2122</div>");
        sb.AppendLine("        </div>");
        sb.AppendLine("      </td>");

        // Right column: title + meta rows (matches Webscan header layout exactly)
        sb.AppendLine("      <td>");
        sb.AppendLine("        <div style=\"text-align:center;font-size:13pt;font-weight:bold;\">");
        sb.AppendLine("          <h1 style=\"font-size:13pt;\">");
        sb.AppendLine("            Hybrid Webscan TruCheck\u2122 + VCCS FlexWedge\u2122 RFID Verification Report");
        sb.AppendLine("          </h1>");
        sb.AppendLine($"          <h2>{H(dateTime)}</h2>");
        sb.AppendLine($"          <h2>Device: {device} &nbsp;&middot;&nbsp; Firmware: {fw}</h2>");
        sb.AppendLine($"          <h2>Serial Number: {serial}</h2>");
        sb.AppendLine($"          <h2>RFID Tag: {rfidBadge}</h2>");
        sb.AppendLine("        </div>");
        sb.AppendLine("      </td>");

        sb.AppendLine("    </tr>");
        sb.AppendLine("  </table>");
        sb.AppendLine("</div>");
    }

    private static void AppendPrintButton(StringBuilder sb)
    {
        sb.AppendLine("  <!-- Print button — hidden by @media print -->");
        sb.AppendLine("  <div class=\"no-print\" style=\"text-align:center;margin:6pt;\">");
        sb.AppendLine("    <button class=\"print-btn\" onclick=\"window.print()\">");
        sb.AppendLine("      &#128438; Print / Save as PDF");
        sb.AppendLine("    </button>");
        sb.AppendLine("  </div>");
    }

    private static void AppendReportSummary(StringBuilder sb, VerificationRecord r)
    {
        sb.AppendLine("  <!-- Report Summary -->");
        sb.AppendLine("  <div span=\"all\">");
        sb.AppendLine("    <table style=\"font-size:9pt;border-style:solid;text-align:left;\" width=\"100%\">");
        sb.AppendLine("      <col width=\"1.5in\"/>");
        sb.AppendLine("      <col/>");

        BlackHeaderRow(sb, "Report Summary", 2);

        // Required rows
        LabelValueRow(sb, "Data",
            $"<div style=\"font-size:9pt;padding:0.025in;\">{H(r.DecodedData ?? "NO DECODE")}</div>");
        LabelValueRow(sb, "Symbology", Padded(H(r.Symbology ?? "\u2014")));

        // Optional user-context rows (present only when populated)
        if (!string.IsNullOrWhiteSpace(r.CompanyName))
            LabelValueRow(sb, "Company",  Padded(H(r.CompanyName)));
        if (!string.IsNullOrWhiteSpace(r.OperatorId))
            LabelValueRow(sb, "Operator", Padded(H(r.OperatorId)));
        if (!string.IsNullOrWhiteSpace(r.JobName))
            LabelValueRow(sb, "Job",      Padded(H(r.JobName)));
        if (!string.IsNullOrWhiteSpace(r.BatchNumber))
            LabelValueRow(sb, "Batch",    Padded(H(r.BatchNumber)));
        if (!string.IsNullOrWhiteSpace(r.ProductName))
            LabelValueRow(sb, "Product",  Padded(H(r.ProductName)));

        sb.AppendLine("    </table>");
        sb.AppendLine("  </div>");
    }

    private static void AppendVerificationGrades(StringBuilder sb, VerificationRecord r)
    {
        // Match Webscan OverallGrades output exactly — same structure, same CSS
        string standard   = H(r.Standard ?? "\u2014");
        string grade      = GradeDisplay(r.OverallGrade);
        string aperture   = r.Aperture.HasValue ? r.Aperture.Value.ToString("D2") : "\u2014";
        string wavelength = r.Wavelength.HasValue ? r.Wavelength.Value.ToString() : "\u2014";
        string lighting   = H(r.Lighting ?? "\u2014");
        string formal     = H(r.FormalGrade ?? "\u2014");

        sb.AppendLine("  <!-- Verification Grades -->");
        sb.AppendLine("  <div span=\"all\">");
        sb.AppendLine("    <table width=\"100%\" style=\"font-size:9pt;border-style:solid;text-align:center;margin-top:8pt;\">");
        sb.AppendLine("      <col/>");
        sb.AppendLine("      <tr>");
        sb.AppendLine("        <td style=\"border-style:solid;background-color:black;color:white;\">");
        sb.AppendLine("          <h2 style=\"padding:0.025in;font-weight:bold;font-size:11pt;\">Verification Grades</h2>");
        sb.AppendLine("        </td>");
        sb.AppendLine("      </tr>");
        sb.AppendLine("      <tr><td style=\"border-style:solid;\">");
        sb.AppendLine("        <table width=\"100%\">");
        sb.AppendLine("          <col/><col/><col/><col/><col/><col/>");
        sb.AppendLine("          <tr style=\"font-weight:bold;text-align:center;\">");
        foreach (string col in new[] { "Standard", "Grade", "Aperture", "Wavelength", "Lighting", "Formal Grade" })
            sb.AppendLine($"            <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;\">{col}</div></td>");
        sb.AppendLine("          </tr>");
        sb.AppendLine("          <tr>");
        foreach (string val in new[] { standard, grade, aperture, wavelength, lighting, formal })
            sb.AppendLine($"            <td style=\"border-style:solid;border-width:thin;display-align:center;\"><div style=\"padding:0.025in;\">{val}</div></td>");
        sb.AppendLine("          </tr>");
        sb.AppendLine("        </table>");
        sb.AppendLine("      </td></tr>");
        sb.AppendLine("    </table>");
        sb.AppendLine("  </div>");
    }

    private static void AppendRfidSection(StringBuilder sb, VerificationRecord r)
    {
        bool rfidPerformed = !string.IsNullOrWhiteSpace(r.RfidStatus);

        sb.AppendLine("  <!-- RFID Validation — VCCS FlexWedge™ -->");
        sb.AppendLine("  <div span=\"all\">");
        sb.AppendLine("    <table width=\"100%\" style=\"font-size:9pt;border-style:solid;text-align:left;margin-top:8pt;\">");
        sb.AppendLine("      <col width=\"2.2in\"/>");
        sb.AppendLine("      <col/>");

        // Section header — matches Webscan black-header style, VCCS brand in label
        sb.AppendLine("      <tr>");
        sb.AppendLine("        <td style=\"border-style:solid;background-color:black;color:white;\" colspan=\"2\">");
        sb.AppendLine("          <div style=\"padding:0.025in;font-weight:bold;text-align:center;font-size:11pt;\">");
        sb.AppendLine("            RFID Validation \u2014 VCCS FlexWedge\u2122");
        sb.AppendLine("          </div>");
        sb.AppendLine("        </td>");
        sb.AppendLine("      </tr>");

        if (!rfidPerformed)
        {
            // RFID not configured for this session
            sb.AppendLine("      <tr>");
            sb.AppendLine("        <td colspan=\"2\" style=\"border-style:solid;border-width:thin;\">");
            sb.AppendLine("          <div style=\"padding:0.025in;color:#666;font-style:italic;\">");
            sb.AppendLine("            RFID validation was not performed for this scan.");
            sb.AppendLine("            Enable RFID by configuring the ASR-P35U reader COM port in Application Settings.");
            sb.AppendLine("          </div>");
            sb.AppendLine("        </td>");
            sb.AppendLine("      </tr>");
        }
        else
        {
            // Tag detected inferred from status (NoTag enum member → ToString = "NoTag")
            bool tagDetected = r.RfidStatus is not ("NoTag" or "NO_TAG");
            RfidLabelRow(sb, "Tag Detected", tagDetected ? "Yes" : "No");

            if (tagDetected)
            {
                RfidLabelRow(sb, "EPC (Hex)",        r.RfidEpcHex  ?? "\u2014");
                RfidLabelRow(sb, "Decoded GTIN-14",  r.RfidGtin14  ?? "\u2014");
                RfidLabelRow(sb, "GCP Valid",         r.RfidGcpValid.HasValue
                    ? (r.RfidGcpValid.Value ? "Yes" : "No")
                    : "\u2014");
                RfidLabelRow(sb, "EPC Serial",        r.RfidSerial  ?? "\u2014");
            }

            // Validation result row — colour-coded by status
            (string style, string display) = r.RfidStatus switch
            {
                "Pass"                 => ("color:#155724;font-weight:bold;",
                                           "\u2713 PASS \u2014 EPC matches barcode data"),
                "Fail"                 => ("color:#721c24;font-weight:bold;",
                                           "\u2718 FAIL \u2014 EPC does not match barcode data"),
                "NoTag"                => ("color:#6c757d;",
                                           "\u2014 No tag detected in scan window"),
                "ParseError"           => ("color:#856404;font-weight:bold;",
                                           "! Parse error \u2014 EPC could not be decoded"),
                "MultipleTagsDetected" => ("color:#856404;font-weight:bold;",
                                           "! Multiple tags \u2014 ambiguous read"),
                _                      => ("", H(r.RfidStatus)),
            };

            sb.AppendLine("      <tr style=\"font-weight:bold;\">");
            sb.AppendLine($"        <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;\">Validation Result</div></td>");
            sb.AppendLine($"        <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;{style}\">{display}</div></td>");
            sb.AppendLine("      </tr>");

            // Mismatch detail (only when present)
            if (!string.IsNullOrWhiteSpace(r.RfidMismatchDetail))
                RfidLabelRow(sb, "Mismatch Detail", H(r.RfidMismatchDetail));

            // Scan window
            if (r.RfidScanWindowMs.HasValue)
                RfidLabelRow(sb, "Scan Window", $"{r.RfidScanWindowMs.Value} ms");
        }

        sb.AppendLine("    </table>");
        sb.AppendLine("  </div>");
    }

    private static void AppendImageSection(StringBuilder sb, VerificationRecord r)
    {
        // Prefer ROI (barcode-crop) image; fall back to full-frame SDK image
        string? b64 = r.RoiJpegImageBase64 ?? r.JpegImageBase64;
        if (string.IsNullOrWhiteSpace(b64)) return;

        string label = r.RoiJpegImageBase64 is { Length: > 0 }
            ? "Symbol Image (ROI Crop)"
            : "Symbol Image";

        sb.AppendLine("  <!-- Image Section -->");
        sb.AppendLine("  <div style=\"font-size:8pt;border-style:solid;margin-top:8pt;\">");
        sb.AppendLine($"    <h2 style=\"background-color:black;color:white;padding:0.025in;text-align:center;" +
                      $"border-style:solid;font-weight:bold;font-size:10pt;\">{H(label)}</h2>");
        sb.AppendLine("    <div>");
        sb.AppendLine($"      <img src=\"data:image/jpeg;base64,{b64}\" alt=\"Symbol Image\"" +
                      $" style=\"width:auto;max-height:4in;max-width:100%;\" />");
        sb.AppendLine("    </div>");
        sb.AppendLine("  </div>");
    }

    private static void AppendQualityParameters(StringBuilder sb, VerificationRecord r)
    {
        // Collect populated grade rows depending on symbology family
        var rows = new List<(string Name, string Value, string Grade)>();

        void Add(string name, GradingResult? g, decimal? numericValue = null, string? valueSuffix = null)
        {
            if (g is null && numericValue is null) return;
            string v = numericValue.HasValue
                ? $"{numericValue.Value:F1}{valueSuffix}"
                : g?.NumericGradeString ?? "\u2014";
            rows.Add((name, v, GradeDisplay(g)));
        }

        if (r.Is2D)
        {
            Add("Decode",           r.DECODE_Grade);
            Add("SC",               r.SC_Grade);
            Add("MOD",              r.MOD_Grade);
            Add("ANU",              r.ANU_Grade,  r.ANU_Percent, "%");
            Add("GNU",              r.GNU_Grade,  r.GNU_Percent, "%");
            Add("FPD",              r.FPD_Grade,  r.FPD_Value);
            Add("UEC",              r.UEC_Grade,  r.UEC_Percent, "%");
            Add("Refl. Margin",     r.RM_Grade);
            Add("Avg. Grade",       r.AverageGrade, r.AverageGradeNumeric);
        }
        else if (r.Is1D && r.ScanResults.Count > 0)
        {
            // For 1D, summarise averages from the scan results aggregate
            Add("Avg SC",  null, r.Avg_SC);
            Add("Avg MOD", null, r.Avg_MOD);
            Add("Avg EC",  null, r.Avg_MinEC);
        }

        if (rows.Count == 0) return;

        string sectionTitle = r.Is2D ? "ISO 15415 Quality Parameters" : "ISO 15416 Quality Parameters";

        sb.AppendLine("  <!-- Quality Parameters -->");
        sb.AppendLine("  <table width=\"100%\" style=\"font-size:8pt;border-style:solid;margin-top:8pt;\">");
        sb.AppendLine("    <col/><col/><col/>");

        // Section header
        sb.AppendLine("    <tr>");
        sb.AppendLine("      <td style=\"background-color:black;color:white;\" colspan=\"3\">");
        sb.AppendLine($"        <h2 style=\"padding:0.025in;text-align:center;border-style:solid;font-weight:bold;font-size:10pt;\">{H(sectionTitle)}</h2>");
        sb.AppendLine("      </td>");
        sb.AppendLine("    </tr>");

        // Column headers (matches Webscan: Parameter | Value | Grade)
        sb.AppendLine("    <tr style=\"text-align:center;font-weight:bold;\">");
        foreach (string col in new[] { "Parameter", "Value", "Grade" })
            sb.AppendLine($"      <th style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;\">{col}</div></th>");
        sb.AppendLine("    </tr>");

        foreach (var (name, value, grade) in rows)
        {
            sb.AppendLine("    <tr>");
            sb.AppendLine($"      <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;\">{H(name)}</div></td>");
            sb.AppendLine($"      <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;\">{value}</div></td>");
            sb.AppendLine($"      <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;\">{grade}</div></td>");
            sb.AppendLine("    </tr>");
        }

        sb.AppendLine("  </table>");
    }

    private static void AppendFooter(StringBuilder sb)
    {
        sb.AppendLine("  <div class=\"vccs-footer\">");
        sb.AppendLine("    Generated by VTCCP &nbsp;&middot;&nbsp; VCCS FlexWedge\u2122 RFID Validation &nbsp;&middot;&nbsp; Hybrid Report v1.0");
        sb.AppendLine("  </div>");
    }

    // ── Structural helpers ─────────────────────────────────────────────────────

    private static void BlackHeaderRow(StringBuilder sb, string title, int colspan)
    {
        sb.AppendLine($"      <tr><td style=\"border-style:solid;background-color:black;color:white;\" colspan=\"{colspan}\">");
        sb.AppendLine($"        <div style=\"padding:0.025in;font-weight:bold;text-align:center;font-size:11pt;\">{H(title)}</div>");
        sb.AppendLine("      </td></tr>");
    }

    private static void LabelValueRow(StringBuilder sb, string label, string valueHtml)
    {
        sb.AppendLine("      <tr style=\"font-weight:bold;\">");
        sb.AppendLine($"        <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;\">{H(label)}</div></td>");
        sb.AppendLine($"        <td style=\"border-style:solid;border-width:thin;\" colspan=\"1\">{valueHtml}</td>");
        sb.AppendLine("      </tr>");
    }

    private static void RfidLabelRow(StringBuilder sb, string label, string value)
    {
        sb.AppendLine("      <tr>");
        sb.AppendLine($"        <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;\">{H(label)}</div></td>");
        sb.AppendLine($"        <td style=\"border-style:solid;border-width:thin;\"><div style=\"padding:0.025in;font-family:Consolas,monospace;\">{value}</div></td>");
        sb.AppendLine("      </tr>");
    }

    private static string Padded(string content) =>
        $"<div style=\"padding:0.025in;\">{content}</div>";

    // ── Display formatters ─────────────────────────────────────────────────────

    /// <summary>
    /// RFID status badge inline HTML matching the header style of the Webscan report.
    /// Uses colour spans (no external CSS class references).
    /// </summary>
    private static string RfidStatusBadge(VerificationRecord r) =>
        string.IsNullOrWhiteSpace(r.RfidStatus)
            ? "<span style=\"color:#6c757d;\">\u2014 Not configured</span>"
            : r.RfidStatus switch
            {
                "Pass"  => Badge("#d4edda", "#155724", "#c3e6cb", "\u2713 MATCHED"),
                "Fail"  => Badge("#f8d7da", "#721c24", "#f5c6cb", "\u2718 MISMATCH"),
                "NoTag" => Badge("#e2e3e5", "#383d41", "#d6d8db", "\u2014 No Tag"),
                "ParseError"
                    or "MultipleTagsDetected"
                        => Badge("#fff3cd", "#856404", "#ffeeba", "! " + r.RfidStatus),
                _       => $"<span>{H(r.RfidStatus)}</span>",
            };

    private static string Badge(string bg, string fg, string border, string text) =>
        $"<span style=\"background:{bg};color:{fg};padding:1pt 6pt;" +
        $"border:1px solid {border};font-size:10pt;\">{text}</span>";

    /// <summary>
    /// Formats a <see cref="GradingResult"/> in Webscan style: "B (3.5)" or "F (0.0)".
    /// Returns "—" when null.
    /// </summary>
    private static string GradeDisplay(GradingResult? g)
    {
        if (g is null) return "\u2014";
        string letter  = g.LetterGradeString;
        bool hasLetter = letter is { Length: > 0 };
        bool hasNum    = g.NumericGrade.HasValue;

        if (hasLetter && hasNum)  return H($"{letter} ({g.NumericGrade!.Value:F1})");
        if (hasLetter)            return H(letter);
        if (hasNum)               return g.NumericGrade!.Value.ToString("F1");
        return "\u2014";
    }

    /// <summary>Minimal HTML entity encoding for untrusted text content.</summary>
    private static string H(string? s) =>
        s is null ? string.Empty
        : s.Replace("&", "&amp;")
           .Replace("<", "&lt;")
           .Replace(">", "&gt;")
           .Replace("\"", "&quot;")
           .Replace("'", "&#39;");

    // ── CSS ────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Base CSS.  The first block is a verbatim copy of the Webscan TruCheck
    /// inline style rules (extracted from html_stylesheet.xslt).
    /// The second block adds VCCS-specific utilities and print rules.
    /// </summary>
    private const string BaseCss = """
        /* ── Webscan TruCheck base (verbatim match) ─────────────────────────── */
        table {
            border-color: Black;
            border-collapse: collapse;
            border-spacing: 0;
        }
        td  { border-color: Black; padding: 0; }
        th  { border-color: Black; padding: 0; }
        div { border-color: Black; }
        h1  { border-color: Black; margin: 0; }
        h2  { border-color: Black; margin: 0; }
        h3  { border-color: Black; margin: 0; }
        ul  { list-style-type: none; }
        li  { list-style-type: none; }
        body { text-align: center; font-family: Arial, Helvetica, sans-serif; }

        /* ── VCCS FlexWedge™ additions ──────────────────────────────────────── */
        .print-btn {
            display: inline-block;
            margin: 8pt;
            padding: 5pt 14pt;
            background: #333;
            color: #fff;
            border: none;
            cursor: pointer;
            font-size: 9pt;
            font-family: inherit;
        }
        .print-btn:hover { background: #555; }

        .vccs-footer {
            margin-top: 16pt;
            padding-top: 6pt;
            font-size: 7.5pt;
            color: #666;
            border-top: 1px solid #ccc;
            text-align: center;
        }

        /* ── Print / PDF ────────────────────────────────────────────────────── */
        @media print {
            .no-print { display: none !important; }
            body { text-align: center; }
            @page { margin: 0.65in; }
        }
        """;
}
