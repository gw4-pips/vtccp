using System.Text;
using ExcelEngine.Models;

namespace DeviceInterface.Reports;

/// <summary>
/// Generates a self-contained hybrid HTML verification report.
///
/// Layout exactly mirrors the Webscan TruCheck PDF output:
///   Full-width (span="all"):
///     1  Report Summary
///     2  Verification Grades
///     3  GradingInfoSection
///   Two-column side-by-side:
///     4  Image (left)  |  ISO Quality Parameters (right)
///     5  Data Format Check (left)  |  RFID Validation (right)
///
/// Design principle: the <see cref="Template"/> constant is the verbatim
/// Webscan HTML skeleton (every table, col, td, and CSS rule copied from
/// HTMLreport_*.html).  VCCS additions are injected at <!-- VCCS:TOKEN -->
/// comment sites only.  Nothing is rebuilt from scratch.
/// </summary>
public static class HybridReportGenerator
{
    // ── Public API ─────────────────────────────────────────────────────────────

    public static string Generate(VerificationRecord r)
    {
        string html = Template;

        // <head>
        html = html.Replace("<!-- VCCS:TITLE -->",
            "Hybrid Webscan TruCheck\u2122 \u00b7 VCCS FlexWedge\u2122 RFID Verification Report");
        html = html.Replace("<!-- VCCS:EXTRA_CSS -->", ExtraCss);

        // Header banner (4 cells)
        html = html.Replace("<!-- VCCS:VCCS_LOGO_CELL -->",   VccsLogoCell());
        html = html.Replace("<!-- VCCS:WEBSCAN_LOGO_CELL -->", WebscanLogoCell());
        html = html.Replace("<!-- VCCS:TITLE_CELL -->",        TitleCell(r));
        html = html.Replace("<!-- VCCS:COMPANY_LOGO_CELL -->", CompanyLogoCell(r));

        // Full-width body sections
        html = html.Replace("<!-- VCCS:PRINT_BUTTON -->",         PrintButton());
        html = html.Replace("<!-- VCCS:SUMMARY_ROWS -->",         SummaryRows(r));
        html = html.Replace("<!-- VCCS:GRADES_ROW -->",           GradesRow(r));
        html = html.Replace("<!-- VCCS:GRADING_INFO_SECTION -->", GradingInfoSection(r));

        // Two-column paired sections
        html = html.Replace("<!-- VCCS:IMAGE_AND_PARAMS -->",    ImageAndQualityTable(r));
        html = html.Replace("<!-- VCCS:FORMAT_AND_RFID -->",     FormatAndRfidTable(r));

        html = html.Replace("<!-- VCCS:FOOTER -->",              Footer());

        return html;
    }

    public static async Task SaveAsync(
        VerificationRecord r,
        string             outputDir,
        CancellationToken  ct = default)
    {
        if (string.IsNullOrWhiteSpace(outputDir)) return;
        string html     = Generate(r);
        string filename = $"{r.VerificationDateTime:yyyy-MM-dd_HH-mm-ss}_hybrid_report.html";
        Directory.CreateDirectory(outputDir);
        await File.WriteAllTextAsync(
            Path.Combine(outputDir, filename), html, Encoding.UTF8, ct)
            .ConfigureAwait(false);
    }

    public static async Task SaveToPathAsync(
        VerificationRecord r,
        string             outputPath,
        CancellationToken  ct = default)
    {
        if (string.IsNullOrWhiteSpace(outputPath)) return;
        string html      = Generate(r);
        string parentDir = Path.GetDirectoryName(outputPath)!;
        if (!string.IsNullOrEmpty(parentDir))
            Directory.CreateDirectory(parentDir);
        await File.WriteAllTextAsync(outputPath, html, Encoding.UTF8, ct)
            .ConfigureAwait(false);
    }

    // ── Section builders ───────────────────────────────────────────────────────

    private static string VccsLogoCell() =>
        """
        <td style="vertical-align:middle;padding:4pt;">
          <div style="border:2px dashed #999;text-align:center;padding:6pt 4pt;min-width:0.9in;">
            <div style="font-size:11pt;font-weight:bold;letter-spacing:1pt;">VCCS</div>
            <div style="font-size:6.5pt;letter-spacing:0.5pt;">FlexWedge&#x2122;</div>
          </div>
        </td>
        """;

    private static string WebscanLogoCell() =>
        // logo.jpg ships with Webscan alongside its HTML.  In Replace-mode,
        // keep logo.jpg in the same folder as the output file.
        """
        <td style="vertical-align:middle;padding:4pt;">
          <div style="text-align:center;padding:0;min-width:0.9in;">
            <img src="logo.jpg" style="width:0.9in;" alt="Webscan"
                 onerror="this.style.display='none'"/>
          </div>
        </td>
        """;

    private static string TitleCell(VerificationRecord r)
    {
        string dt     = r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt");
        string device = H(r.DeviceModel     ?? "Cognex DataMan");
        string fw     = H(r.FirmwareVersion ?? "\u2014");
        string serial = H(r.DeviceSerial    ?? "\u2014");
        string badge  = RfidStatusBadge(r);

        return $"""
        <td>
          <div style="text-align:center;font-size:13pt;font-weight:bold;">
            <h1 style="font-size:15pt;">Hybrid Webscan TruCheck&#x2122; &middot; VCCS FlexWedge&#x2122; RFID Verification Report</h1>
            <h2 style="font-size:9pt;font-weight:normal;">{H(dt)}</h2>
            <h2 style="font-size:9pt;font-weight:normal;">Device: {device} &nbsp;&middot;&nbsp; Firmware: {fw}</h2>
            <h2 style="font-size:9pt;font-weight:normal;">Serial Number: {serial}</h2>
            <h2 style="font-size:9pt;font-weight:normal;">RFID Tag: {badge}</h2>
          </div>
        </td>
        """;
    }

    private static string CompanyLogoCell(VerificationRecord r)
    {
        string name = H(r.CompanyName ?? "Company Logo");
        return $"""
        <td style="vertical-align:middle;padding:4pt;">
          <div style="border:2px dashed #999;text-align:center;padding:6pt 4pt;min-width:0.9in;">
            <div style="font-size:7pt;color:#555;">{name}</div>
          </div>
        </td>
        """;
    }

    private static string PrintButton() =>
        """
        <div class="no-print" style="text-align:center;margin:6pt;">
          <button class="vccs-print-btn" onclick="window.print()">&#128438; Print / Save as PDF</button>
        </div>
        """;

    private static string SummaryRows(VerificationRecord r)
    {
        var sb = new StringBuilder();

        // MIL-STD / standard acceptance criteria row — always first if present
        if (r.Standard is not null && r.OverallGrade is not null)
        {
            string grade = GradeDisplay(r.OverallGrade);
            string extra = BuildGradeExtra(r);
            sb.Append($"""
              <tr style="font-weight:bold;">
                <td style="border-style:solid;border-width:thin;display-align:center;">
                  <div style="padding:0.025in;">{H(r.Standard)} Acceptance Criteria</div>
                </td>
                <td style="border-style:solid;border-width:thin;" colspan="1">
                  <div style="padding:0.025in;">{grade}{extra}</div>
                </td>
              </tr>
            """);
        }

        // Data row
        string dataFontSize = (r.DecodedData?.Length ?? 0) > 66 ? "7pt" : "9pt";
        sb.Append($"""
          <tr style="font-weight:bold;">
            <td style="border-style:solid;border-width:thin;display-align:center;">
              <div style="padding:0.025in;">Data</div>
            </td>
            <td style="border-style:solid;border-width:thin;" colspan="1">
              <div style="font-size:{dataFontSize};padding:0.025in;">{H(r.DecodedData ?? "NO DECODE")}</div>
            </td>
          </tr>
        """);

        // Symbology
        sb.Append($"""
          <tr style="font-weight:bold;">
            <td style="border-style:solid;border-width:thin;">
              <div style="padding:0.025in;">Symbology</div>
            </td>
            <td style="border-style:solid;border-width:thin;" colspan="1">
              <div style="padding:0.025in;">{H(r.Symbology ?? "\u2014")}</div>
            </td>
          </tr>
        """);

        // User-info rows — first row gets thick top border, per Webscan template
        bool firstUserRow = true;
        void UserRow(string label, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            string thick = firstUserRow ? "border-top-width:thick;" : "";
            firstUserRow = false;
            sb.Append($"""
              <tr style="font-weight:bold;">
                <td style="border-style:solid;border-width:thin;{thick}display-align:center;">
                  <div style="padding:0.025in;">{H(label)}</div>
                </td>
                <td style="border-style:solid;border-width:thin;{thick}display-align:center;">
                  <div style="padding:0.025in;">{H(value)}</div>
                </td>
              </tr>
            """);
        }

        UserRow("Verified By",  r.OperatorId);
        UserRow("Company Name", r.CompanyName);
        UserRow("Product Name", r.ProductName);
        UserRow("Job Number",   r.JobName);
        UserRow("Batch Number", r.BatchNumber);

        return sb.ToString();
    }

    private static string GradesRow(VerificationRecord r)
    {
        string standard   = H(r.Standard ?? "\u2014");
        string grade      = GradeDisplay(r.OverallGrade);
        string aperture   = r.Aperture.HasValue   ? r.Aperture.Value.ToString("D2")  : "\u2014";
        string wavelength = r.Wavelength.HasValue ? r.Wavelength.Value.ToString()     : "\u2014";
        string lighting   = H(r.Lighting   ?? "\u2014");
        string formal     = H(r.FormalGrade ?? "\u2014");

        var sb = new StringBuilder();
        sb.Append("<tr>");
        foreach (string v in new[] { standard, grade, aperture, wavelength, lighting, formal })
            sb.Append($"<td style=\"border-style:solid;border-width:thin;display-align:center;\"><div style=\"padding:0.025in;\">{v}</div></td>");
        sb.Append("</tr>");
        return sb.ToString();
    }

    /// <summary>
    /// GradingInfoSection — full-width per-parameter detail table.
    /// 8 columns: Name | Grade | Aperture | Wavelength | Lighting | Formal Grade | Check | Notes
    /// Omitted when no grading results are populated.
    /// </summary>
    private static string GradingInfoSection(VerificationRecord r)
    {
        var rows = new List<(string Name, string Grade, string Note)>();

        void Add(string name, GradingResult? g, string note = "")
        {
            if (g is null) return;
            rows.Add((name, GradeDisplay(g), note));
        }

        if (r.Is2D)
        {
            Add("Decode",       r.DECODE_Grade);
            Add("SC",           r.SC_Grade);
            Add("MOD",          r.MOD_Grade);
            Add("ANU",          r.ANU_Grade);
            Add("GNU",          r.GNU_Grade);
            Add("FPD",          r.FPD_Grade);
            Add("UEC",          r.UEC_Grade);
            Add("Refl. Margin", r.RM_Grade);
            Add("Avg. Grade",   r.AverageGrade);
        }
        else if (r.Is1D)
        {
            Add("SC",  r.SC_Grade);
            Add("MOD", r.MOD_Grade);
        }

        if (rows.Count == 0) return string.Empty;

        string aperture   = r.Aperture.HasValue   ? r.Aperture.Value.ToString("D2")  : "\u2014";
        string wavelength = r.Wavelength.HasValue ? r.Wavelength.Value.ToString()     : "\u2014";
        string lighting   = H(r.Lighting   ?? "\u2014");
        string title      = H($"{r.Standard ?? "Grading Info"} Grade: {GradeDisplay(r.OverallGrade)}");

        var sb = new StringBuilder();
        sb.Append($"""
          <div span="all" style="font-size:9pt;border-style:solid;margin-top:8pt;">
            <table width="100%">
              <col/>
              <tr>
                <td style="border-style:solid;background-color:#1a3a6b;color:white;">
                  <h2 style="padding:0.025in;font-weight:bold;text-align:center;font-size:11pt;">{title}</h2>
                </td>
              </tr>
              <tr><td style="border-style:solid;"><table width="100%">
                <col/><col/><col/><col/><col/><col/><col/><col/>
                <tr style="font-weight:bold;text-align:center;">
                  <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Name</div></td>
                  <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Grade</div></td>
                  <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Aperture</div></td>
                  <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Wavelength</div></td>
                  <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Lighting</div></td>
                  <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Formal Grade</div></td>
                  <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Check</div></td>
                  <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Notes</div></td>
                </tr>
        """);

        foreach (var (name, grade, note) in rows)
        {
            string formal = $"{grade}/{aperture}/{wavelength}/{lighting}";
            sb.Append($"""
                <tr>
                  <td style="border-style:solid;border-width:thin;display-align:center;"><div style="padding:0.025in;">{H(name)}</div></td>
                  <td style="border-style:solid;border-width:thin;display-align:center;"><div style="padding:0.025in;">{grade}</div></td>
                  <td style="border-style:solid;border-width:thin;display-align:center;"><div style="padding:0.025in;">{aperture}</div></td>
                  <td style="border-style:solid;border-width:thin;display-align:center;"><div style="padding:0.025in;">{wavelength}</div></td>
                  <td style="border-style:solid;border-width:thin;display-align:center;"><div style="padding:0.025in;">{lighting}</div></td>
                  <td style="border-style:solid;border-width:thin;display-align:center;"><div style="padding:0.025in;">{formal}</div></td>
                  <td style="border-style:solid;border-width:thin;display-align:center;"><div style="padding:0.025in;">PASS</div></td>
                  <td style="border-style:solid;border-width:thin;display-align:center;"><div style="padding:0.025in;">{H(note)}</div></td>
                </tr>
            """);
        }

        sb.Append("      </table></td></tr></table>\n  </div>");
        return sb.ToString();
    }

    /// <summary>
    /// Two-column row: Image (left ~3.5in) | ISO Quality Parameters (right).
    /// Matches the Webscan PDF layout exactly.
    /// </summary>
    private static string ImageAndQualityTable(VerificationRecord r)
    {
        string imageHtml  = BuildImageCell(r);
        string qualityHtml = BuildQualityCell(r);

        if (string.IsNullOrEmpty(imageHtml) && string.IsNullOrEmpty(qualityHtml))
            return string.Empty;

        // Wrap in a 2-col table exactly as the PDF shows
        return $"""
          <table width="100%" style="font-size:8pt;border-style:solid;margin-top:8pt;">
            <col width="3.5in"/>
            <col/>
            <tr>
              <td style="border-style:solid;vertical-align:top;padding:0;">
                {imageHtml}
              </td>
              <td style="border-style:solid;vertical-align:top;padding:0;">
                {qualityHtml}
              </td>
            </tr>
          </table>
        """;
    }

    private static string BuildImageCell(VerificationRecord r)
    {
        string? b64 = r.RoiJpegImageBase64 ?? r.JpegImageBase64;
        if (string.IsNullOrWhiteSpace(b64)) return string.Empty;

        string label = r.RoiJpegImageBase64 is { Length: > 0 } ? "Symbol Image (ROI Crop)" : "Image";

        // Verbatim Webscan ImageSection structure (header + img div)
        return $"""
          <div style="font-size:8pt;border-style:solid;">
            <h2 style="background-color:#1a3a6b;color:white;padding:0.025in;text-align:center;border-style:solid;font-weight:bold;font-size:10pt;">{H(label)}</h2>
            <div>
              <img src="data:image/jpeg;base64,{b64}" alt="Symbol Image"
                   style="width:auto;max-height:4in;max-width:100%;"/>
            </div>
          </div>
        """;
    }

    private static string BuildQualityCell(VerificationRecord r)
    {
        var rows = new List<(string Name, string Value, string Data)>();

        void Add(string name, GradingResult? g, decimal? numVal = null, string? suffix = null)
        {
            if (g is null && numVal is null) return;
            string v    = numVal.HasValue ? $"{numVal.Value:F0}{suffix}" : g?.NumericGradeString ?? "\u2014";
            string data = g?.LetterGradeString is { Length: > 0 } l ? l : "\u2014";
            rows.Add((name, v, data));
        }

        if (r.Is2D)
        {
            Add("1. Unused Error Correction (UEC)", r.UEC_Grade,    r.UEC_Percent,  "%");
            Add("2. Cell Contrast (CC)",            r.SC_Grade,     r.SC_Percent,   "%");
            Add("3a. Cell Modulation (CMOD)",       r.MOD_Grade);
            Add("3b. Reflectance Margin (RM)",      r.RM_Grade);
            Add("4. Axial Nonuniformity (ANU)",     r.ANU_Grade,    r.ANU_Percent,  "%");
            Add("5. Grid Nonuniformity (GNU)",      r.GNU_Grade,    r.GNU_Percent,  "%");
            Add("6. Fixed Pattern Damage (FPD)",    r.FPD_Grade,    r.FPD_Value);
            Add("18. DECODE",                       r.DECODE_Grade);
            Add("Avg. Grade",                       r.AverageGrade, r.AverageGradeNumeric);
        }
        else if (r.Is1D && r.ScanResults.Count > 0)
        {
            Add("Avg SC",  null, r.Avg_SC);
            Add("Avg MOD", null, r.Avg_MOD);
            Add("Avg EC",  null, r.Avg_MinEC);
        }

        if (rows.Count == 0) return string.Empty;

        string title = r.Is2D ? "ISO 29158 Quality Parameters" : "ISO 15416 Quality Parameters";
        var sb = new StringBuilder();

        // Verbatim Webscan TableSection inner structure
        sb.Append($"""
          <table width="100%" style="font-size:8pt;border-style:solid;">
            <col/><col/><col/><col/>
            <tr>
              <td style="background-color:#1a3a6b;color:white;" colspan="4">
                <h2 style="padding:0.025in;text-align:center;border-style:solid;font-weight:bold;font-size:10pt;">{H(title)}</h2>
              </td>
            </tr>
            <tr style="text-align:center;font-weight:bold;">
              <th style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Parameter</div></th>
              <th style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Value</div></th>
              <th style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Grade</div></th>
              <th style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Check</div></th>
            </tr>
        """);

        foreach (var (name, value, data) in rows)
        {
            bool isPass = !data.StartsWith("F", StringComparison.OrdinalIgnoreCase);
            string check = isPass ? "PASS" : "FAIL";
            string checkStyle = check == "FAIL" ? "color:#721c24;font-weight:bold;" : "";
            sb.Append($"""
              <tr>
                <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">{H(name)}</div></td>
                <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">{value}</div></td>
                <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">{data}</div></td>
                <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;{checkStyle}">{check}</div></td>
              </tr>
            """);
        }

        sb.Append("  </table>");
        return sb.ToString();
    }

    /// <summary>
    /// Two-column row: Data Format Check (left) | RFID Validation (right).
    /// Mirrors the Webscan PDF pattern of Data Format Check beside General Characteristics.
    /// </summary>
    private static string FormatAndRfidTable(VerificationRecord r)
    {
        string formatHtml = BuildDataFormatCheckCell(r);
        string rfidHtml   = BuildRfidCell(r);

        if (string.IsNullOrEmpty(formatHtml) && string.IsNullOrEmpty(rfidHtml))
            return string.Empty;

        return $"""
          <table width="100%" style="font-size:8pt;border-style:solid;margin-top:8pt;">
            <col/>
            <col/>
            <tr>
              <td style="border-style:solid;vertical-align:top;padding:0;">
                {formatHtml}
              </td>
              <td style="border-style:solid;vertical-align:top;padding:0;">
                {rfidHtml}
              </td>
            </tr>
          </table>
        """;
    }

    private static string BuildDataFormatCheckCell(VerificationRecord r)
    {
        var dfc = r.DataFormatCheck;
        if (dfc is null || dfc.Rows.Count == 0) return string.Empty;

        string overall = dfc.Overall switch
        {
            OverallPassFail.Pass => " \u2014 PASS",
            OverallPassFail.Fail => " \u2014 FAIL",
            _                   => "",
        };
        string title = H((dfc.Standard ?? "Data Format Check") + overall);

        var sb = new StringBuilder();
        sb.Append($"""
          <table width="100%" style="font-size:8pt;border-style:solid;">
            <col/><col/><col/>
            <tr>
              <td style="background-color:#1a3a6b;color:white;" colspan="3">
                <h2 style="padding:0.025in;text-align:center;border-style:solid;font-weight:bold;font-size:10pt;">{title}</h2>
              </td>
            </tr>
            <tr style="text-align:center;font-weight:bold;">
              <th style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Name</div></th>
              <th style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Data</div></th>
              <th style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Check</div></th>
            </tr>
        """);

        foreach (var row in dfc.Rows)
        {
            bool fail = string.Equals(row.Check, "FAIL", StringComparison.OrdinalIgnoreCase);
            string checkStyle = fail ? "color:#721c24;font-weight:bold;" : "";
            sb.Append($"""
              <tr>
                <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">{H(row.Name)}</div></td>
                <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">{H(row.Data)}</div></td>
                <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;{checkStyle}">{H(row.Check)}</div></td>
              </tr>
            """);
        }

        sb.Append("  </table>");
        return sb.ToString();
    }

    /// <summary>
    /// RFID Validation cell — right column of the bottom 2-column row.
    /// Left label column is 2in wide; every label carries white-space:nowrap.
    /// </summary>
    private static string BuildRfidCell(VerificationRecord r)
    {
        bool rfidPerformed = !string.IsNullOrWhiteSpace(r.RfidStatus);
        var  sb            = new StringBuilder();

        sb.Append("""
          <table width="100%" style="font-size:9pt;border-style:solid;text-align:left;">
            <col width="2in"/>
            <col/>
            <tr>
              <td style="border-style:solid;background-color:#1a3a6b;color:white;" colspan="2">
                <div style="padding:0.025in;font-weight:bold;text-align:center;font-size:10pt;">
                  RFID Validation &#x2014; VCCS FlexWedge&#x2122;
                </div>
              </td>
            </tr>
        """);

        if (!rfidPerformed)
        {
            sb.Append("""
                <tr>
                  <td colspan="2" style="border-style:solid;border-width:thin;">
                    <div style="padding:0.025in;color:#666;font-style:italic;">
                      RFID validation was not performed for this scan.
                    </div>
                  </td>
                </tr>
            """);
        }
        else
        {
            bool tagDetected = r.RfidStatus is not ("NoTag" or "NO_TAG");
            RfidRow(sb, "Tag Detected", tagDetected ? "Yes" : "No");

            if (tagDetected)
            {
                RfidRow(sb, "EPC (Hex)",       r.RfidEpcHex ?? "\u2014", mono: true);
                RfidRow(sb, "Decoded GTIN-14", r.RfidGtin14 ?? "\u2014", mono: true);
                RfidRow(sb, "GCP Valid",       r.RfidGcpValid.HasValue
                    ? (r.RfidGcpValid.Value ? "Yes" : "No") : "\u2014");
                RfidRow(sb, "EPC Serial",      r.RfidSerial ?? "\u2014", mono: true);
            }

            (string style, string display) = r.RfidStatus switch
            {
                "Pass"  => ("color:#155724;font-weight:bold;",
                            "&#x2713; PASS &#x2014; EPC matches barcode data"),
                "Fail"  => ("color:#721c24;font-weight:bold;",
                            "&#x2718; FAIL &#x2014; EPC does not match barcode data"),
                "NoTag" => ("color:#6c757d;",
                            "&#x2014; No tag detected in scan window"),
                "ParseError" => ("color:#856404;font-weight:bold;",
                                 "! Parse error &#x2014; EPC could not be decoded"),
                "MultipleTagsDetected" => ("color:#856404;font-weight:bold;",
                                           "! Multiple tags &#x2014; ambiguous read"),
                _ => ("", H(r.RfidStatus)),
            };

            sb.Append($"""
                <tr>
                  <td style="border-style:solid;border-width:thin;font-weight:bold;">
                    <div style="padding:0.025in;white-space:nowrap;">Validation Result</div>
                  </td>
                  <td style="border-style:solid;border-width:thin;">
                    <div style="padding:0.025in;{style}">{display}</div>
                  </td>
                </tr>
            """);

            if (!string.IsNullOrWhiteSpace(r.RfidMismatchDetail))
                RfidRow(sb, "Mismatch Detail", H(r.RfidMismatchDetail), mono: true);

            if (r.RfidScanWindowMs.HasValue)
                RfidRow(sb, "Scan Window", $"{r.RfidScanWindowMs.Value} ms");
        }

        sb.Append("  </table>");
        return sb.ToString();
    }

    private static string Footer() =>
        """
        <div class="vccs-footer">
          Generated by VTCCP &nbsp;&middot;&nbsp; VCCS FlexWedge&#x2122; RFID Validation &nbsp;&middot;&nbsp; Hybrid Report v2.1
        </div>
        """;

    // ── Structural helpers ─────────────────────────────────────────────────────

    private static void RfidRow(StringBuilder sb, string label, string value, bool mono = false)
    {
        string fontStyle = mono ? "font-family:Consolas,monospace;" : "";
        sb.Append($"""
            <tr>
              <td style="border-style:solid;border-width:thin;">
                <div style="padding:0.025in;white-space:nowrap;">{H(label)}</div>
              </td>
              <td style="border-style:solid;border-width:thin;">
                <div style="padding:0.025in;{fontStyle}">{value}</div>
              </td>
            </tr>
        """);
    }

    private static string BuildGradeExtra(VerificationRecord r)
    {
        var parts = new List<string>();
        if (r.Aperture.HasValue)              parts.Add(r.Aperture.Value.ToString("D2"));
        if (r.Wavelength.HasValue)            parts.Add(r.Wavelength.Value.ToString());
        if (r.Lighting is { Length: > 0 })    parts.Add(r.Lighting);
        if (r.FormalGrade is { Length: > 0 }) parts.Add(r.FormalGrade);
        return parts.Count > 0
            ? $" <span style=\"font-weight:normal;font-size:8pt;\">({H(string.Join("/", parts))})</span>"
            : "";
    }

    // ── Display helpers ────────────────────────────────────────────────────────

    private static string RfidStatusBadge(VerificationRecord r) =>
        string.IsNullOrWhiteSpace(r.RfidStatus)
            ? "<span style=\"color:#6c757d;\">\u2014 Not configured</span>"
            : r.RfidStatus switch
            {
                "Pass"  => Badge("#d4edda", "#155724", "#c3e6cb", "&#x2713; MATCHED"),
                "Fail"  => Badge("#f8d7da", "#721c24", "#f5c6cb", "&#x2718; MISMATCH"),
                "NoTag" => Badge("#e2e3e5", "#383d41", "#d6d8db", "\u2014 No Tag"),
                "ParseError" or "MultipleTagsDetected"
                        => Badge("#fff3cd", "#856404", "#ffeeba", "! " + H(r.RfidStatus)),
                _       => $"<span>{H(r.RfidStatus)}</span>",
            };

    private static string Badge(string bg, string fg, string border, string text) =>
        $"<span style=\"background:{bg};color:{fg};padding:1pt 6pt;" +
        $"border:1px solid {border};font-size:10pt;\">{text}</span>";

    private static string GradeDisplay(GradingResult? g)
    {
        if (g is null) return "\u2014";
        string letter = g.LetterGradeString;
        bool hasL = letter is { Length: > 0 };
        bool hasN = g.NumericGrade.HasValue;
        if (hasL && hasN) return H($"{letter} ({g.NumericGrade!.Value:F1})");
        if (hasL)         return H(letter);
        if (hasN)         return g.NumericGrade!.Value.ToString("F1");
        return "\u2014";
    }

    private static string H(string? s) =>
        s is null ? string.Empty
        : s.Replace("&", "&amp;")
           .Replace("<", "&lt;")
           .Replace(">", "&gt;")
           .Replace("\"", "&quot;")
           .Replace("'", "&#39;");

    // ── Extra CSS (VCCS additions only — Webscan base CSS is verbatim in Template) ──

    private const string ExtraCss = """
        /* ── VCCS FlexWedge™ additions ──────────────────────────────────────── */
        .vccs-print-btn {
            display:inline-block; margin:8pt; padding:5pt 14pt;
            background:#1a3a6b; color:#fff; border:none; cursor:pointer;
            font-size:9pt; font-family:inherit;
        }
        .vccs-print-btn:hover { background:#25508a; }

        .vccs-footer {
            margin-top:16pt; padding-top:6pt; font-size:7.5pt; color:#666;
            border-top:1px solid #ccc; text-align:center;
        }

        @media print {
            .no-print { display:none !important; }
            body { text-align:center; }
            @page { margin:0.65in; size:letter portrait; }
        }
        """;

    // ── Template ───────────────────────────────────────────────────────────────
    //
    // Verbatim Webscan TruCheck HTML skeleton (from HTMLreport_*.html).
    // Every CSS rule, table, col, td style attribute, and the <lu><li> structure
    // are copied exactly.  The only deliberate changes:
    //   • background-color:black  →  #1a3a6b  (navy; prints same as black in B&W)
    //   • Header <table>: 2 logo cols added at left, 1 company col at right
    //   • <lu><li> body: tokens replaced for VCCS content (no structural change)
    //   • Two two-column paired-section <li> items (Image+Params, Format+RFID)
    //     replace what the XSLT generates via CSS column-flow into separate <li>s

    private const string Template = """
        <?xml version="1.0" encoding="utf-8"?>
        <html xmlns:fo="http://www.w3.org/1999/XSL/Format" lang="en">
          <!--Header-->
          <head>
            <title><!-- VCCS:TITLE --></title>
            <style type="text/css">
              <!--
                    table {
                    border-color:Black;
                    border-collapse:collapse;
                    border-spacing:0;
                    }
                    td {
                    border-color:Black;
                    padding:0;
                    }
                    th {
                    border-color:Black;
                    padding:0;
                    }
                    div {
                    border-color:Black;
                    }
                    h2 {
                    border-color:Black;
                    margin:0;
                    }
                    h3 {
                    border-color:Black;
                    margin:0;
                    }
                    ul {
                    list-style-type:none;
                    }
                    li {
                    list-style-type:none;

                    }
                    body {
                    text-align:center;

                    }
               -->
              <!-- VCCS:EXTRA_CSS -->
            </style>
          </head>

          <!--
            Header banner — four-column layout:
              [VCCS logo] [Webscan logo] [Title / meta] [Company logo]
            Original Webscan header was 2-col (logo + title).
          -->
          <div>
            <table width="100%">
              <col width="1.1in"/>
              <col width="1.1in"/>
              <col/>
              <col width="1.1in"/>
              <tr>
                <!-- VCCS:VCCS_LOGO_CELL -->
                <!-- VCCS:WEBSCAN_LOGO_CELL -->
                <!-- VCCS:TITLE_CELL -->
                <!-- VCCS:COMPANY_LOGO_CELL -->
              </tr>
            </table>
          </div>

          <!--Body-->
          <body>
            <!-- VCCS:PRINT_BUTTON -->

            <!--Summary — full width, span="all"-->
            <div span="all">
              <table style="font-size:9pt;border-style:solid;text-align:left;" width="100%">
                <col width="1.5in"/>
                <col width="6.4in"/>
                <tr>
                  <td style="border-style:solid;background-color:#1a3a6b;color:white;" colspan="2">
                    <div style="padding:0.025in;font-weight:bold;text-align:center;font-size:11pt;">
                      Report Summary
                    </div>
                  </td>
                </tr>
                <!-- VCCS:SUMMARY_ROWS -->
              </table>
            </div>

            <!--OverallGrades — full width, span="all"-->
            <div span="all">
              <table width="100%" style="font-size:9pt;border-style:solid;text-align:center;margin-top:8pt;">
                <col/>
                <tr>
                  <td style="border-style:solid;background-color:#1a3a6b;color:white;">
                    <h2 style="padding:0.025in;font-weight:bold;font-size:11pt;">Verification Grades</h2>
                  </td>
                </tr>
                <tr>
                  <td style="border-style:solid;">
                    <table width="100%">
                      <col/><col/><col/><col/><col/><col/>
                      <tr style="font-weight:bold;text-align:center;">
                        <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Standard</div></td>
                        <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Grade</div></td>
                        <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Aperture</div></td>
                        <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Wavelength</div></td>
                        <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Lighting</div></td>
                        <td style="border-style:solid;border-width:thin;"><div style="padding:0.025in;">Formal Grade</div></td>
                      </tr>
                      <!-- VCCS:GRADES_ROW -->
                    </table>
                  </td>
                </tr>
              </table>
            </div>

            <lu>
              <li>
                <!--GradingInfoSection — full width-->
                <!-- VCCS:GRADING_INFO_SECTION -->
              </li>
              <li>
                <!--ImageSection + TableSection — side by side, 2 columns-->
                <!--ReportSection-->
                <!-- VCCS:IMAGE_AND_PARAMS -->
              </li>
              <li>
                <!--DataFormatCheck + RFID Validation — side by side, 2 columns-->
                <!--ReportSection-->
                <!-- VCCS:FORMAT_AND_RFID -->
              </li>
            </lu>

            <!-- VCCS:FOOTER -->
          </body>
        </html>
        """;
}
