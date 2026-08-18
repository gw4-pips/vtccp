// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  VCCS RFID Validation Report — HTML generator (v23-faithful)             ║
// ║                                                                          ║
// ║  Produces a fully self-contained HTML document matching the canonical    ║
// ║  v23 design (dist/vccs-pdf-preview-v23.html) character-for-character in  ║
// ║  CSS, layout and section structure.  Real VerificationRecord data is     ║
// ║  substituted via section builders (same pattern as HybridReportGenerator)║
// ║  and both logos (VCCS + PIPS) are base64-inlined so the file renders     ║
// ║  identically with no external references.                                ║
// ║                                                                          ║
// ║  The HTML is rendered to PDF by VccsPdfRenderer (WebView2 primary,       ║
// ║  wkhtmltopdf silent fallback).                                           ║
// ╚══════════════════════════════════════════════════════════════════════════╝

using System.Text;
using ExcelEngine.Models;

namespace DeviceInterface.Reports;

/// <summary>
/// Generates the v23-design VCCS FlexWedge™ Pro RFID Validation Report as a
/// self-contained HTML string.  All CSS is verbatim from the canonical v23
/// preview file; only data values differ.
/// </summary>
public static class VccsHtmlReportGenerator
{
    /// <summary>Report format version — bump on ANY layout/content/logic change.</summary>
    public const string ReportVersion = "v1.4.12";

    // ── Public API ─────────────────────────────────────────────────────────

    public static string Generate(VerificationRecord r)
    {
        var sb = new StringBuilder(96 * 1024);

        sb.Append("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\"/>\n");
        sb.Append("<title>VCCS FlexWedge\u2122 Pro RFID Validation Report</title>\n");
        sb.Append("<style>\n").Append(Css).Append("\n</style>\n</head>\n<body>\n\n");

        sb.Append("<div class=\"page\">\n\n");
        BuildHeader(sb, r);
        sb.Append("  <div class=\"content\">\n\n");
        BuildBarcodeSummarySection(sb, r);
        BuildRfidSection(sb, r);
        BuildImageSection(sb, r);
        BuildDataFormatSection(sb, r);
        sb.Append("  </div><!-- /content -->\n\n");
        BuildFooter(sb, r);
        sb.Append("</div><!-- /page -->\n</body>\n</html>\n");

        return sb.ToString();
    }

    // ── Header ─────────────────────────────────────────────────────────────

    private static void BuildHeader(StringBuilder sb, VerificationRecord r)
    {
        string dt     = r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt");
        string device = H(r.DeviceModel ?? r.VerifierBrand ?? "\u2014");
        string serial = H(r.DeviceSerial ?? "\u2014");
        string sw     = H(r.SoftwareVersion ?? "\u2014");
        string fw     = H(r.FirmwareVersion ?? "\u2014");

        (string badgeCls, string badgeTxt) = r.RfidStatus switch
        {
            "Pass"                 => ("badge-pass", "&#x2713; RFID MATCHED"),
            "Fail"                 => ("badge-fail", "&#x2717; RFID MISMATCH"),
            "NoTag"                => ("badge-warn", "&#x26a0; NO RFID TAG"),
            "MultipleTagsDetected" => ("badge-warn", "&#x26a0; MULTIPLE TAGS"),
            "Skipped"              => ("badge-warn", "&#x2014; RFID SKIPPED"),
            null or ""             => ("badge-warn", "&#x2014; NO RFID DATA"),
            var s                  => ("badge-warn", H(s.ToUpperInvariant())),
        };

        // VCCS logo — base64-inlined; text fallback when the PNG is absent.
        string? vccsB64 = LoadLogoBase64("vccs_logo.png");
        string vccsInner = vccsB64 is not null
            ? $"      <img src=\"data:image/png;base64,{vccsB64}\" style=\"max-height:65pt;max-width:68pt;object-fit:contain;\" alt=\"VCCS\" />\n"
            : "      <div class=\"logo-name\">VCCS</div>\n      <div class=\"logo-sub\">FlexWedge&#x2122; Pro</div>\n";

        // Company logo — session logo (LogoPath) preferred, PIPS logo as default;
        // company-name text as last resort.
        string? companyB64 = LoadLogoBase64FromPath(r.LogoPath) ?? LoadLogoBase64("pips_logo.png");
        string companyInner = companyB64 is not null
            ? $"      <img src=\"data:image/png;base64,{companyB64}\" style=\"max-height:48pt;max-width:68pt;object-fit:contain;\" alt=\"{H(r.CompanyName ?? "Company")}\" />\n"
            : $"      {H(r.CompanyName ?? "Company Logo")}\n";

        sb.Append($"""
          <!-- ── HEADER ─────────────────────────────────────────────── -->
          <div class="header">
            <div class="logo-box" style="padding:3pt 2pt;">
        {vccsInner}    </div>
            <div class="header-meta">
              <div class="dt">{H(dt)}</div>
              <div class="dev">Device: {device}</div>
              <div class="ln">Serial: {serial}</div>
              <div class="ln">Software: {sw}</div>
              <div class="ln">Firmware: {fw}</div>
            </div>
            <div class="header-title">
              <h1>VCCS <em>FlexWedge&#x2122; Pro</em><br>RFID Validation Report</h1>
              <span class="rfid-badge {badgeCls}">{badgeTxt}</span>
            </div>
            <div class="company-box" style="display:flex;align-items:center;justify-content:center;padding:2pt 4pt;">
        {companyInner}    </div>
          </div>


        """);
    }

    // ── ① Barcode Verification Results Summary + Grades ────────────────────

    private static void BuildBarcodeSummarySection(StringBuilder sb, VerificationRecord r)
    {
        string? brand = r.VerifierBrand ?? (r.DeviceModel is { Length: > 0 } ? "COGNEX" : null);
        string? dataManInfix = (brand == "COGNEX" &&
            r.DeviceModel?.Contains("DataMan", StringComparison.OrdinalIgnoreCase) == true)
            ? "DataMan " : null;
        string sectionTitle = brand != null
            ? $"{H(brand)} {dataManInfix}TruCheck Barcode Verification Results Summary"
            : "TruCheck Barcode Verification Results Summary";

        bool multiMode = !string.IsNullOrWhiteSpace(r.LinearSymbology);

        string reportName = !string.IsNullOrWhiteSpace(r.WebscanSourcePath)
            ? Path.GetFileName(r.WebscanSourcePath)
            : $"{r.VerificationDateTime:yyyy-MM-dd_HH-mm-ss}_vccs_rfid.pdf";

        sb.Append($"""
            <!-- ① Barcode Verification Results Summary -->
            <div>
              <div class="sec-sub-hdr">{sectionTitle}<span class="sec-note"> &#x2014; <em>See associated TruCheck report for additional details</em></span></div>
              <table class="sum-table">
                <colgroup>
                  <col style="width:76pt">
                  <col>
                  <col style="width:88pt">
                </colgroup>
                <thead>
                  <tr>
                    <th>Symbology</th>
                    <th>Encoded Data</th>
                    <th>Application Spec.</th>
                  </tr>
                </thead>
                <tbody>

        """);

        if (multiMode)
            AppendSymbolRow(sb, r.LinearSymbology!, r.LinearDecodedData, LinearAppSpecCell(r.LinearSymbology));
        AppendSymbolRow(sb, r.Symbology ?? "\u2014", r.DecodedData, AppSpecCell(r));

        sb.Append($"""
                  <tr class="sum-meta-start">
                    <td class="meta-lbl">Report Name</td>
                    <td colspan="2">{H(reportName)}</td>
                  </tr>
                  <tr>
                    <td class="meta-lbl">Report Date/Time</td>
                    <td colspan="2">{H(r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt"))}</td>
                  </tr>
                </tbody>
              </table>

              <div class="sec-sub-hdr sec-sub-hdr-gap">Barcode Verification Grades</div>
              <table class="grades-table">
                <thead>
                  <tr>
                    <th>Symbology</th>
                    <th>Standard</th>
                    <th>Grade</th>
                    <th>Aperture (mil)</th>
                    <th>Wavelength (nm)</th>
                    <th>Lighting</th>
                    <th>Formal Grade</th>
                  </tr>
                </thead>
                <tbody>

        """);

        if (multiMode)
        {
            AppendGradeRow(sb, r.LinearSymbology!, r.LinearStandard ?? "ISO/IEC 15416",
                GradeDisplay(r.LinearOverallGrade),
                r.LinearAperture.HasValue ? r.LinearAperture.Value.ToString("D2") : "\u2014",
                r.LinearWavelength?.ToString() ?? "\u2014",
                r.LinearLighting ?? "\u2014", r.LinearFormalGrade ?? "\u2014");
        }
        AppendGradeRow(sb, r.Symbology ?? "\u2014", r.Standard ?? "\u2014",
            GradeDisplay(r.OverallGrade),
            r.Aperture.HasValue ? r.Aperture.Value.ToString("D2") : "\u2014",
            r.Wavelength?.ToString() ?? "\u2014",
            r.Lighting ?? "\u2014", r.FormalGrade ?? "\u2014");

        sb.Append("""
                </tbody>
              </table>
            </div>


        """);
    }

    private static void AppendSymbolRow(StringBuilder sb, string symb, string? encoded, string appSpecCell)
    {
        sb.Append($"""
                  <tr>
                    <td style="font-size:8pt;">{H(symb)}</td>
                    <td style="font-family:Consolas,monospace;">{H(encoded ?? "\u2014")}</td>
                    {appSpecCell}
                  </tr>

        """);
    }

    // 2D symbols use the flex two-liner app-spec cell ("GS1 —" | stacked pair);
    // linear symbols use the single-line form — both verbatim v23 structures.
    private static string AppSpecCell(VerificationRecord r)
    {
        string? s = r.Symbology;
        if (s?.Contains("DataMatrix", StringComparison.OrdinalIgnoreCase) == true)
            return AppSpecStackCell("Element", "String");
        if (s?.Contains("QR", StringComparison.OrdinalIgnoreCase) == true)
            return AppSpecStackCell("Digital", "Link");
        string fallback = !string.IsNullOrWhiteSpace(r.ApplicationStandard)
            ? r.ApplicationStandard : r.Standard ?? "\u2014";
        return $"<td style=\"font-size:8pt;\">{H(fallback)}</td>";
    }

    private static string AppSpecStackCell(string line1, string line2) => $"""
        <td class="app-spec">
                      <div class="app-spec-inner">
                        <span class="app-spec-prefix">GS1 &#x2014;</span>
                        <div class="app-spec-stack">{H(line1)}<br>{H(line2)}</div>
                      </div>
                    </td>
        """;

    private static string LinearAppSpecCell(string? symbology)
    {
        string txt = symbology switch
        {
            var s when s?.Contains("UPC",     StringComparison.OrdinalIgnoreCase) == true => "GS1 \u2014 GTIN-12",
            var s when s?.Contains("EAN-8",   StringComparison.OrdinalIgnoreCase) == true => "GS1 \u2014 GTIN-8",
            var s when s?.Contains("EAN",     StringComparison.OrdinalIgnoreCase) == true => "GS1 \u2014 GTIN-13",
            var s when s?.Contains("DataBar", StringComparison.OrdinalIgnoreCase) == true => "GS1 \u2014 GTIN-14",
            _ => "GS1",
        };
        return $"<td style=\"font-size:8pt;\">{H(txt)}</td>";
    }

    private static void AppendGradeRow(StringBuilder sb, string symb, string standard,
        string grade, string aperture, string wavelength, string lighting, string formal)
    {
        sb.Append($"""
                  <tr>
                    <td>{H(symb)}</td>
                    <td>{H(standard)}</td>
                    <td>{H(grade)}</td>
                    <td>{H(aperture)}</td>
                    <td>{H(wavelength)}</td>
                    <td>{H(lighting)}</td>
                    <td>{H(formal)}</td>
                  </tr>

        """);
    }

    // ── ② VCCS FlexWedge™ Pro EPC RFID Validation Summary ──────────────────

    private static void BuildRfidSection(StringBuilder sb, VerificationRecord r)
    {
        bool isGS1Rfid = r.ApplicationStandard?
            .StartsWith("GS1", StringComparison.OrdinalIgnoreCase) != false;
        string rfidAdj = isGS1Rfid ? "EPC" : "UHF";

        bool tagDetected = r.RfidStatus is "Pass" or "Fail" or "MultipleTagsDetected";
        string tagLabel = r.RfidStatus switch
        {
            "Pass" or "Fail"       => "Yes",
            "NoTag"                => "No",
            "MultipleTagsDetected" => "Yes (multiple)",
            "Skipped"              => "Skipped",
            _                      => r.RfidStatus ?? "\u2014",
        };
        string lockDisplay = !tagDetected ? "N/A" : r.RfidTagLockStatus switch
        {
            "Locked"      => "Locked",
            "PermaLocked" => "Permanently Locked",
            "Unlocked"    => "Unlocked",
            "Unknown"     => "Unknown",
            null          => "\u2014",
            var other     => other,
        };

        // EPC encoding scheme — derived from Tag URI (urn:epc:tag:<scheme>:…).
        string? epcScheme = null;
        if (!string.IsNullOrWhiteSpace(r.RfidEpcTagUri))
        {
            var p = r.RfidEpcTagUri.Split(':');
            if (p.Length >= 4 && p[0] == "urn" && p[1] == "epc" && p[2] == "tag")
                epcScheme = p[3].ToUpperInvariant();
        }
        string schemePart = tagDetected ? (epcScheme ?? "\u2014") : "N/A";

        string gcpLenPart = r.RfidGcpLength.HasValue ? $" ({r.RfidGcpLength.Value})" : string.Empty;
        string gcpDisplay = r.RfidGcpValid switch
        {
            true  => $"Valid{gcpLenPart}",
            false => $"Invalid{gcpLenPart}",
            null  => "\u2014",
        };
        string gcpNote = !string.IsNullOrWhiteSpace(r.RfidGcpTableDate)
            ? $"<em class=\"gcp-inline-note\"> &ndash; From GCP prefix table as of {H(r.RfidGcpTableDate)}</em>"
            : string.Empty;

        sb.Append($$"""
            <!-- ② VCCS FlexWedge™ Pro EPC RFID Validation Summary -->
            <div>
              <div class="sec-hdr">VCCS <em>FlexWedge&#x2122; Pro</em> {{rfidAdj}} RFID Validation Summary</div>
              <table class="rfid-table">
                <thead>
                  <tr><th class="lbl-col">Field</th><th>Value</th></tr>
                </thead>
                <tbody>
                  <tr>
                    <td>Tag Detected / Lock Status</td>
                    <td>{{H(tagLabel)}} &#x2014; {{H(lockDisplay)}} <span style="color:#555;font-size:8pt;">(Permalocked / Locked / Unlocked / Unknown)</span></td>
                  </tr>
                  <tr class="row-hi">
                    <td>EPC Encoding Scheme / GCP Length</td>
                    <td><span style="font-family:Consolas,monospace;">{{H(schemePart)}}</span> &#x2014; {{H(gcpDisplay)}}{{gcpNote}}</td>
                  </tr>
                  <tr class="row-hi">
                    <td>EPC Hex</td>
                    <td style="font-family:Consolas,monospace;">{{H(r.RfidEpcHex ?? "\u2014")}}</td>
                  </tr>
                  <tr class="row-hi">
                    <td>EPC Tag URI</td>
                    <td style="font-family:Consolas,monospace;font-size:8pt;">{{H(r.RfidEpcTagUri ?? "\u2014")}}</td>
                  </tr>
                  <tr class="row-hi">
                    <td>GTIN-14</td>
                    <td style="font-family:Consolas,monospace;">{{H(r.RfidGtin14 ?? "\u2014")}}</td>
                  </tr>
                  <tr class="row-hi">
                    <td>Serial Number</td>
                    <td style="font-family:Consolas,monospace;">{{H(r.RfidSerial ?? "\u2014")}}</td>
                  </tr>

        """);

        // Result rows — named by the actual symbology(ies) in this scan.
        // Multi-mode: linear row (GTIN only) + 2D row (GTIN + Serial Number).
        string rowCls = r.RfidStatus switch
        {
            "Pass" => "row-result-pass",
            "Fail" => "row-result-fail",
            _      => "row-result-warn",
        };

        bool multiMode = !string.IsNullOrWhiteSpace(r.LinearSymbology);
        bool is1DOnly  = r.Is1D && !multiMode;

        void ResultRow(string label, string value) =>
            sb.Append($"""
                      <tr class="{rowCls}">
                        <td>{H(label)}</td>
                        <td>{value}</td>
                      </tr>

            """);

        if (multiMode)
        {
            string linVal = r.RfidStatus switch
            {
                "Pass" => "Pass &#x2014; EPC data matches barcode GTIN",
                "Fail" => "Fail &#x2014; GTIN mismatch",
                var s  => H(s ?? "\u2014"),
            };
            ResultRow($"{r.LinearSymbology} Validation Result", linVal);

            string twoDSym = string.IsNullOrWhiteSpace(r.Symbology) ? "2D" : r.Symbology;
            string twoDVal = r.RfidStatus switch
            {
                "Pass" => "Pass &#x2014; EPC data matches barcode GTIN and Serial Number",
                "Fail" => "Fail &#x2014; GTIN or Serial Number mismatch",
                var s  => H(s ?? "\u2014"),
            };
            if (!string.IsNullOrWhiteSpace(r.RfidMismatchDetail) && r.RfidStatus is "Fail")
                twoDVal += $" &#x2014; {H(r.RfidMismatchDetail)}";
            ResultRow($"{twoDSym} Validation Result", twoDVal);
        }
        else
        {
            string symName = string.IsNullOrWhiteSpace(r.Symbology)
                ? (is1DOnly ? "EAN/UPC" : "2D Symbol") : r.Symbology;
            string singleVal = r.RfidStatus switch
            {
                "Pass" when is1DOnly => "Pass &#x2014; EPC data matches barcode GTIN",
                "Fail" when is1DOnly => "Fail &#x2014; GTIN mismatch",
                "Pass"               => "Pass &#x2014; EPC data matches barcode GTIN and Serial Number",
                "Fail"               => "Fail &#x2014; GTIN or Serial Number mismatch",
                var s                => H(s ?? "\u2014"),
            };
            if (!string.IsNullOrWhiteSpace(r.RfidMismatchDetail) && r.RfidStatus is "Fail")
                singleVal += $" &#x2014; {H(r.RfidMismatchDetail)}";
            ResultRow($"{symName} Validation Result", singleVal);
        }

        sb.Append("""
                </tbody>
              </table>
            </div>


        """);
    }

    // ── ③ Barcode Image ─────────────────────────────────────────────────────

    private static void BuildImageSection(StringBuilder sb, VerificationRecord r)
    {
        string? img2D     = r.RoiJpegImageBase64 ?? r.JpegImageBase64;
        string? imgLinear = r.LinearJpegImageBase64;
        bool multiMode    = !string.IsNullOrWhiteSpace(r.LinearSymbology);

        sb.Append("""
            <!-- ③ Barcode Image -->
            <div>

        """);

        if (multiMode && !string.IsNullOrWhiteSpace(img2D) && !string.IsNullOrWhiteSpace(imgLinear))
        {
            // Dual layout: linear ROI frame left, 2D crop right.
            string img2DForDual = r.JpegImageBase64 ?? img2D;
            sb.Append($"""
                  <div class="sec-sub-hdr">Barcode Images</div>
                  <div class="img-frame">
                    <table style="width:100%;border-collapse:collapse;"><tr>
                      <td style="width:50%;text-align:center;vertical-align:middle;">
                        <div style="font-size:7pt;color:#6c757d;font-style:italic;">{H(r.LinearSymbology!)}</div>
                        <img src="data:image/jpeg;base64,{imgLinear}" alt="{H(r.LinearSymbology!)}" style="max-width:3.4in;max-height:1.6in;object-fit:contain;"/>
                      </td>
                      <td style="width:50%;text-align:center;vertical-align:middle;">
                        <div style="font-size:7pt;color:#6c757d;font-style:italic;">{H(r.Symbology ?? "2D")}</div>
                        <img src="data:image/jpeg;base64,{img2DForDual}" alt="{H(r.Symbology ?? "2D")}" style="max-width:3.4in;max-height:1.6in;object-fit:contain;"/>
                      </td>
                    </tr></table>
                  </div>

            """);
        }
        else if (!string.IsNullOrWhiteSpace(img2D))
        {
            sb.Append($"""
                  <div class="sec-sub-hdr">Barcode Image</div>
                  <div class="img-frame">
                    <img src="data:image/jpeg;base64,{img2D}" alt="Barcode" style="max-width:100%;max-height:1.8in;object-fit:contain;"/>
                  </div>

            """);
        }
        else
        {
            sb.Append("""
                  <div class="sec-sub-hdr">Barcode Image</div>
                  <div class="img-frame">
                    <div class="img-placeholder">
                      [No barcode image available for this scan]
                    </div>
                  </div>

            """);
        }

        sb.Append("""
            </div>


        """);
    }

    // ── ④ Data Format Check ─────────────────────────────────────────────────

    private static void BuildDataFormatSection(StringBuilder sb, VerificationRecord r)
    {
        bool has2D     = r.DataFormatCheck       is { Rows.Count: > 0 };
        bool hasLinear = r.LinearDataFormatCheck is { Rows.Count: > 0 };
        if (!has2D && !hasLinear) return;

        // Overall = fail if either check fails.
        bool anyFail =
            (has2D     && r.DataFormatCheck!.Overall       == OverallPassFail.Fail) ||
            (hasLinear && r.LinearDataFormatCheck!.Overall == OverallPassFail.Fail);
        string pillCls = anyFail ? "pill-fail" : "pill-pass";
        string pillTxt = anyFail ? "OVERALL: FAIL" : "OVERALL: PASS";

        string std = r.DataFormatCheck?.Standard
                  ?? r.LinearDataFormatCheck?.Standard ?? "GS1";
        string hdrSuffix = std.Contains("GS1", StringComparison.OrdinalIgnoreCase) ? "GS1" : std;

        sb.Append($"""
            <!-- ④ Data Format Check -->
            <div class="cf">
              <div class="sec-sub-hdr">Data Format Check &#x2014; {H(hdrSuffix)}<span class="sec-note"> &#x2014; <em>See associated TruCheck report for additional details</em></span></div>
              <table class="dfc-table">
                <thead>
                  <tr><th>Field</th><th>Data</th><th class="chk">Check</th></tr>
                </thead>
                <tbody>

        """);

        void DfcRows(DataFormatCheckResult dfc, string? symbLabel)
        {
            if (symbLabel is not null)
                sb.Append($"""
                          <tr><td colspan="3" style="background:#e8eef5;font-weight:bold;">{H(symbLabel)}</td></tr>

                """);
            foreach (var row in dfc.Rows)
            {
                bool fail = string.Equals(row.Check, "FAIL", StringComparison.OrdinalIgnoreCase);
                string cls = fail ? "chk fail-fg" : "chk pass-fg";
                bool mono = row.Name.Contains("GTIN", StringComparison.OrdinalIgnoreCase);
                string dataStyle = mono ? " style=\"font-family:Consolas,monospace;\"" : "";
                sb.Append($"""
                          <tr><td>{H(row.Name)}</td><td{dataStyle}>{H(row.Data)}</td><td class="{cls}">{H(row.Check)}</td></tr>

                """);
            }
        }

        bool bothPresent = has2D && hasLinear;
        if (hasLinear) DfcRows(r.LinearDataFormatCheck!, bothPresent ? r.LinearSymbology : null);
        if (has2D)     DfcRows(r.DataFormatCheck!,       bothPresent ? r.Symbology      : null);

        sb.Append($"""
                </tbody>
              </table>
              <span class="overall-pill {pillCls}">{pillTxt}</span>
            </div>


        """);
    }

    // ── Footer ──────────────────────────────────────────────────────────────

    private static void BuildFooter(StringBuilder sb, VerificationRecord r)
    {
        sb.Append($"""
          <!-- ── FOOTER ─────────────────────────────────────────────── -->
          <div class="footer">
            <div class="fl">
              VCCS <em>FlexWedge&#x2122; Pro</em> RFID Validation Report {ReportVersion}
              <span>&nbsp;&#x2014;&nbsp; Generated {DateTime.Now:yyyy-MM-dd HH:mm:ss}</span>
            </div>
            <div>Verification Command &amp; Control System (VCCS)</div>
          </div>


        """);
    }

    // ── Logo loading ─────────────────────────────────────────────────────────

    // Logos live at <ExeDir>/resources/ (copied by VtccpApp.csproj Content rules).
    // Cached per file name — loaded once per process.
    private static readonly Dictionary<string, string?> _logoCache = new();
    private static readonly object _logoLock = new();

    private static string? LoadLogoBase64(string fileName)
    {
        lock (_logoLock)
        {
            if (_logoCache.TryGetValue(fileName, out var cached)) return cached;
            string? b64 = null;
            try
            {
                string? dir = Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (dir is not null)
                {
                    string p = Path.Combine(dir, "resources", fileName);
                    if (File.Exists(p))
                        b64 = Convert.ToBase64String(File.ReadAllBytes(p));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VCCS-HTML] Logo load failed for '{fileName}': {ex.Message}");
            }
            _logoCache[fileName] = b64;
            return b64;
        }
    }

    private static string? LoadLogoBase64FromPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
        try   { return Convert.ToBase64String(File.ReadAllBytes(path)); }
        catch { return null; }
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static string GradeDisplay(GradingResult? g)
    {
        if (g is null) return "\u2014";
        string letter = g.LetterGradeString;
        bool hasL = letter is { Length: > 0 };
        bool hasN = g.NumericGrade.HasValue;
        if (hasL && hasN) return $"{letter} ({g.NumericGrade!.Value:F1})";
        if (hasL)         return letter;
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

    // ── CSS — verbatim from dist/vccs-pdf-preview-v23.html ───────────────────
    // Only the .print-hint / .design-note preview-chrome rules are retained
    // unchanged; they are inert because the generated document has no such
    // elements.  DO NOT edit values here without a report-version bump.

    private const string Css = """
      * { box-sizing: border-box; margin: 0; padding: 0; }

      table th, table td { vertical-align: middle; }

      body {
        background: #888;
        font-family: Arial, sans-serif;
        font-size: 9pt;
        padding: 24px;
      }

      .print-hint {
        text-align: center; margin-bottom: 12px;
        font-family: Arial, sans-serif; font-size: 11px; color: #ccc;
      }
      .print-hint button {
        background: #1a3a6b; color: white; border: none;
        padding: 6px 16px; border-radius: 3px; cursor: pointer; font-size: 11px;
      }

      .page {
        width: 8.5in; height: 11in; overflow: hidden; background: white;
        margin: 0 auto; padding: 0.25in 0.5in 0.5in 0.5in;
        box-shadow: 0 4px 24px rgba(0,0,0,0.4);
        display: flex; flex-direction: column;
      }

      /* ── HEADER ─────────────────────────────────────────── */
      .header {
        background: #edf1f7; border-bottom: 2px solid #1a3a6b;
        padding: 6pt 0; display: flex; flex-direction: row;
        align-items: stretch; gap: 0; margin-bottom: 0;
      }
      .logo-box {
        width: 90pt; min-width: 90pt;
        padding: 3pt 0 3pt 18pt;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
      }
      .logo-box .logo-name { font-size: 12pt; font-weight: bold; letter-spacing: 1pt; }
      .logo-box .logo-sub  { font-size: 7pt; color: #6c757d; }
      .header-meta {
        flex: 1; padding: 0 8pt;
        display: flex; flex-direction: column; justify-content: center; gap: 2pt;
      }
      .header-meta .dt  { font-size: 8pt; color: #6c757d; }
      .header-meta .dev { font-size: 9pt; font-weight: bold; }
      .header-meta .ln  { font-size: 8pt; }
      .header-title {
        flex: 1; text-align: center;
        display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 6pt;
      }
      .header-title h1 { font-size: 11pt; font-weight: bold; color: #1a3a6b; }
      .rfid-badge {
        display: inline-block; font-size: 9pt; font-weight: bold;
        padding: 3pt 10pt; border: 1px solid;
      }
      .badge-pass { background:#d4edda; color:#155724; border-color:#155724; }
      .badge-fail { background:#f8d7da; color:#721c24; border-color:#721c24; }
      .badge-warn { background:#fff3cd; color:#856404; border-color:#856404; }
      .company-box {
        width: 90pt; min-width: 90pt;
        padding: 3pt 18pt 3pt 0; font-size: 7pt; color: #6c757d;
        display: flex; align-items: center; justify-content: center;
      }

      /* ── CONTENT ─────────────────────────────────────────── */
      .content { flex: 1; margin-top: 6pt; display: flex; flex-direction: column; gap: 6pt; }

      .sec-hdr {
        background: #1a3a6b; color: white;
        font-size: 10pt; font-weight: bold; padding: 3pt 6pt;
      }
      .sec-sub-hdr {
        background: #2c5296; color: white;
        font-size: 8pt; font-weight: bold; padding: 3pt 6pt;
      }
      /* Inline annotation appended to a section sub-header — inherits white colour */
      .sec-note { font-size: 7.5pt; font-weight: normal; }
      .sec-sub-hdr-gap { margin-top: 6pt; }

      /* ── BARCODE SUMMARY TABLE ───────────────────────────── */
      .sum-table {
        width: 100%; border-collapse: collapse; table-layout: fixed;
        border: 1px solid #1a3a6b; font-size: 9pt;
      }
      .sum-table th {
        background: white; border-bottom: 1px solid #999; border-right: 1px solid #999;
        padding: 3pt 4pt; text-align: left; font-size: 8pt; font-weight: bold;
      }
      .sum-table th:last-child { border-right: none; }
      .sum-table td {
        border-bottom: 1px solid #aaa;
        padding: 2.5pt 4pt 2pt;
        white-space: nowrap; overflow: hidden;
      }
      .sum-table td:not(:last-child) { border-right: 1px solid #aaa; }
      .sum-table tr:last-child td { border-bottom: none; }
      /* Navy separator above Report Name row — 1px matches the exterior border weight;
         navy colour beats #aaa in border-collapse with equal width + style. */
      .sum-table tr.sum-meta-start td { border-top: 1px solid #1a3a6b; }
      .sum-table .meta-lbl { color: #555; font-size: 8.5pt; }

      /* App-spec cell: flex row — "GS1 —" left, stacked two-liner right */
      .sum-table td.app-spec {
        white-space: normal;
        overflow: visible;
        padding: 0 4pt;    /* horizontal only; inner div handles vertical */
      }
      .app-spec-inner {
        display: flex;
        align-items: center;
        gap: 3pt;
        padding: 2.5pt 0 2pt;
      }
      .app-spec-prefix { font-size: 8pt; white-space: nowrap; }
      .app-spec-stack  { font-size: 7pt; line-height: 1.2; white-space: nowrap; }

      /* ── BARCODE GRADES TABLE ────────────────────────────── */
      .grades-table {
        width: 100%; border-collapse: collapse;
        border: 1px solid #1a3a6b; font-size: 8pt;
      }
      .grades-table th {
        border-bottom: 1px solid #999; border-right: 1px solid #999;
        padding: 3pt 4pt; text-align: center; font-weight: bold; background: white;
      }
      .grades-table th:last-child { border-right: none; }
      .grades-table td {
        border-right: 1px solid #aaa;
        padding: 2.5pt 4pt 2pt; text-align: center;
      }
      .grades-table td:last-child { border-right: none; }
      .grades-table tr:not(:last-child) td { border-bottom: 1px solid #aaa; }

      /* ── RFID TABLE ──────────────────────────────────────── */
      .rfid-table {
        width: 100%; border-collapse: collapse;
        border: 1px solid #1a3a6b; font-size: 9pt;
      }
      .rfid-table th {
        background: #e8eef5; border-bottom: 1px solid #aaa;
        padding: 3pt 4pt; text-align: left; font-size: 8pt;
      }
      .rfid-table td {
        border-bottom: 1px solid #aaa;
        padding: 2.5pt 4pt 2pt;
      }
      .rfid-table .lbl-col { width: 178pt; }
      .row-hi          td { background: #f0f7ff; }
      .row-result-pass td { background:#d4edda; color:#155724; font-weight:bold; }
      .row-result-fail td { background:#f8d7da; color:#721c24; font-weight:bold; }
      .row-result-warn td { background:#fff3cd; color:#856404; font-weight:bold; }
      .rfid-table tr:last-child td { border-bottom: none; }

      .gcp-inline-note { font-size: 7.5pt; font-style: italic; }

      /* ── BARCODE IMAGE ───────────────────────────────────── */
      .img-frame { border: 1px solid #1a3a6b; text-align: center; padding: 8pt; }
      .img-placeholder {
        background: #f5f5f5; border: 1px dashed #aaa;
        display: inline-block; padding: 18pt 32pt; color: #888; font-size: 8pt;
      }

      /* ── DATA FORMAT CHECK ───────────────────────────────── */
      .dfc-table {
        width: 100%; border-collapse: collapse;
        border: 1px solid #1a3a6b; font-size: 8pt;
      }
      .dfc-table th {
        background: #e8eef5; border-bottom: 1px solid #aaa;
        padding: 3pt 4pt; text-align: left;
      }
      .dfc-table td { border-bottom: 1px solid #aaa; padding: 2.5pt 4pt 2pt; }
      .dfc-table .chk { width: 50pt; font-weight: bold; }
      .dfc-table tr:last-child td { border-bottom: none; }
      .pass-fg { color: #155724; }
      .fail-fg { color: #721c24; }

      .overall-pill {
        display: inline-block; font-size: 8pt; font-weight: bold;
        padding: 3pt 8pt; border: 1px solid; margin-top: 4pt; float: right;
      }
      .pill-pass { background:#d4edda; color:#155724; border-color:#155724; }
      .pill-fail { background:#f8d7da; color:#721c24; border-color:#721c24; }
      .cf::after { content:""; display:table; clear:both; }

      /* ── FOOTER ──────────────────────────────────────────── */
      .footer {
        border-top: 1px solid #ccc; padding-top: 4pt; margin-top: 8pt;
        display: flex; justify-content: space-between;
        font-size: 7pt; color: #6c757d;
      }
      .footer .fl { font-weight: bold; color: #222; }
      .footer .fl span { font-weight: normal; color: #6c757d; }

      @media print {
        body { background: white; padding: 0; }
        .page { box-shadow: none; margin: 0; width: 100%; }
        .print-hint { display: none; }
        * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
      }
      """;
}
