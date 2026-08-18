// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  VCCS RFID Validation Report — HTML generator                            ║
// ║                                                                          ║
// ║  Loads vccs-report-template.html (EmbeddedResource) and replaces         ║
// ║  {{TOKEN}} and {{SLOT_*}} markers with live VerificationRecord data.     ║
// ║                                                                          ║
// ║  ALL visual/layout changes belong in the template file.                  ║
// ║  C# only provides data values and generates variable-length HTML blocks  ║
// ║  for the five dynamic table sections.                                    ║
// ║                                                                          ║
// ║  The HTML is rendered to PDF by VccsPdfRenderer (WebView2 primary,       ║
// ║  wkhtmltopdf silent fallback).                                           ║
// ╚══════════════════════════════════════════════════════════════════════════╝

using System.Text;
using ExcelEngine.Models;

namespace DeviceInterface.Reports;

/// <summary>
/// Generates the v23-design VCCS FlexWedge™ Pro RFID Validation Report as a
/// self-contained HTML string by loading the embedded template and substituting
/// live data via token replacement.
/// </summary>
public static class VccsHtmlReportGenerator
{
    /// <summary>Report format version — bump on ANY layout/content/logic change.</summary>
    public const string ReportVersion = "v1.5.5";

    // ── Template ────────────────────────────────────────────────────────────

    // Loaded once per process from the EmbeddedResource.
    // Resource name: DeviceInterface.Reports.vccs-report-template.html
    private static readonly Lazy<string> _template = new(() =>
    {
        var asm = typeof(VccsHtmlReportGenerator).Assembly;
        const string Name = "DeviceInterface.Reports.vccs-report-template.html";
        using var stream = asm.GetManifestResourceStream(Name)
            ?? throw new InvalidOperationException(
                $"Embedded resource '{Name}' not found. " +
                "Ensure vccs-report-template.html is included as EmbeddedResource in DeviceInterface.csproj.");
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    });

    // ── Public API ─────────────────────────────────────────────────────────

    public static string Generate(VerificationRecord r)
    {
        // ── logos ──────────────────────────────────────────────────────────
        string? vccsB64    = LoadLogoBase64("vccs_logo.png");
        string? companyB64 = LoadLogoBase64FromPath(r.LogoPath) ?? LoadLogoBase64("pips_logo.png");

        string vccsLogoHtml = vccsB64 is not null
            ? $"<img src=\"data:image/png;base64,{vccsB64}\" style=\"max-height:65pt;max-width:68pt;object-fit:contain;\" alt=\"VCCS\" />"
            : "<div class=\"logo-name\">VCCS</div><div class=\"logo-sub\">FlexWedge&#x2122; Pro</div>";

        string companyLogoHtml = companyB64 is not null
            ? $"<img src=\"data:image/png;base64,{companyB64}\" style=\"max-height:48pt;max-width:68pt;object-fit:contain;\" alt=\"{H(r.CompanyName ?? "Company")}\" />"
            : H(r.CompanyName ?? "Company Logo");

        // ── badge ─────────────────────────────────────────────────────────
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

        // ── section ① title ───────────────────────────────────────────────
        string? brand = r.VerifierBrand ?? (r.DeviceModel is { Length: > 0 } ? "COGNEX" : null);
        string? dataManInfix = (brand == "COGNEX" &&
            r.DeviceModel?.Contains("DataMan", StringComparison.OrdinalIgnoreCase) == true)
            ? "DataMan " : null;
        string section1Title = brand is not null
            ? $"{H(brand)} {dataManInfix}TruCheck Barcode Verification Results Summary"
            : "TruCheck Barcode Verification Results Summary";

        // ── report name ───────────────────────────────────────────────────
        string reportName = !string.IsNullOrWhiteSpace(r.WebscanSourcePath)
            ? H(Path.GetFileName(r.WebscanSourcePath))
            : H($"{r.VerificationDateTime:yyyy-MM-dd_HH-mm-ss}_vccs_rfid.pdf");

        // ── RFID section adjective (EPC vs UHF) ───────────────────────────
        bool isGS1 = r.ApplicationStandard?.StartsWith("GS1", StringComparison.OrdinalIgnoreCase) != false;
        string rfidAdj = isGS1 ? "EPC" : "UHF";

        // ── token replacement ─────────────────────────────────────────────
        return _template.Value
            .Replace("{{HDR_VCCS_LOGO}}",      vccsLogoHtml)
            .Replace("{{HDR_COMPANY_LOGO}}",    companyLogoHtml)
            .Replace("{{HDR_DATETIME}}",        H(r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt")))
            .Replace("{{HDR_DEVICE}}",          H(r.DeviceName ?? r.DeviceModel ?? r.VerifierBrand ?? "\u2014"))
            .Replace("{{HDR_SERIAL}}",          H(r.DeviceSerial ?? "\u2014"))
            .Replace("{{HDR_SW}}",              H(r.SoftwareVersion ?? "\u2014"))
            .Replace("{{HDR_FW}}",              H(r.FirmwareVersion ?? "\u2014"))
            .Replace("{{HDR_BADGE_CLASS}}",     badgeCls)
            .Replace("{{HDR_BADGE_TEXT}}",      badgeTxt)
            .Replace("{{SECTION1_TITLE}}",      section1Title)
            .Replace("{{SLOT_SYMBOL_ROWS}}",    BuildSymbolRows(r))
            .Replace("{{REPORT_NAME}}",         reportName)
            .Replace("{{REPORT_DATETIME}}",     H(r.VerificationDateTime.ToString("ddd dd-MMM-yyyy hh:mm:ss tt")))
            .Replace("{{SLOT_GRADE_ROWS}}",     BuildGradeRows(r))
            .Replace("{{RFID_ADJ}}",            rfidAdj)
            .Replace("{{SLOT_RFID_ROWS}}",      BuildRfidRows(r))
            .Replace("{{SLOT_IMAGE}}",          BuildImageSlot(r))
            .Replace("{{SLOT_DFC_SECTION}}",    BuildDfcSection(r))
            .Replace("{{FOOTER_VERSION}}",      ReportVersion)
            .Replace("{{FOOTER_GENERATED}}",    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
    }

    // ── Slot builders ───────────────────────────────────────────────────────

    private static string BuildSymbolRows(VerificationRecord r)
    {
        bool multiMode = !string.IsNullOrWhiteSpace(r.LinearSymbology);
        var sb = new StringBuilder();
        if (multiMode)
            AppendSymbolRow(sb, r.LinearSymbology!, r.LinearDecodedData, LinearAppSpecCell(r.LinearSymbology));
        AppendSymbolRow(sb, r.Symbology ?? "\u2014", r.DecodedData, AppSpecCell(r));
        return sb.ToString();
    }

    private static void AppendSymbolRow(StringBuilder sb, string symb, string? encoded, string appSpecCell)
    {
        sb.Append($"          <tr>\n");
        sb.Append($"            <td style=\"font-size:8pt;\">{H(symb)}</td>\n");
        sb.Append($"            <td style=\"font-family:Consolas,monospace;\">{H(encoded ?? "\u2014")}</td>\n");
        sb.Append($"            {appSpecCell}\n");
        sb.Append($"          </tr>\n");
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

    private static string AppSpecStackCell(string line1, string line2) =>
        $"<td class=\"app-spec\"><div class=\"app-spec-inner\"><span class=\"app-spec-prefix\">GS1 &#x2014;</span><div class=\"app-spec-stack\">{H(line1)}<br>{H(line2)}</div></div></td>";

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

    private static string BuildGradeRows(VerificationRecord r)
    {
        bool multiMode = !string.IsNullOrWhiteSpace(r.LinearSymbology);
        var sb = new StringBuilder();

        if (multiMode)
            AppendGradeRow(sb, r.LinearSymbology!, r.LinearStandard ?? "ISO/IEC 15416",
                GradeDisplay(r.LinearOverallGrade),
                r.LinearAperture.HasValue ? r.LinearAperture.Value.ToString("D2") : "\u2014",
                r.LinearWavelength?.ToString() ?? "\u2014",
                r.LinearLighting ?? "\u2014", r.LinearFormalGrade ?? "\u2014");

        AppendGradeRow(sb, r.Symbology ?? "\u2014", r.Standard ?? "\u2014",
            GradeDisplay(r.OverallGrade),
            r.Aperture.HasValue ? r.Aperture.Value.ToString("D2") : "\u2014",
            r.Wavelength?.ToString() ?? "\u2014",
            r.Lighting ?? "\u2014", r.FormalGrade ?? "\u2014");

        return sb.ToString();
    }

    private static void AppendGradeRow(StringBuilder sb, string symb, string standard,
        string grade, string aperture, string wavelength, string lighting, string formal)
    {
        sb.Append($"          <tr>\n");
        sb.Append($"            <td>{H(symb)}</td>\n");
        sb.Append($"            <td>{H(standard)}</td>\n");
        sb.Append($"            <td>{H(grade)}</td>\n");
        sb.Append($"            <td>{H(aperture)}</td>\n");
        sb.Append($"            <td>{H(wavelength)}</td>\n");
        sb.Append($"            <td>{H(lighting)}</td>\n");
        sb.Append($"            <td>{H(formal)}</td>\n");
        sb.Append($"          </tr>\n");
    }

    private static string BuildRfidRows(VerificationRecord r)
    {
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

        var sb = new StringBuilder();

        // Fixed data rows — Row 1: tag detected + lock status
        sb.Append($"          <tr>\n");
        sb.Append($"            <td>Tag Detected / Lock Status</td>\n");
        sb.Append($"            <td>{H(tagLabel)} &#x2014; {H(lockDisplay)} <span style=\"color:#555;font-size:8pt;\">(Permalocked / Locked / Unlocked / Unknown)</span></td>\n");
        sb.Append($"          </tr>\n");

        // Row 2: TID — always shown immediately after Tag Detected; blank when no tag
        sb.Append($"          <tr class=\"row-hi\">\n");
        sb.Append($"            <td>TID (Chip Identity)</td>\n");
        sb.Append($"            <td style=\"font-family:Consolas,monospace;\">{H(r.RfidTid ?? (tagDetected ? "\u2014" : ""))}</td>\n");
        sb.Append($"          </tr>\n");

        // Row 3+: EPC decoded fields
        sb.Append($"          <tr class=\"row-hi\">\n");
        sb.Append($"            <td>EPC Encoding Scheme / GCP Length</td>\n");
        sb.Append($"            <td><span style=\"font-family:Consolas,monospace;\">{H(schemePart)}</span> &#x2014; {H(gcpDisplay)}{gcpNote}</td>\n");
        sb.Append($"          </tr>\n");

        sb.Append($"          <tr class=\"row-hi\">\n");
        sb.Append($"            <td>EPC Hex</td>\n");
        sb.Append($"            <td style=\"font-family:Consolas,monospace;\">{H(r.RfidEpcHex ?? "\u2014")}</td>\n");
        sb.Append($"          </tr>\n");

        sb.Append($"          <tr class=\"row-hi\">\n");
        sb.Append($"            <td>EPC Tag URI</td>\n");
        sb.Append($"            <td style=\"font-family:Consolas,monospace;font-size:8pt;\">{H(r.RfidEpcTagUri ?? "\u2014")}</td>\n");
        sb.Append($"          </tr>\n");

        sb.Append($"          <tr class=\"row-hi\">\n");
        sb.Append($"            <td>GTIN-14</td>\n");
        sb.Append($"            <td style=\"font-family:Consolas,monospace;\">{H(r.RfidGtin14 ?? "\u2014")}</td>\n");
        sb.Append($"          </tr>\n");

        sb.Append($"          <tr class=\"row-hi\">\n");
        sb.Append($"            <td>Serial Number</td>\n");
        sb.Append($"            <td style=\"font-family:Consolas,monospace;\">{H(r.RfidSerial ?? "\u2014")}</td>\n");
        sb.Append($"          </tr>\n");

        // Result row(s) — coloured by pass/fail/warn
        string rowCls = r.RfidStatus switch
        {
            "Pass" => "row-result-pass",
            "Fail" => "row-result-fail",
            _      => "row-result-warn",
        };

        bool multiMode = !string.IsNullOrWhiteSpace(r.LinearSymbology);
        bool is1DOnly  = r.Is1D && !multiMode;

        void ResultRow(string label, string value)
        {
            sb.Append($"          <tr class=\"{rowCls}\">\n");
            sb.Append($"            <td>{H(label)}</td>\n");
            sb.Append($"            <td>{value}</td>\n");
            sb.Append($"          </tr>\n");
        }

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
                "Fail" => BuildMismatch2DLabel(r.RfidMismatchDetail),
                var s  => H(s ?? "\u2014"),
            };
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
                "Fail"               => BuildMismatch2DLabel(r.RfidMismatchDetail),
                var s                => H(s ?? "\u2014"),
            };
            ResultRow($"{symName} Validation Result", singleVal);
        }

        return sb.ToString();
    }

    private static string BuildImageSlot(VerificationRecord r)
    {
        string? img2D     = r.RoiJpegImageBase64 ?? r.JpegImageBase64;
        string? imgLinear = r.LinearJpegImageBase64;
        bool multiMode    = !string.IsNullOrWhiteSpace(r.LinearSymbology);

        var sb = new StringBuilder();
        sb.Append("    <div>\n");

        if (multiMode && !string.IsNullOrWhiteSpace(img2D) && !string.IsNullOrWhiteSpace(imgLinear))
        {
            string img2DForDual = r.JpegImageBase64 ?? img2D;
            sb.Append("      <div class=\"sec-sub-hdr\">Barcode Images</div>\n");
            sb.Append("      <div class=\"img-frame\">\n");
            sb.Append("        <table style=\"width:100%;border-collapse:collapse;\"><tr>\n");
            sb.Append($"          <td style=\"width:50%;text-align:center;vertical-align:middle;\">\n");
            sb.Append($"            <div style=\"font-size:7pt;color:#6c757d;font-style:italic;\">{H(r.LinearSymbology!)}</div>\n");
            sb.Append($"            <img src=\"data:image/jpeg;base64,{imgLinear}\" alt=\"{H(r.LinearSymbology!)}\" style=\"max-width:3.4in;max-height:1.2in;object-fit:contain;\"/>\n");
            sb.Append($"          </td>\n");
            sb.Append($"          <td style=\"width:50%;text-align:center;vertical-align:middle;\">\n");
            sb.Append($"            <div style=\"font-size:7pt;color:#6c757d;font-style:italic;\">{H(r.Symbology ?? "2D")}</div>\n");
            sb.Append($"            <img src=\"data:image/jpeg;base64,{img2DForDual}\" alt=\"{H(r.Symbology ?? "2D")}\" style=\"max-width:3.4in;max-height:1.2in;object-fit:contain;\"/>\n");
            sb.Append($"          </td>\n");
            sb.Append("        </tr></table>\n");
            sb.Append("      </div>\n");
        }
        else if (!string.IsNullOrWhiteSpace(img2D))
        {
            sb.Append("      <div class=\"sec-sub-hdr\">Barcode Image</div>\n");
            sb.Append("      <div class=\"img-frame\">\n");
            sb.Append($"        <img src=\"data:image/jpeg;base64,{img2D}\" alt=\"Barcode\" style=\"max-width:100%;max-height:1.35in;object-fit:contain;\"/>\n");
            sb.Append("      </div>\n");
        }
        else
        {
            sb.Append("      <div class=\"sec-sub-hdr\">Barcode Image</div>\n");
            sb.Append("      <div class=\"img-frame\">\n");
            sb.Append("        <div class=\"img-placeholder\">[No barcode image available for this scan]</div>\n");
            sb.Append("      </div>\n");
        }

        sb.Append("    </div>\n");
        return sb.ToString();
    }

    private static string BuildDfcSection(VerificationRecord r)
    {
        bool has2D     = r.DataFormatCheck       is { Rows.Count: > 0 };
        bool hasLinear = r.LinearDataFormatCheck is { Rows.Count: > 0 };
        if (!has2D && !hasLinear) return string.Empty;

        bool anyFail =
            (has2D     && r.DataFormatCheck!.Overall       == OverallPassFail.Fail) ||
            (hasLinear && r.LinearDataFormatCheck!.Overall == OverallPassFail.Fail);
        string pillCls = anyFail ? "pill-fail" : "pill-pass";
        string pillTxt = anyFail ? "OVERALL: FAIL" : "OVERALL: PASS";

        string std = r.DataFormatCheck?.Standard
                  ?? r.LinearDataFormatCheck?.Standard ?? "GS1";
        string hdrSuffix = std.Contains("GS1", StringComparison.OrdinalIgnoreCase) ? "GS1" : std;

        var sb = new StringBuilder();
        sb.Append($"    <div class=\"cf\">\n");
        sb.Append($"      <div class=\"sec-sub-hdr\">Data Format Check &#x2014; {H(hdrSuffix)}<span class=\"sec-note\"> &#x2014; <em>See associated TruCheck report for additional details</em></span></div>\n");
        sb.Append($"      <table class=\"dfc-table\">\n");
        sb.Append($"        <thead><tr><th>Field</th><th>Data</th><th class=\"chk\">Check</th></tr></thead>\n");
        sb.Append($"        <tbody>\n");

        bool bothPresent = has2D && hasLinear;

        void DfcRows(DataFormatCheckResult dfc, string? symbLabel)
        {
            if (symbLabel is not null)
                sb.Append($"          <tr><td colspan=\"3\" style=\"background:#e8eef5;font-weight:bold;\">{H(symbLabel)}</td></tr>\n");
            foreach (var row in dfc.Rows)
            {
                bool fail = string.Equals(row.Check, "FAIL", StringComparison.OrdinalIgnoreCase);
                string cls = fail ? "chk fail-fg" : "chk pass-fg";
                bool mono = row.Name.Contains("GTIN", StringComparison.OrdinalIgnoreCase);
                string dataStyle = mono ? " style=\"font-family:Consolas,monospace;\"" : "";
                sb.Append($"          <tr><td>{H(row.Name)}</td><td{dataStyle}>{H(row.Data)}</td><td class=\"{cls}\">{H(row.Check)}</td></tr>\n");
            }
        }

        if (hasLinear) DfcRows(r.LinearDataFormatCheck!, bothPresent ? r.LinearSymbology : null);
        if (has2D)     DfcRows(r.DataFormatCheck!,       bothPresent ? r.Symbology      : null);

        sb.Append($"        </tbody>\n");
        sb.Append($"      </table>\n");
        sb.Append($"      <span class=\"overall-pill {pillCls}\">{pillTxt}</span>\n");
        sb.Append($"    </div>\n");
        return sb.ToString();
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

    /// <summary>
    /// Produces a specific "Fail — …" label for a 2D RFID cross-validation failure
    /// by inspecting <paramref name="mismatchDetail"/> to name which field(s) failed.
    ///
    /// Detail format (from RfidValidator): semicolon-separated tokens, each beginning
    /// with the field name followed by a colon, e.g.:
    ///   "GTIN14:RFID=00012345678905,BC=00012345678906;Serial:RFID=12345,BC=12346"
    ///
    /// Returned strings use HTML entity &#x2014; for the em dash and are already
    /// safe to embed in HTML (no user-supplied data is interpolated).
    /// </summary>
    private static string BuildMismatch2DLabel(string? mismatchDetail)
    {
        if (string.IsNullOrWhiteSpace(mismatchDetail))
            return "Fail &#x2014; GTIN or Serial Number mismatch";

        bool gtinNoData     = mismatchDetail.Contains("GTIN14:NoBarcodeData",   StringComparison.Ordinal);
        bool gtinMismatch   = mismatchDetail.Contains("GTIN14:RFID=",           StringComparison.Ordinal);
        bool serialMissing  = mismatchDetail.Contains("Serial:MissingFromTag",  StringComparison.Ordinal);
        bool serialMismatch = mismatchDetail.Contains("Serial:RFID=",           StringComparison.Ordinal);

        var parts = new List<string>(4);
        if (gtinNoData)     parts.Add("GTIN not in barcode");
        if (gtinMismatch)   parts.Add("GTIN mismatch");
        if (serialMissing)  parts.Add("Serial Number missing from tag");
        if (serialMismatch) parts.Add("Serial Number mismatch");

        return parts.Count switch
        {
            0 => "Fail &#x2014; mismatch",
            1 => $"Fail &#x2014; {parts[0]} mismatch",
            2 => $"Fail &#x2014; {parts[0]} and {parts[1]} mismatch",
            _ => $"Fail &#x2014; {string.Join(", ", parts[..^1])} and {parts[^1]} mismatch",
        };
    }

    private static string H(string? s) =>
        s is null ? string.Empty
        : s.Replace("&", "&amp;")
           .Replace("<", "&lt;")
           .Replace(">", "&gt;")
           .Replace("\"", "&quot;")
           .Replace("'", "&#39;");
}
