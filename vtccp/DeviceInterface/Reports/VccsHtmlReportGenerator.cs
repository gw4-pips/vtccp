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
using System.Globalization;
using System.Text.RegularExpressions;
using ExcelEngine.Models;

namespace DeviceInterface.Reports;

/// <summary>
/// Generates the v23-design RFID VeriWedge™ PowerPro Validation Report as a
/// self-contained HTML string by loading the embedded template and substituting
/// live data via token replacement.
/// </summary>
public static class VccsHtmlReportGenerator
{
    /// <summary>Report format version — bump on ANY layout/content/logic change.</summary>
    public const string ReportVersion = "v1.5.23";
    internal const int MaxRenderedSymbolGroups = 2;

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
        bool hasCorrelatedHtml = HasCorrelatedFilesystemHtml(r);

        // ── logos ──────────────────────────────────────────────────────────
        string? vccsB64    = LoadLogoBase64("vccs_logo.png");
        string? companyB64 = LoadLogoBase64FromPath(r.LogoPath) ?? LoadLogoBase64("pips_logo.png");

        string vccsLogoHtml = vccsB64 is not null
            ? $"<img src=\"data:image/png;base64,{vccsB64}\" style=\"max-height:65pt;max-width:68pt;object-fit:contain;\" alt=\"VCCS\" />"
            : "<div class=\"logo-name\">VCCS</div><div class=\"logo-sub\">RFID VeriWedge&#x2122; PowerPro</div>";

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
        // Do not infer a verifier brand or report identity from reader metadata.
        string section1Title = hasCorrelatedHtml
            ? "TruCheck Barcode Verification Results Summary"
            : "[BARCODE VERIFICATION UNAVAILABLE — NO CORRELATED DMST HTML]";

        // ── report name ───────────────────────────────────────────────────
        string reportName = hasCorrelatedHtml
            ? H(r.HtmlSourceFileName!)
            : "[NO CORRELATED DMST HTML REPORT]";

        // ── RFID section adjective (EPC vs UHF) ───────────────────────────
        bool isGS1 = r.ApplicationStandard?.StartsWith("GS1", StringComparison.OrdinalIgnoreCase) != false;
        string rfidAdj = isGS1 ? "EPC" : "UHF";

        // ── report date/time ──────────────────────────────────────────────
        // The DMST filename wins only when it contains a timestamp. Otherwise
        // preserve the HTML Verified: string exactly as displayed by the device.
        string reportDateTime = H(GetSourceDateTimeText(r));

        // ── token replacement ─────────────────────────────────────────────
        return _template.Value
            .Replace("{{HDR_VCCS_LOGO}}",      vccsLogoHtml)
            .Replace("{{HDR_COMPANY_LOGO}}",    companyLogoHtml)
            .Replace("{{HDR_DEVICE}}",          H(r.DeviceName ?? "\u2014"))
            .Replace("{{HDR_SERIAL}}",          H(r.DeviceSerial ?? "\u2014"))
            .Replace("{{HDR_FW}}",              H(r.FirmwareVersion ?? "\u2014"))
            .Replace("{{HDR_BADGE_CLASS}}",     badgeCls)
            .Replace("{{HDR_BADGE_TEXT}}",      badgeTxt)
            .Replace("{{SECTION1_TITLE}}",      section1Title)
            .Replace("{{SLOT_SYMBOL_ROWS}}",    BuildSymbolRows(r))
            .Replace("{{REPORT_NAME}}",         reportName)
            .Replace("{{REPORT_DATETIME}}",     reportDateTime)
            .Replace("{{SLOT_GRADE_ROWS}}",     BuildGradeRows(r))
            .Replace("{{RFID_ADJ}}",            rfidAdj)
            .Replace("{{SLOT_RFID_ROWS}}",      BuildRfidRows(r))
            .Replace("{{SLOT_BARCODE_DETAIL}}", BuildBarcodeDetailSection(r))
            .Replace("{{FOOTER_VERSION}}",      ReportVersion)
            .Replace("{{FOOTER_GENERATED}}",    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
    }

    /// <summary>
    /// Returns the literal local <c>Verified:</c> value from the correlated HTML.
    /// </summary>
    internal static string GetSourceDateTimeText(VerificationRecord r)
    {
        if (HasCorrelatedFilesystemHtml(r) &&
            !string.IsNullOrWhiteSpace(r.HtmlVerifiedString))
            return r.HtmlVerifiedString;

        return "[UNAVAILABLE — NO CORRELATED DMST HTML]";
    }

    /// <summary>Timestamp used for the generated PDF filename, with no offset conversion.</summary>
    internal static string GetOutputTimestamp(VerificationRecord r)
    {
        string sourceText = GetSourceDateTimeText(r);
        string parseable = Regex.Replace(sourceText, @"\(\d+ms\)", string.Empty,
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Trim();
        parseable = Regex.Replace(parseable, @"^[A-Za-z]{3}\s+", string.Empty,
            RegexOptions.CultureInvariant);
        return DateTime.TryParseExact(parseable,
                ["dd-MMM-yyyy hh:mm:ss tt"],
                CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime dt)
            ? dt.ToString("yyyy-MM-dd_HH-mm-ss")
            : "unknown";
    }

    // ── Slot builders ───────────────────────────────────────────────────────

    private static string BuildSymbolRows(VerificationRecord r)
    {
        if (!HasCorrelatedFilesystemHtml(r))
            return UnavailableSymbolRow(r);

        bool multiMode = !string.IsNullOrWhiteSpace(r.LinearSymbology);
        int renderedGroups = 0;
        var sb = new StringBuilder();
        if (multiMode && renderedGroups < MaxRenderedSymbolGroups)
        {
            AppendSymbolRow(sb, r.LinearSymbology!, r.LinearDecodedData,
                ApplicationSettingsCell(r, hasCorrelatedHtml: true));
            renderedGroups++;
        }
        if (renderedGroups < MaxRenderedSymbolGroups)
        {
            AppendSymbolRow(sb, r.HtmlSymbology ?? "\u2014", r.HtmlDecodedData,
                ApplicationSettingsCell(r, hasCorrelatedHtml: true));
        }
        return sb.ToString();
    }

    private static string UnavailableSymbolRow(VerificationRecord r)
        => "          <tr>\n" +
           "            <td>[UNAVAILABLE]</td>\n" +
           "            <td>[UNAVAILABLE — NOT PRESENT IN CORRELATED TRUCHECK HTML]</td>\n" +
           $"            {ApplicationSettingsCell(r, hasCorrelatedHtml: false)}\n" +
           "          </tr>\n";

    private static void AppendSymbolRow(
        StringBuilder sb, string symb, string? encoded, string applicationSettingsCell)
    {
        sb.Append($"          <tr data-vccs-symbol-group=\"true\">\n");
        sb.Append($"            <td style=\"font-size:8pt;\">{H(symb)}</td>\n");
        sb.Append($"            <td style=\"font-family:Consolas,monospace;\">{H(encoded ?? "\u2014")}</td>\n");
        sb.Append($"            {applicationSettingsCell}\n");
        sb.Append($"          </tr>\n");
    }

    private static string ApplicationSettingsCell(VerificationRecord r, bool hasCorrelatedHtml)
    {
        // App Standard and DFC are sourced only from the correlated TruCheck HTML.
        // The aperture mode is the TC setting captured via GET TRUCHECK.APERTURE.
        // Do not substitute local validation or a numeric grade aperture here.
        string applicationStandard = hasCorrelatedHtml
            ? r.HtmlApplicationStandard ?? "\u2014"
            : "\u2014";
        string dataFormatCheck = hasCorrelatedHtml
            ? DisplayDataFormatCheckSetting(r.HtmlDataFormatCheck)
            : "\u2014";
        string apertureSetting = r.ApertureSettingMode ?? "\u2014";

        return $"<td class=\"app-settings\">{H(applicationStandard)}" +
               $"<span class=\"app-settings-separator\"> / </span>{H(dataFormatCheck)}" +
               $"<span class=\"app-settings-separator\"> / </span>{H(apertureSetting)}</td>";
    }

    private static string DisplayDataFormatCheckSetting(DataFormatCheckResult? dfc)
    {
        string? standard = dfc?.Standard?.Trim();
        if (string.IsNullOrWhiteSpace(standard))
            return "None";

        // These are display aliases for the exact TruCheck HTML standard labels,
        // matching the values exposed in the TruCheck Application Settings UI.
        if (standard.Contains("GS1", StringComparison.OrdinalIgnoreCase))
            return "GS1";
        if (standard.Contains("HIBCC", StringComparison.OrdinalIgnoreCase))
            return "HIBCC";
        if (standard.Contains("15434", StringComparison.OrdinalIgnoreCase))
            return "ISO 15434";

        return standard;
    }

    private static string BuildGradeRows(VerificationRecord r)
    {
        if (!HasCorrelatedFilesystemHtml(r))
            return UnavailableGradeRow();

        bool multiMode = !string.IsNullOrWhiteSpace(r.LinearSymbology);
        int renderedGroups = 0;
        var sb = new StringBuilder();

        if (multiMode && renderedGroups < MaxRenderedSymbolGroups)
        {
            AppendGradeRow(sb, r.LinearSymbology!,
                r.HtmlLinearStandard,
                r.HtmlLinearGradeDisplay,
                r.HtmlLinearAperture,
                r.HtmlLinearWavelength,
                r.HtmlLinearLighting,
                r.HtmlLinearFormalGrade);
            renderedGroups++;
        }

        if (renderedGroups < MaxRenderedSymbolGroups)
        {
            AppendGradeRow(sb, r.HtmlSymbology,
                r.HtmlStandard,
                r.HtmlOverallGradeDisplay,
                r.HtmlAperture,
                r.HtmlWavelength,
                r.HtmlLighting,
                r.HtmlFormalGrade);
        }

        return sb.ToString();
    }

    private static string UnavailableGradeRow()
        => "          <tr>\n" +
           "            <td>[UNAVAILABLE]</td>\n" +
           "            <td colspan=\"6\">[UNAVAILABLE — NO CORRELATED DMST HTML]</td>\n" +
           "          </tr>\n";

    private static void AppendGradeRow(StringBuilder sb, string? symb, string? standard,
        string? grade, string? aperture, string? wavelength,
        string? lighting, string? formal)
    {
        static string Display(string? value)
            => H(value ?? "[UNAVAILABLE — NOT PRESENT IN CORRELATED TRUCHECK HTML]");

        sb.Append($"          <tr>\n");
        sb.Append($"            <td>{Display(symb)}</td>\n");
        sb.Append($"            <td>{Display(standard)}</td>\n");
        sb.Append($"            <td>{Display(grade)}</td>\n");
        sb.Append($"            <td>{Display(aperture)}</td>\n");
        sb.Append($"            <td>{Display(wavelength)}</td>\n");
        sb.Append($"            <td>{Display(lighting)}</td>\n");
        sb.Append($"            <td>{Display(formal)}</td>\n");
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
            "PermaLocked" => "Permalocked",
            "Unlocked"    => "Unlocked",
            "Unknown"     => "Unknown",
            null          => "Unknown",
            _             => "Unknown",
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
            ? $"<em class=\"gcp-inline-note\"> &ndash; from official GS1 GCP prefix table as of {H(r.RfidGcpTableDate)}</em>"
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

    private static string BuildBarcodeDetailSection(VerificationRecord r)
    {
        bool hasHtml       = HasCorrelatedFilesystemHtml(r);
        string? img2D      = hasHtml ? r.HtmlBarcodeImageBase64 : null;
        bool multiMode    = !string.IsNullOrWhiteSpace(r.LinearSymbology);
        DataFormatCheckResult? htmlDfc = hasHtml ? r.HtmlDataFormatCheck : null;

        var sb = new StringBuilder();
        sb.Append("    <div class=\"barcode-detail-section\">\n");
        sb.Append(hasHtml
            ? "      <div class=\"sec-sub-hdr trucheck-barcode-hdr barcode-detail-header\"><span class=\"trucheck-header-title\">TruCheck Barcode Image <span class=\"detail-separator\">|</span> Data Format Check &#x2014; GS1</span><span class=\"sec-note\"> &#x2014; <em>Native TruCheck data and VCCS Digital Link validation are separately labelled</em></span></div>\n"
            : "      <div class=\"sec-sub-hdr trucheck-barcode-hdr barcode-detail-header\"><span class=\"trucheck-header-title\">Barcode Verification Capture Unavailable</span><span class=\"sec-note\"> &#x2014; <em>No correlated DMST HTML report</em></span></div>\n");
        sb.Append("      <table class=\"barcode-detail-grid\"><tbody><tr>\n");
        sb.Append("        <td class=\"barcode-image-column\">\n");

        if (hasHtml && !string.IsNullOrWhiteSpace(img2D))
        {
            sb.Append($"          <img class=\"barcode-image\" src=\"data:image/jpeg;base64,{img2D}\" alt=\"TruCheck Barcode Image\"/>\n");
        }
        else
        {
            sb.Append(hasHtml
                ? "        <div class=\"img-placeholder\">[BARCODE IMAGE NOT EMBEDDED IN CORRELATED TRUCHECK HTML]</div>\n"
                : "        <div class=\"img-placeholder\">[BARCODE IMAGE UNAVAILABLE — NO CORRELATED DMST HTML]</div>\n");
        }

        sb.Append("        </td>\n");
        sb.Append("        <td class=\"barcode-dfc-column\">\n");

        AppendVendorDataFormatCheck(sb, htmlDfc, hasHtml);
        AppendVccsDigitalLinkValidation(sb, r.VccsDigitalLinkValidation);

        sb.Append("        </td>\n");
        sb.Append("      </tr></tbody></table>\n");
        sb.Append("    </div>\n");
        return sb.ToString();
    }

    private static void AppendVendorDataFormatCheck(
        StringBuilder sb,
        DataFormatCheckResult? htmlDfc,
        bool hasCorrelatedHtml)
    {
        sb.Append("          <div class=\"sec-note\" style=\"margin:0 0 3pt 0;\"><strong>Native TruCheck Data Format Check</strong></div>\n");
        if (htmlDfc is not { Rows.Count: > 0 })
        {
            string unavailableReason = hasCorrelatedHtml
                ? "[DATA FORMAT CHECK UNAVAILABLE — NOT PRESENT IN TRUCHECK HTML]"
                : "[DATA FORMAT CHECK UNAVAILABLE — NO DMST HTML REPORT CORRELATED]";
            sb.Append("          <table class=\"dfc-table\"><thead><tr><th>Field</th><th>Data</th><th class=\"chk\">Check</th></tr></thead><tbody>\n");
            sb.Append($"            <tr><td>Source status</td><td>{H(unavailableReason)}</td><td class=\"chk\">UNAVAILABLE</td></tr>\n");
            sb.Append("          </tbody></table>\n");
            return;
        }

        (string pillCls, string pillTxt) = htmlDfc.Overall switch
        {
            OverallPassFail.Fail => ("pill-fail", "OVERALL: FAIL"),
            OverallPassFail.Pass => ("pill-pass", "OVERALL: PASS"),
            _                    => ("pill-warn", "OVERALL: UNAVAILABLE"),
        };
        sb.Append("          <table class=\"dfc-table\">\n");
        sb.Append("            <thead><tr><th>Field</th><th>Data</th><th class=\"chk\">Check</th></tr></thead>\n");
        sb.Append("            <tbody>\n");
        foreach (var row in htmlDfc.Rows)
        {
            bool fail = string.Equals(row.Check, "FAIL", StringComparison.OrdinalIgnoreCase);
            string cls = fail ? "chk fail-fg" : "chk pass-fg";
            bool mono = row.Name.Contains("GTIN", StringComparison.OrdinalIgnoreCase);
            string dataStyle = mono ? " style=\"font-family:Consolas,monospace;\"" : "";
            sb.Append($"              <tr><td>{H(row.Name)}</td><td{dataStyle}>{H(row.Data)}</td><td class=\"{cls}\">{H(row.Check)}</td></tr>\n");
        }
        sb.Append("            </tbody>\n");
        sb.Append("          </table>\n");
        sb.Append($"          <div class=\"barcode-dfc-footer\"><span class=\"overall-pill {pillCls}\">{pillTxt}</span></div>\n");
    }

    private static void AppendVccsDigitalLinkValidation(
        StringBuilder sb,
        DigitalLinkValidationResult? validation)
    {
        DigitalLinkValidationStatus status =
            validation?.Status ?? DigitalLinkValidationStatus.Unavailable;
        string detail = validation?.Detail ??
            "VCCS validation was not calculated for this record.";
        (string cls, string label) = status switch
        {
            DigitalLinkValidationStatus.Valid => ("pass-fg", "PASS"),
            DigitalLinkValidationStatus.Invalid => ("fail-fg", "FAIL"),
            DigitalLinkValidationStatus.NotApplicable => ("", "NOT APPLICABLE"),
            _ => ("", "UNAVAILABLE"),
        };
        string engine = string.IsNullOrWhiteSpace(validation?.EngineVersion)
            ? "VCCS validation"
            : validation.EngineVersion!;

        sb.Append("          <div class=\"sec-note\" style=\"margin:7pt 0 3pt 0;\"><strong>VCCS / GS1 Digital Link syntax validation</strong></div>\n");
        sb.Append("          <table class=\"dfc-table\"><thead><tr><th>Source</th><th>Detail</th><th class=\"chk\">Check</th></tr></thead><tbody>\n");
        sb.Append($"            <tr><td>{H(engine)}</td><td>{H(detail)}</td><td class=\"chk {cls}\">{label}</td></tr>\n");
        sb.Append("          </tbody></table>\n");
    }

    public static bool HasCorrelatedFilesystemHtml(VerificationRecord r)
        => r.HtmlReportProvenance == HtmlReportProvenance.CorrelatedFilesystem &&
           !string.IsNullOrWhiteSpace(r.HtmlSourceFileName) &&
           !string.IsNullOrWhiteSpace(r.HtmlVerifiedString);

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
            1 => $"Fail &#x2014; {parts[0]}",
            2 => $"Fail &#x2014; {parts[0]} and {parts[1]}",
            _ => $"Fail &#x2014; {string.Join(", ", parts[..^1])} and {parts[^1]}",
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
