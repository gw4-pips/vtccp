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
    public const string ReportVersion = "v1.5.47";
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
            .Replace("{{APP_SETTINGS_INLINE}}", BuildApplicationSettingsInline(r))
            .Replace("{{SLOT_SYMBOL_ROWS}}",    BuildSymbolRows(r))
            .Replace("{{REPORT_NAME}}",         reportName)
            .Replace("{{REPORT_DATETIME}}",     reportDateTime)
            .Replace("{{SLOT_GRADE_ROWS}}",     BuildGradeRows(r))
            .Replace("{{RFID_ADJ}}",            rfidAdj)
            .Replace("{{SLOT_RFID_ROWS}}",      BuildRfidRows(r))
            .Replace("{{SLOT_BARCODE_DETAIL}}", BuildBarcodeDetailSection(r))
            .Replace("{{FOOTER_VERSION}}",      ReportVersion)
            .Replace("{{APP_VERSION}}",         GetApplicationVersion())
            .Replace("{{FOOTER_GENERATED}}",    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
    }

    private static string GetApplicationVersion()
    {
        Version? version = System.Reflection.Assembly.GetEntryAssembly()?.GetName().Version;
        return version is null ? "unknown" : $"v{version.ToString(3)}";
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
            AppendSymbolRow(sb, r.LinearSymbology!, r.LinearDecodedData);
            renderedGroups++;
        }
        if (renderedGroups < MaxRenderedSymbolGroups)
        {
            AppendSymbolRow(sb, r.HtmlSymbology ?? "\u2014", r.HtmlDecodedData);
        }
        return sb.ToString();
    }

    private static string UnavailableSymbolRow(VerificationRecord r)
        => "          <tr>\n" +
           "            <td>[UNAVAILABLE]</td>\n" +
           "            <td colspan=\"2\">[UNAVAILABLE — NOT PRESENT IN CORRELATED TRUCHECK HTML]</td>\n" +
           "          </tr>\n";

    private static void AppendSymbolRow(StringBuilder sb, string symb, string? encoded)
    {
        sb.Append($"          <tr data-vccs-symbol-group=\"true\">\n");
        sb.Append($"            <td style=\"font-size:8pt;\">{H(symb)}</td>\n");
        sb.Append($"            <td colspan=\"2\" style=\"font-family:Consolas,monospace;\">{H(encoded ?? "\u2014")}</td>\n");
        sb.Append($"          </tr>\n");
    }

    private static string BuildApplicationSettingsInline(VerificationRecord r)
    {
        // These are live TruCheck configuration values queried after each
        // completed result. Do not substitute values inferred from the correlated
        // HTML's Data Format Check results or from the numeric grade aperture.
        string applicationStandard = r.ApplicationStandardSetting ?? "\u2014";
        string dataFormatCheck = r.DataFormatCheckSetting ??
                                 (HasCorrelatedFilesystemHtml(r)
                                     ? DisplayDataFormatCheckSetting(r.HtmlDataFormatCheck)
                                     : "\u2014");
        string apertureSetting = r.ApertureSettingMode ?? "\u2014";

        return "<span class=\"app-settings-label\">Application Std. / Data Format Check / Aperture:</span> " +
               $"<span class=\"app-settings-values\">{H(applicationStandard)}" +
               $" / {H(dataFormatCheck)} / {H(apertureSetting)}</span>";
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

        string gcpValidLenPart = r.RfidGcpLength.HasValue
            ? $" (={r.RfidGcpLength.Value})"
            : string.Empty;
        string gcpInvalidLenPart = r.RfidGcpLength.HasValue
            ? $" ({r.RfidGcpLength.Value})"
              + (r.RfidGcpRegisteredLength.HasValue
                  ? $"; Valid = {r.RfidGcpRegisteredLength.Value}"
                  : string.Empty)
            : string.Empty;
        string gcpDisplay = r.RfidGcpStatus switch
        {
            "Valid" => $"Valid{gcpValidLenPart}",
            "Invalid" => $"Invalid{gcpInvalidLenPart}",
            "NotFound" => $"NOT FOUND{gcpValidLenPart}",
            "NotChecked" => "\u2014",
            _ => r.RfidGcpValid switch
            {
                true  => $"Valid{gcpValidLenPart}",
                false => $"Invalid{gcpInvalidLenPart}",
                null  => "\u2014",
            },
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
        string resultLabel = IsRfidCrossValidation(r)
            ? "RFID Cross-Validation Result"
            : "RFID Validation Result";

        void ResultRow(string label, string value)
        {
            sb.Append($"          <tr class=\"{rowCls}\">\n");
            sb.Append($"            <td class=\"rfid-result-label\">{H(label)}</td>\n");
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
            ResultRow($"{r.LinearSymbology} {resultLabel}", linVal);

            string twoDSym = string.IsNullOrWhiteSpace(r.Symbology) ? "2D" : r.Symbology;
            string twoDVal = r.RfidStatus switch
            {
                "Pass" => "Pass &#x2014; EPC data matches barcode GTIN and Serial Number",
                "Fail" => BuildMismatch2DLabel(r.RfidMismatchDetail),
                var s  => H(s ?? "\u2014"),
            };
            ResultRow($"{twoDSym} {resultLabel}", twoDVal);
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
            ResultRow($"{symName} {resultLabel}", singleVal);
        }

        return sb.ToString();
    }

    private static string BuildBarcodeDetailSection(VerificationRecord r)
    {
        bool hasHtml       = HasCorrelatedFilesystemHtml(r);
        string? img2D      = hasHtml ? r.HtmlBarcodeImageBase64 : null;
        string imageMimeType = hasHtml
            ? (r.HtmlBarcodeImageMimeType ?? "image/jpeg")
            : "image/jpeg";
        string imageSourceNote = r.HtmlBarcodeImageProvenance switch
        {
            "SiblingExport" =>
                " &#x2014; <em>Image1 sibling export; not embedded in the HTML</em>",
            "EmbeddedHtml" =>
                " &#x2014; <em>image referenced by the HTML export</em>",
            _ => string.Empty,
        };
        bool multiMode    = !string.IsNullOrWhiteSpace(r.LinearSymbology);
        DataFormatCheckResult? htmlDfc = hasHtml ? r.HtmlDataFormatCheck : null;
        bool veriWedgeValidationUsed = IsVeriWedgeValidationUsed(r);
        bool useElementStringLayout = veriWedgeValidationUsed &&
            IsElementStringValidation(r.VccsDigitalLinkValidation);
        bool useParserComparisonLayout = veriWedgeValidationUsed;
        string? nativeDigitalLinkSupportNote = useParserComparisonLayout && !useElementStringLayout
            ? BuildNativeDigitalLinkSupportNote(r, htmlDfc, r.VccsDigitalLinkValidation)
            : null;

        var sb = new StringBuilder();
        sb.Append("    <div class=\"barcode-detail-section\">\n");
        sb.Append(hasHtml
            ? useParserComparisonLayout
                ? $"      <div class=\"sec-sub-hdr trucheck-barcode-hdr barcode-detail-header barcode-dual-header\">{BuildVeriWedgeDfcHeader(r, imageSourceNote)}</div>\n"
                : $"      <div class=\"sec-sub-hdr trucheck-barcode-hdr barcode-detail-header\"><span class=\"trucheck-header-title\">TruCheck Barcode Image <span class=\"detail-separator\">|</span> Data Format Check &#x2014; GS1</span><span class=\"sec-note\"> &#x2014; <em>Native TruCheck data and VCCS Digital Link validation are separately labelled</em>{imageSourceNote}</span></div>\n"
            : "      <div class=\"sec-sub-hdr trucheck-barcode-hdr barcode-detail-header\"><span class=\"trucheck-header-title\">Barcode Verification Capture Unavailable</span><span class=\"sec-note\"> &#x2014; <em>No correlated DMST HTML report</em></span></div>\n");
        sb.Append("      <table class=\"barcode-detail-grid\"><tbody><tr>\n");
        sb.Append("        <td class=\"barcode-image-column\">\n");

        if (hasHtml && !string.IsNullOrWhiteSpace(img2D))
        {
            sb.Append($"          <img class=\"barcode-image\" src=\"data:{imageMimeType};base64,{img2D}\" alt=\"TruCheck Barcode Image\"/>\n");
        }
        else
        {
            sb.Append(hasHtml
                ? "        <div class=\"img-placeholder\">[BARCODE IMAGE NOT EMBEDDED IN CORRELATED TRUCHECK HTML]</div>\n"
                : "        <div class=\"img-placeholder\">[BARCODE IMAGE UNAVAILABLE — NO CORRELATED DMST HTML]</div>\n");
        }

        sb.Append("        </td>\n");
        sb.Append("        <td class=\"barcode-dfc-column\">\n");

        if (useParserComparisonLayout)
        {
            AppendElementStringDualValidation(
                sb,
                r,
                htmlDfc,
                r.VccsDigitalLinkValidation,
                hasHtml,
                string.Equals(r.DataFormatCheckSetting, "None", StringComparison.OrdinalIgnoreCase),
                nativeDigitalLinkSupportNote);
        }
        else
        {
            AppendVendorDataFormatCheck(sb, htmlDfc, hasHtml);
        }

        sb.Append("        </td>\n");
        sb.Append("      </tr></tbody></table>\n");
        sb.Append("    </div>\n");
        return sb.ToString();
    }

    private static string BuildVeriWedgeDfcHeader(
        VerificationRecord r,
        string imageSourceNote)
    {
        DigitalLinkValidationResult? validation = r.VccsDigitalLinkValidation;
        string algorithm = string.Equals(
            validation?.Source,
            DigitalLinkValidationResult.VccsElementStringSource,
            StringComparison.Ordinal)
            ? "Element String"
            : "Digital Link";

        return "<span class=\"barcode-header-image-title\">TruCheck Barcode Image</span>" +
               "<span class=\"barcode-header-dfc-title\"><span class=\"detail-separator\">|</span> " +
               $"Data Format Check (DFC) &#x2014; GS1 {algorithm}</span>" +
               (imageSourceNote.Length == 0
                   ? string.Empty
                   : $"<span class=\"sec-note\">{imageSourceNote}</span>");
    }

    private static bool IsElementStringValidation(DigitalLinkValidationResult? validation)
        => string.Equals(
            validation?.Source,
            DigitalLinkValidationResult.VccsElementStringSource,
            StringComparison.Ordinal);

    private static bool IsVeriWedgeValidationUsed(VerificationRecord record)
        // Saved records created before explicit provenance retain their
        // populated parser result as evidence that the comparison panel exists.
        // New records always set VeriWedgeValidationUsed, which takes priority.
        => record.VeriWedgeValidationUsed ??
           record.VccsDigitalLinkValidation is not null;

    private static bool IsRfidCrossValidation(VerificationRecord record)
        => IsVeriWedgeValidationUsed(record) &&
           (!record.TruCheckValidationUsable || record.TruCheckValidationFailed);

    private static string? BuildNativeDigitalLinkSupportNote(
        VerificationRecord record,
        DataFormatCheckResult? nativeDfc,
        DigitalLinkValidationResult? validation)
    {
        if (nativeDfc?.Overall != OverallPassFail.Fail ||
            validation?.Status != DigitalLinkValidationStatus.Valid)
        {
            return null;
        }

        if (string.Equals(record.VerifierBrand, "WEBSCAN", StringComparison.OrdinalIgnoreCase) &&
            IsVersionAtOrBefore(record.SoftwareVersion, "3.3.74"))
        {
            return $"Software {record.SoftwareVersion} does not support GS1 Digital Link parsing.";
        }

        // Cognex's DM475V verifier-line release record identifies 6.1.16_sr4
        // (numeric release 6.1.16) as the latest released firmware without
        // GS1 Digital Link parsing support. Pre-release suffixes retain the same
        // numeric compatibility boundary until Cognex publishes a newer release.
        if (IsVersionAtOrBefore(record.FirmwareVersion, "6.1.16"))
        {
            return $"Firmware {record.FirmwareVersion} does not support GS1 Digital Link parsing.";
        }

        return null;
    }

    private static bool IsVersionAtOrBefore(string? value, string supportedThrough)
    {
        if (string.IsNullOrWhiteSpace(value) ||
            !Version.TryParse(supportedThrough, out Version? boundary))
        {
            return false;
        }

        Match match = Regex.Match(value, @"\d+(?:\.\d+){1,3}");
        return match.Success && Version.TryParse(match.Value, out Version? version) &&
            version.CompareTo(boundary) <= 0;
    }

    private static void AppendVendorDataFormatCheck(
        StringBuilder sb,
        DataFormatCheckResult? htmlDfc,
        bool hasCorrelatedHtml)
    {
        sb.Append("          <div class=\"sec-note\" style=\"margin:0 0 3pt 0;\"><strong>Native TruCheck Data Format Check</strong></div>\n");
        if (htmlDfc is null)
        {
            string unavailableReason = hasCorrelatedHtml
                ? "[DATA FORMAT CHECK UNAVAILABLE — NOT PRESENT IN TRUCHECK HTML]"
                : "[DATA FORMAT CHECK UNAVAILABLE — NO DMST HTML REPORT CORRELATED]";
            sb.Append("          <table class=\"dfc-table\"><thead><tr><th>Field</th><th>Data</th><th class=\"chk\">Check</th></tr></thead><tbody>\n");
            sb.Append($"            <tr><td>Source status</td><td>{H(unavailableReason)}</td><td class=\"chk\">UNAVAILABLE</td></tr>\n");
            sb.Append("          </tbody></table>\n");
            return;
        }

        if (!string.IsNullOrWhiteSpace(htmlDfc.Standard))
            sb.Append($"          <div class=\"sec-note\" style=\"margin:0 0 3pt 0;\">Standard: {H(htmlDfc.Standard)}</div>\n");

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
        sb.Append("          <table class=\"dfc-table\"><thead><tr><th>Data</th><th class=\"chk\">Check</th></tr></thead><tbody>\n");
        sb.Append($"            <tr><td>{H(detail)}</td><td class=\"chk {cls}\">{label}</td></tr>\n");
        sb.Append("          </tbody></table>\n");
    }

    private static void AppendElementStringDualValidation(
        StringBuilder sb,
        VerificationRecord record,
        DataFormatCheckResult? htmlDfc,
        DigitalLinkValidationResult? validation,
        bool hasCorrelatedHtml,
        bool noVerifierDfcSelected,
        string? nativeDigitalLinkSupportNote)
    {
        var verifierRows = htmlDfc?.Rows?.ToList() ?? [];
        if (!string.IsNullOrWhiteSpace(htmlDfc?.Standard))
        {
            sb.Append(
                $"          <div class=\"sec-note\" style=\"margin:0 0 3pt 0;\">Native standard: {H(htmlDfc.Standard)}</div>\n");
        }
        if (verifierRows.Count == 0)
        {
            verifierRows.Add(new DataFormatCheckRow
            {
                Name = "Verifier DFC",
                Data = noVerifierDfcSelected
                    ? "No verifier GS1 parser selected."
                    : hasCorrelatedHtml
                    ? "[DATA FORMAT CHECK UNAVAILABLE — NOT PRESENT IN TRUCHECK HTML]"
                    : "[DATA FORMAT CHECK UNAVAILABLE — NO DMST HTML REPORT CORRELATED]",
                Check = noVerifierDfcSelected ? "NOT APPLICABLE" : "UNAVAILABLE",
            });
        }

        string detail = validation?.Detail ??
            "VCCS parsing was not calculated for this record.";
        DigitalLinkValidationStatus status =
            validation?.Status ?? DigitalLinkValidationStatus.Unavailable;
        string parserCheck = status switch
        {
            DigitalLinkValidationStatus.Valid => "PASS",
            DigitalLinkValidationStatus.Invalid => "FAIL",
            DigitalLinkValidationStatus.NotApplicable => "NOT APPLICABLE",
            _ => "UNAVAILABLE",
        };
        string parserClass = parserCheck switch
        {
            "PASS" => "pass-fg",
            "FAIL" => "fail-fg",
            _ => "",
        };
        string? parsedAiData = ExtractParsedAiData(detail);
        List<Gs1ParserRow> parserRows = BuildGs1ParserRows(record, validation, parsedAiData, detail);

        sb.Append("          <div class=\"dfc-accordion-section dfc-dual-block\">\n");
        sb.Append("            <table class=\"dfc-dual-table\">\n");
        sb.Append("              <colgroup><col class=\"dual-left-field\"><col class=\"dual-left-data\"><col class=\"dual-left-check\"><col class=\"dual-divider\"><col class=\"dual-right-field\"><col class=\"dual-right-data\"><col class=\"dual-right-check\"></colgroup>\n");
        sb.Append($"              <thead><tr class=\"dual-subhead\"><th colspan=\"3\">DataMan TruCheck GS1 Parser</th><th class=\"dual-divider\"></th><th colspan=\"3\" class=\"dual-right-parser-header\">VeriWedge GS1 Parser — GS1 Barcode Syntax Engine (v. {H(GetParserVersion(validation))})</th></tr>\n");
        sb.Append("              <tr><th>Field</th><th>Data</th><th>Check</th><th class=\"dual-divider\"></th><th class=\"dual-right-field\">Field</th><th>Data</th><th>Check</th></tr></thead>\n");
        sb.Append("              <tbody>\n");

        int rowCount = Math.Max(verifierRows.Count, parserRows.Count);
        for (int index = 0; index < rowCount; index++)
        {
            DataFormatCheckRow? verifierRow = index < verifierRows.Count ? verifierRows[index] : null;
            Gs1ParserRow? parserRow = index < parserRows.Count ? parserRows[index] : null;
            string? verifierCheck = verifierRow?.Check;
            if (nativeDigitalLinkSupportNote is not null &&
                string.Equals(verifierCheck, "FAIL", StringComparison.OrdinalIgnoreCase))
            {
                verifierCheck = "FAIL*";
            }
            string leftClass = verifierRow?.Check switch
            {
                "PASS" => "pass-fg",
                "FAIL" => "fail-fg",
                _ => "",
            };
            string parserData = parserRow?.IsCanonicalAiString == true
                ? FormatCanonicalAiForHtml(parserRow.Data)
                : H(parserRow?.Data);
            string parserDataClass = parserRow?.IsCanonicalAiString == true
                ? "dual-data parser-element-string-data"
                : "dual-data";
            sb.Append($"                <tr><td>{H(verifierRow?.Name)}</td><td class=\"dual-data\">{H(verifierRow?.Data)}</td><td class=\"dual-check {leftClass}\">{H(verifierCheck)}</td><td class=\"dual-divider\"></td><td class=\"dual-right-field\">{H(parserRow?.Field)}</td><td class=\"{parserDataClass}\">{parserData}</td><td class=\"dual-check {parserClass}\">{(parserRow is not null ? parserCheck : "")}</td></tr>\n");
        }

        (string leftOverallClass, string leftOverallText) = htmlDfc is null
            ? ("pill-warn", "OVERALL: UNAVAILABLE")
            : htmlDfc.Overall switch
            {
                OverallPassFail.Pass => ("pill-pass", "OVERALL: PASS"),
                OverallPassFail.Fail => ("pill-fail", "OVERALL: FAIL"),
                _ => ("pill-warn", "OVERALL: UNAVAILABLE"),
            };
        string parserOverallClass = parserCheck switch
        {
            "PASS" => "pill-pass",
            "FAIL" => "pill-fail",
            _ => "pill-warn",
        };
        string nativeNoteMarkup = string.Empty;
        if (nativeDigitalLinkSupportNote is not null)
        {
            leftOverallText = leftOverallText.Replace("FAIL", "FAIL*", StringComparison.Ordinal);
            leftOverallClass = "pill-native-limitation";
            nativeNoteMarkup = $"<div class=\"dual-native-note\">{H(nativeDigitalLinkSupportNote)}</div>";
        }
        sb.Append($"                <tr class=\"dual-overall\"><td colspan=\"3\" class=\"dual-overall-cell\"><span class=\"overall-pill {leftOverallClass}\">{leftOverallText}</span>{nativeNoteMarkup}</td><td class=\"dual-divider\"></td><td colspan=\"3\" class=\"dual-overall-cell\"><span class=\"overall-pill {parserOverallClass}\">OVERALL: {parserCheck}</span></td></tr>\n");
        sb.Append("              </tbody>\n");
        sb.Append("            </table>\n");
        sb.Append("          </div>\n");
    }

    private static string? ExtractParsedAiData(string detail)
    {
        const string prefix = "Parsed GS1 AI data: ";
        int start = detail.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
        if (start < 0) return null;
        start += prefix.Length;
        int end = detail.IndexOf(" Validated with", start, StringComparison.OrdinalIgnoreCase);
        return (end >= 0 ? detail[start..end] : detail[start..]).Trim();
    }

    private static List<Gs1ParserRow> BuildGs1ParserRows(
        VerificationRecord record,
        DigitalLinkValidationResult? validation,
        string? parsedAiData,
        string detail)
    {
        var rows = new List<Gs1ParserRow>();
        bool isElementString = IsElementStringValidation(validation);

        if (!isElementString)
        {
            rows.Add(new Gs1ParserRow(
                "Web URI",
                FindDigitalLinkUri(record) ?? "[DIGITAL LINK URI NOT AVAILABLE]",
                false));
        }

        foreach ((string ai, string value) in ParseAiElements(parsedAiData))
            rows.Add(new Gs1ParserRow($"AI ({ai}) {GetGs1AiName(ai)}", value, false));

        if (!string.IsNullOrWhiteSpace(parsedAiData))
        {
            rows.Add(new Gs1ParserRow("GS1 Element String", parsedAiData, true));
        }
        else if (rows.Count == 0)
        {
            rows.Add(new Gs1ParserRow("Parser Detail", detail, false));
        }

        return rows;
    }

    private static string? FindDigitalLinkUri(VerificationRecord record)
    {
        foreach (string? candidate in new[] { record.DecodedData, record.HtmlDecodedData })
        {
            if (Uri.TryCreate(candidate, UriKind.Absolute, out Uri? uri) &&
                (uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                 uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
            {
                return candidate;
            }
        }
        return null;
    }

    private static List<(string Ai, string Value)> ParseAiElements(string? parsedAiData)
    {
        var fields = new List<(string Ai, string Value)>();
        if (string.IsNullOrWhiteSpace(parsedAiData)) return fields;

        foreach (Match match in Regex.Matches(
                     parsedAiData,
                     @"\((\d{2,4})\)(.*?)(?=\(\d{2,4}\)|$)",
                     RegexOptions.Singleline))
        {
            fields.Add((match.Groups[1].Value, match.Groups[2].Value));
        }
        return fields;
    }

    private static string FormatCanonicalAiForHtml(string value)
    {
        MatchCollection elements = Regex.Matches(
            value,
            @"\(\d{2,4}\).*?(?=\(\d{2,4}\)|$)",
            RegexOptions.Singleline);
        if (elements.Count == 0) return H(value);

        var formatted = new StringBuilder();
        foreach (Match element in elements)
            formatted.Append(H(element.Value)).Append("<wbr>");
        return formatted.ToString();
    }

    private static string GetGs1AiName(string ai)
        => Gs1AiNames.TryGetValue(ai, out string? name) ? name : "GS1 Application Identifier";

    private sealed record Gs1ParserRow(string Field, string Data, bool IsCanonicalAiString);

    private static readonly IReadOnlyDictionary<string, string> Gs1AiNames =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["00"] = "SSCC",
            ["01"] = "GTIN",
            ["02"] = "GTIN of Contained Trade Item",
            ["10"] = "Batch or Lot Number",
            ["11"] = "Production Date",
            ["12"] = "Due Date",
            ["13"] = "Packaging Date",
            ["15"] = "Best Before Date",
            ["17"] = "Expiration Date",
            ["20"] = "Variant",
            ["21"] = "Serial Number",
            ["22"] = "Consumer Product Variant",
            ["30"] = "Count",
            ["37"] = "Count of Trade Items",
            ["240"] = "Additional Product Identification",
            ["241"] = "Customer Part Number",
            ["250"] = "Secondary Serial Number",
            ["251"] = "Reference to Source Entity",
            ["253"] = "Global Document Type Identifier",
            ["400"] = "Customer Purchase Order Number",
            ["401"] = "Global Identification Number for Consignment",
            ["402"] = "Global Shipment Identification Number",
            ["403"] = "Routing Code",
            ["410"] = "Ship To Global Location Number",
            ["411"] = "Bill To Global Location Number",
            ["412"] = "Purchased From Global Location Number",
            ["413"] = "Ship For Global Location Number",
            ["414"] = "Physical Location Global Location Number",
            ["415"] = "Pay To Global Location Number",
            ["416"] = "Production or Service Location",
            ["417"] = "Party Global Location Number",
            ["420"] = "Ship To Postal Code",
            ["421"] = "Ship To Postal Code with Country Code",
            ["422"] = "Country of Origin",
            ["423"] = "Country of Initial Processing",
            ["424"] = "Country of Processing",
            ["425"] = "Country of Disassembly",
            ["426"] = "Country of Full Processing",
            ["7001"] = "NATO Stock Number",
            ["7003"] = "Expiration Date and Time",
            ["7004"] = "Active Potency",
            ["7006"] = "First Freeze Date",
            ["7007"] = "Harvest Date",
            ["7009"] = "Fishing Gear Type",
            ["7010"] = "Production Method",
            ["7020"] = "Refurbishment Lot Identifier",
            ["7021"] = "Functional Status",
            ["7022"] = "Revision Status",
        };

    private static string GetParserVersion(DigitalLinkValidationResult? validation)
    {
        Match match = Regex.Match(
            validation?.EngineVersion ?? string.Empty,
            @"\b\d+\.\d+\.\d+\b");
        return match.Success ? match.Value : "1.4.1";
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
