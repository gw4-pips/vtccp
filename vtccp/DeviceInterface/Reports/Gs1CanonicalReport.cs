using System.Globalization;
using System.Net;
using System.Text;
using ExcelEngine.Models;

namespace DeviceInterface.Reports;

/// <summary>
/// Dependency-free implementation of the GS1 General Specifications Release 26.0
/// report templates (5.12.7.3 and 5.12.7.4).  It deliberately does not infer,
/// normalise, grade, or substitute verifier values.
/// </summary>
public static class Gs1CanonicalReport
{
    public const string ReportVersion = "v1.0.0";
    public const string Release = "Release 26.0, Ratified, Jan 26";

    public static Gs1ReportValidationResult Validate(Gs1ReportData report)
    {
        ArgumentNullException.ThrowIfNull(report);
        var missing = new List<string>();
        AddSummaryRequirements(report, missing);
        return new Gs1ReportValidationResult(missing);
    }

    /// <summary>Validates the summary plus the mandatory section 5.12.7.3 evidence rows.</summary>
    public static Gs1ReportValidationResult ValidateLinear(Gs1ReportData report)
    {
        ArgumentNullException.ThrowIfNull(report);
        var missing = new List<string>();
        AddSummaryRequirements(report, missing);
        AddIsoEvidence(report, missing, "Overall ISO/IEC grade", "Decode", "Symbol contrast",
            "Rmin (minimum reflectance)", "Edge contrast", "Modulation", "Defects",
            "Decodability", "Print growth (+/- %) – Process control parameter");
        AddGs1AssessedEvidence(report, missing, "X-dimension");
        AddGs1ComplianceEvidence(report, missing, "Data structure¹ (syntax)",
            "Validity of GS1 Company Prefix", "Human readable");
        return new Gs1ReportValidationResult(missing);
    }

    /// <summary>Validates the summary plus the mandatory section 5.12.7.4 evidence rows.</summary>
    public static Gs1ReportValidationResult ValidateTwoDimensional(Gs1ReportData report)
    {
        ArgumentNullException.ThrowIfNull(report);
        var missing = new List<string>();
        AddSummaryRequirements(report, missing);
        AddIsoEvidence(report, missing, "Overall ISO/IEC grade", "Decode", "Symbol contrast",
            "Modulation", "Axial nonuniformity", "Grid nonuniformity", "Unused Error Correction",
            "Fixed pattern damage");
        AddGs1AssessedEvidence(report, missing, "Matrix size", "X-dimension");
        AddGs1ComplianceEvidence(report, missing, "Data structure¹ (syntax)",
            "Validity of GS1 Company Prefix", "Human readable");
        return new Gs1ReportValidationResult(missing);
    }

    private static void AddSummaryRequirements(Gs1ReportData report, ICollection<string> missing)
    {
        Required(report.IssueDate, "Issue date", missing);
        Required(report.TestingAgencyName, "Name", missing);
        Required(report.ProductDescription, "Product Description", missing);
        Required(report.BarcodeType, "Type of barcode", missing);
        Required(report.PrintMethod, "Print method", missing);
        Required(report.NumberOfBarcodesOnProduct, "Number of barcodes on product", missing);
        Required(report.VerifierDevice, "Verifier device", missing);
        Required(report.VerificationSoftwareVersion, "Verification software version", missing);
        Required(report.LastVerifierCalibrationDate, "Last verifier calibration date", missing);
        Required(report.TestedEnvironments, "Tested environments", missing);
        Required(report.SymbolSpecificationTable, "GS1 symbol specification table", missing);
        Required(report.PlacementResult, "Placement result", missing);
        Required(report.TwoDimensionalProximity, "Yes/No", missing);
        Required(report.OverallIsoIecGrade, "ISO/IEC grade", missing);
        Required(report.DecodedText, "Decoded text", missing);
    }

    private static void AddIsoEvidence(Gs1ReportData report, ICollection<string> missing, params string[] labels)
    {
        foreach (string label in labels)
        {
            if (!report.IsoParameters.TryGetValue(label, out Gs1IsoAssessment? assessment) ||
                string.IsNullOrWhiteSpace(assessment?.LiteralValue))
                missing.Add($"ISO/IEC evidence: {label}");
        }
    }

    private static void AddGs1ComplianceEvidence(
        Gs1ReportData report, ICollection<string> missing, params string[] labels)
    {
        foreach (string label in labels)
        {
            if (!report.Gs1Parameters.TryGetValue(label, out Gs1ParameterAssessment? assessment) ||
                string.IsNullOrWhiteSpace(assessment?.ComplianceResult))
                missing.Add($"GS1 compliance evidence: {label}");
        }
    }

    private static void AddGs1AssessedEvidence(
        Gs1ReportData report, ICollection<string> missing, params string[] labels)
    {
        foreach (string label in labels)
        {
            if (!report.Gs1Parameters.TryGetValue(label, out Gs1ParameterAssessment? assessment) ||
                string.IsNullOrWhiteSpace(assessment?.LiteralAssessedValue))
                missing.Add($"GS1 assessed evidence: {label}");
        }
    }

    /// <summary>Renders the linear-barcode template in section 5.12.7.3.</summary>
    public static string GenerateLinearHtml(Gs1ReportData report)
    {
        ArgumentNullException.ThrowIfNull(report);
        return GenerateLinearHtml(report, Gs1PrintProfile.A4);
    }

    /// <summary>Renders using the selected print profile; A4 is the canonical default.</summary>
    public static string GenerateLinearHtml(Gs1ReportData report, Gs1PrintProfile printProfile)
    {
        ArgumentNullException.ThrowIfNull(report);
        return Generate(report, false, ValidateLinear(report), printProfile);
    }

    /// <summary>Renders the two-dimensional-barcode template in section 5.12.7.4.</summary>
    public static string GenerateTwoDimensionalHtml(Gs1ReportData report)
    {
        ArgumentNullException.ThrowIfNull(report);
        return GenerateTwoDimensionalHtml(report, Gs1PrintProfile.A4);
    }

    /// <summary>Renders using the selected print profile; A4 is the canonical default.</summary>
    public static string GenerateTwoDimensionalHtml(Gs1ReportData report, Gs1PrintProfile printProfile)
    {
        ArgumentNullException.ThrowIfNull(report);
        return Generate(report, true, ValidateTwoDimensional(report), printProfile);
    }

    /// <summary>
    /// Explicit adapter for existing verifier records. Values are copied literally;
    /// fields which the verifier does not supply remain missing and validation exposes
    /// them rather than inventing a compliant-looking value.
    /// </summary>
    public static Gs1ReportData FromVerificationRecord(VerificationRecord record) =>
        new()
        {
            IssueDate = record.VerificationDateTime.ToString("d", CultureInfo.InvariantCulture),
            ProductDescription = record.ProductName,
            BarcodeType = record.Symbology,
            VerifierDevice = record.DeviceModel ?? record.DeviceName ?? record.VerifierBrand,
            VerificationSoftwareVersion = record.SoftwareVersion ?? record.FirmwareVersion,
            LastVerifierCalibrationDate = record.CalibrationDate?.ToString("d", CultureInfo.InvariantCulture),
            DecodedText = record.HtmlDecodedData ?? record.DecodedData,
            OverallIsoIecGrade = record.HtmlOverallGradeDisplay ?? record.HtmlLinearGradeDisplay ??
                                 record.OverallGrade?.ToString(),
            IsoStandard = record.HtmlStandard ?? record.HtmlLinearStandard ?? record.Standard,
            Rfid = new Gs1RfidSupplement
            {
                Status = record.RfidStatus,
                EpcTagUri = record.RfidEpcTagUri,
                EpcHex = record.RfidEpcHex,
                Tid = record.RfidTid,
                Gtin14 = record.RfidGtin14,
                Serial = record.RfidSerial,
                Detail = record.RfidMismatchDetail
            }
        };

    private static void Required(string? value, string label, ICollection<string> missing)
    {
        if (string.IsNullOrWhiteSpace(value)) missing.Add(label);
    }

    private static string Generate(Gs1ReportData r, bool twoDimensional, Gs1ReportValidationResult validation,
        Gs1PrintProfile printProfile)
    {
        string kind = twoDimensional ? "two-dimensional" : "linear";
        var sb = new StringBuilder();
        string paper = printProfile == Gs1PrintProfile.Letter ? "Letter" : "A4";
        string pageMargin = printProfile == Gs1PrintProfile.Letter ? "8mm" : "12mm";
        string bodyClass = printProfile == Gs1PrintProfile.Letter ? "letter" : "a4";
        sb.Append($$"""
<!doctype html><html><head><meta charset="utf-8"><style>
@page { size:{{paper}}; margin:{{pageMargin}}; } @media print { .page { width:auto; min-height:auto; margin:0; border:0; } }
* { box-sizing:border-box } body { font:10pt Arial,sans-serif; color:#000; margin:0; background:#eee }
.page { width:186mm; min-height:273mm; margin:8mm auto; padding:0; background:#fff } h1 { font-size:15pt; margin:0 0 4mm }
h2 { font-size:12pt; margin:5mm 0 2mm } h3 { font-size:10pt; margin:3mm 0 1mm }
table { width:100%; border-collapse:collapse; margin:2mm 0 } th,td { border:1px solid #222; padding:1.5mm; vertical-align:top } th { background:#e9e9e9; text-align:left }
.identity td:first-child { width:43%; font-weight:bold }.warning { border:2px solid #b00020; padding:2mm; color:#b00020; font-weight:bold }.columns{display:grid;grid-template-columns:1fr 1fr;gap:5mm}.small{font-size:8.5pt}.rfid{border:1px solid #555;padding:2mm;margin-top:5mm}.missing{color:#b00020;font-weight:bold}
.technical-page { break-before:page; page-break-before:always; font-size:8pt; line-height:1.08 }
.technical-page h2 { font-size:10.5pt; margin:2mm 0 1mm }.technical-page h3 { font-size:8.5pt; margin:1.5mm 0 .5mm }
.technical-page table { margin:1mm 0 }.technical-page th,.technical-page td { padding:.65mm 1mm }
.technical-page p { margin:1mm 0 }.technical-page .small { font-size:7.2pt }
.technical-page .rfid { margin-top:2mm; padding:1mm }.technical-page .rfid h2 { margin-top:0 }
.technical-page .columns { gap:3mm }.provenance td { border:0; padding:.25mm 1mm }
.letter .technical-page { font-size:7.2pt; line-height:1 }.letter .technical-page h2 { font-size:9.5pt; margin:1.4mm 0 .7mm }
.letter .technical-page h3 { font-size:7.8pt; margin:1mm 0 .3mm }.letter .technical-page th,.letter .technical-page td { padding:.45mm .8mm }
.letter .technical-page p { margin:.6mm 0 }.letter .technical-page .small { font-size:6.4pt }.letter .technical-page .rfid { margin-top:1mm; padding:.7mm }
</style></head><body class="{{bodyClass}}"><main class="page">
""");
        sb.Append("<h1>GS1 barcode verification template for ").Append(kind).Append(" barcodes</h1>");
        if (!validation.IsValid)
            sb.Append("<div class=\"warning\">INVALID REPORT — mandatory fields missing: ")
              .Append(H(string.Join("; ", validation.MissingFields))).Append("</div>");
        sb.Append("<table class=\"identity\">")
          .Append(Row("Issue date:", V(r.IssueDate, "Issue date")))
          .Append(Row("Name:", V(r.TestingAgencyName, "Name") + Br(r.TestingAgencyAddress)))
          .Append(Row("Product Description:", V(r.ProductDescription, "Product Description")))
          .Append(Row("Type of barcode:", V(r.BarcodeType, "Type of barcode")))
          .Append(Row("Print method:", V(r.PrintMethod, "Print method")))
          .Append(Row("Number of barcodes on product:", V(r.NumberOfBarcodesOnProduct, "Number of barcodes on product")))
          .Append(Row("Verifier device" + (twoDimensional ? "" : ":") , V(r.VerifierDevice, "Verifier device")))
          .Append(Row("Verification software version:", V(r.VerificationSoftwareVersion, "Verification software version")))
          .Append(Row("Last verifier calibration date:", V(r.LastVerifierCalibrationDate, "Last verifier calibration date")))
          .Append("</table><h3>Please Note:</h3><p>These assessments are based on meeting the minimum GS1 standards.<br>To ensure efficient scanning, the barcode should exceed the minimum.</p>");
        sb.Append("<h2>Testing summary of the ").Append(kind).Append(" barcode</h2>")
          .Append("<p>GS1 General Specifications for ").Append(kind).Append(twoDimensional ? " barcodes, environments tested:" : " barcodes tested environments:")
          .Append("<br>").Append(V(r.TestedEnvironments, "Tested environments")).Append("</p>")
          .Append("<p>PASS or FAIL when verified in accordance with GS1 symbol specification table ")
          .Append(V(r.SymbolSpecificationTable, "GS1 symbol specification table")).Append("</p>")
          .Append("<p class=\"small\">Note: Provide symbol specification table name and ")
          .Append(twoDimensional ? "number" : "table number").Append(" used from the GS1 General Specifications. If barcode is tested against more than one symbol specification table, one report should be provided for each table.</p>");
        sb.Append("<table><tr><th>").Append(twoDimensional ? "Complies to GS1 barcode placement recommendations" : "Complies with GS1 barcode placement rules")
          .Append("</th><td>").Append(V(r.PlacementResult, "Placement result")).Append("</td></tr>")
          .Append(Row("If multiple retail barcodes are present: Is the GS1 compliant two-dimensional barcode within a 50 mm radius from the centre of the linear POS barcode? (informative only)", V(r.TwoDimensionalProximity, "Yes/No")))
          .Append(Row(twoDimensional ? "Overall ISO/IEC 15415 print quality grade" : "ISO/IEC 15416 print quality grade", V(r.OverallIsoIecGrade, "ISO/IEC grade")))
          .Append(Row("Decoded text", V(r.DecodedText, "Decoded text")))
          .Append(Row("Business critical comments", H(r.BusinessCriticalComments))).Append("</table>");
        sb.Append("<section class=\"technical-page\">");
        AppendAnalysis(sb, r, twoDimensional);
        AppendRfid(sb, r.Rfid);
        AppendProvenance(sb, r.Provenance);
        AppendNotes(sb, twoDimensional);
        return sb.Append("<p class=\"small\">Canonical report version ").Append(ReportVersion)
          .Append("</p></section></main></body></html>").ToString();
    }

    private static void AppendAnalysis(StringBuilder sb, Gs1ReportData r, bool twoD)
    {
        string[] iso = twoD
            ? ["Overall ISO/IEC grade", "Decode", "Symbol contrast", "Modulation", "Axial nonuniformity", "Grid nonuniformity", "Unused Error Correction", "Print growth (horizontal)", "Print growth (vertical)", "Fixed pattern damage", "• Clock track and solid area regularity³", "• Quite Zones (QZL1, QZL2)³", "• L1 and L2³", "• Format information⁴", "• Version information⁴"]
            : ["Overall ISO/IEC grade", "Decode", "Symbol contrast", "Rmin (minimum reflectance)", "Rmax (maximum reflectance)", "Edge contrast", "Modulation", "Defects", "Decodability", "Print growth (+/- %) – Process control parameter"];
        sb.Append("<h2>Technical analysis of the ").Append(twoD ? "two-dimensional" : "linear").Append(" barcode</h2><table><tr><th>ISO/IEC parameters</th><th>Values</th><th>Comment Reference</th></tr>");
        foreach (string parameter in iso)
        {
            r.IsoParameters.TryGetValue(parameter, out Gs1IsoAssessment? value);
            sb.Append("<tr><td>").Append(parameter).Append("</td><td>").Append(H(value?.LiteralValue))
              .Append("</td><td>").Append(H(value?.CommentReference)).Append("</td></tr>");
        }
        sb.Append("</table><table><tr><th>GS1 parameters</th><th>Required</th><th>")
          .Append(twoD ? "Compliant to standard" : "Within standard range").Append("</th><th>Assessed</th><th>Comment reference</th></tr>")
          .Append(Gs1Row(r.Gs1Parameters, twoD ? "Matrix size" : "Barcode structure"))
          .Append(Gs1Row(r.Gs1Parameters, "X-dimension"))
          .Append(Gs1Row(r.Gs1Parameters, "Data structure¹ (syntax)"))
          .Append(Gs1Row(r.Gs1Parameters, "Validity of GS1 Company Prefix"))
          .Append(Gs1Row(r.Gs1Parameters, "Human readable"))
          .Append("</table><h3>Educational comments²").Append(twoD ? ":" : "").Append("</h3><p>").Append(H(r.EducationalComments)).Append("</p>");
        sb.Append(twoD
            ? "<p class=\"small\">(1) Data structure (syntax) indicates that the barcode is compliant with GS1 data syntax rules defined in the GS1 General Specifications or GS1 Digital Link URI Syntax standard.<br>(2) Educational comments are based on the technical analysis of the barcode. In this comment box the operator comments on what the problem is and how to make the barcode better by explaining the parameter’s meanings.<br>(3) Data Matrix Only, see ISO/IEC 16022<br>(4) QR Code Only, see ISO/IEC 18004</p>"
            : "<p class=\"small\">(1) Data structure (syntax) indicates that the barcode is compliant with GS1 data syntax rules defined in the General Specifications.<br>(2) Educational comments are based on the technical analysis of the barcode. In this comment box the operator comments on what the problem is and how to make the symbol better.</p>");
    }

    private static void AppendRfid(StringBuilder sb, Gs1RfidSupplement? rfid)
    {
        if (rfid is null) return;
        sb.Append("<section class=\"rfid\"><h2>RFID supplemental information</h2><p>This supplemental section is independent of GS1 and ISO/IEC barcode verification outcomes.</p><table>")
          .Append(Row("RFID status", H(rfid.Status))).Append(Row("EPC Tag URI", H(rfid.EpcTagUri)))
          .Append(Row("EPC (hex)", H(rfid.EpcHex))).Append(Row("TID", H(rfid.Tid)))
          .Append(Row("GTIN-14", H(rfid.Gtin14))).Append(Row("Serial", H(rfid.Serial)))
          .Append(Row("Detail", H(rfid.Detail))).Append("</table></section>");
    }

    private static void AppendProvenance(StringBuilder sb, Gs1ReportProvenance? provenance)
    {
        if (provenance is null) return;
        sb.Append("<section class=\"small\"><h3>Report provenance</h3><table class=\"provenance\">")
          .Append(Row("Organization", H(provenance.OrganizationName)))
          .Append(Row("Job", H(provenance.JobName)))
          .Append(Row("Verifier source", H(provenance.VerifierSource)))
          .Append(Row("Native source artifact", H(provenance.SourceArtifactPath)))
          .Append(Row("RFID source", H(provenance.RfidSource)))
          .Append("</table></section>");
    }

    private static void AppendNotes(StringBuilder sb, bool twoD)
    {
        string secondaryHeading = twoD
            ? "Important Note (normative localised)"
            : "Notes (informative localised)";
        sb.Append($$"""
<section class="columns small"><div><h3>Notes (informative localised)</h3><p>It is the responsibility of the GS1 identification licensee to ensure the correct use of the GS1 Company Prefix and/or the individually licensed keys and the correct allocation of the data content.</p><p>Rejection of products should not necessarily be based only on an out of specification results</p><p>Barcode verifiers are measuring devices and are tools that can be used for assisting in quality control. The results are not absolute in that they do not necessarily prove or disprove that the barcode will scan.</p><p>This report may not be amended after issue. In the event of a dispute over contents the version held at [TESTING AGENCY] will be deemed to be the correct and original version of this report.</p></div><div><h3>{{secondaryHeading}}</h3><p>This Verification Report may contain privileged and confidential information intended only for the use of the addressee named above. If you are not the intended recipient of this report you are hereby notified that any use, dissemination, distribution or reproduction of this message is prohibited. If you received this message in error please notify [TESTING AGENCY].</p><h3>Disclaimer (legal localised)</h3><p>This report does not constitute evidence for the purpose of any litigation, and [TESTING AGENCY] will not enter into any discussion, or respond to any correspondence in relation to litigation.</p><p>Every possible effort has been made to ensure that the information and specifications in the Barcode Verification Reports are correct, however, [TESTING AGENCY] expressly disclaims liability for any errors.</p></div></section><p class="small">{{Release}}<br>© 2026 GS1 AISBL</p>
""");
    }

    private static string Gs1Row(IReadOnlyDictionary<string, Gs1ParameterAssessment> parameters, string name)
    {
        parameters.TryGetValue(name, out Gs1ParameterAssessment? value);
        return "<tr><td>" + name + "</td><td>" + V(value?.NormativeRequirement, "Normative requirement") + "</td><td>" +
               V(value?.ComplianceResult, "Compliance result") + "</td><td>" +
               H(value?.LiteralAssessedValue) + "</td><td>" + H(value?.CommentReference) + "</td></tr>";
    }
    private static string Row(string name, string value) => "<tr><td>" + name + "</td><td>" + value + "</td></tr>";
    private static string V(string? value, string label) => string.IsNullOrWhiteSpace(value) ? "<span class=\"missing\">[MISSING: " + H(label) + "]</span>" : H(value);
    private static string Br(string? value) => string.IsNullOrWhiteSpace(value) ? "" : "<br>" + H(value);
    private static string H(string? value) => WebUtility.HtmlEncode(value ?? "");
}

public sealed class Gs1ReportData
{
    public string? IssueDate { get; init; }
    public string? TestingAgencyName { get; init; }
    public string? TestingAgencyAddress { get; init; }
    public string? ProductDescription { get; init; }
    public string? BarcodeType { get; init; }
    public string? PrintMethod { get; init; }
    public string? NumberOfBarcodesOnProduct { get; init; }
    public string? VerifierDevice { get; init; }
    public string? VerificationSoftwareVersion { get; init; }
    public string? LastVerifierCalibrationDate { get; init; }
    public string? TestedEnvironments { get; init; }
    public string? SymbolSpecificationTable { get; init; }
    public string? PlacementResult { get; init; }
    public string? TwoDimensionalProximity { get; init; }
    public string? OverallIsoIecGrade { get; init; }
    public string? DecodedText { get; init; }
    public string? BusinessCriticalComments { get; init; }
    public string? EducationalComments { get; init; }
    public string? IsoStandard { get; init; }
    /// <summary>Literal technical values and references keyed by the canonical ISO row label.</summary>
    public IReadOnlyDictionary<string, Gs1IsoAssessment> IsoParameters { get; init; } =
        new Dictionary<string, Gs1IsoAssessment>();
    /// <summary>
    /// GS1 row evidence keyed by canonical row label. NormativeRequirement is template
    /// requirement/range text, never a verifier result; ComplianceResult is the plugin result.
    /// </summary>
    public IReadOnlyDictionary<string, Gs1ParameterAssessment> Gs1Parameters { get; init; } =
        new Dictionary<string, Gs1ParameterAssessment>();
    public Gs1RfidSupplement? Rfid { get; init; }
    public Gs1ReportProvenance? Provenance { get; init; }
}

/// <summary>Keeps organization, job, verifier, RFID, and native artifact origins distinct.</summary>
public sealed class Gs1ReportProvenance
{
    public string? OrganizationName { get; init; }
    public string? JobName { get; init; }
    public string? VerifierSource { get; init; }
    public string? SourceArtifactPath { get; init; }
    public string? RfidSource { get; init; }
}

public sealed class Gs1RfidSupplement
{
    public string? Status { get; init; }
    public string? EpcTagUri { get; init; }
    public string? EpcHex { get; init; }
    public string? Tid { get; init; }
    public string? Gtin14 { get; init; }
    public string? Serial { get; init; }
    public string? Detail { get; init; }
}

public sealed class Gs1ParameterAssessment
{
    public string? NormativeRequirement { get; init; }
    public string? ComplianceResult { get; init; }
    public string? LiteralAssessedValue { get; init; }
    public string? CommentReference { get; init; }
    public string? Provenance { get; init; }
}

public sealed class Gs1IsoAssessment
{
    public string? LiteralValue { get; init; }
    public string? CommentReference { get; init; }
    public string? Provenance { get; init; }
}

public sealed class Gs1ReportValidationResult
{
    internal Gs1ReportValidationResult(IReadOnlyList<string> missingFields) => MissingFields = missingFields;
    public IReadOnlyList<string> MissingFields { get; }
    public bool IsValid => MissingFields.Count == 0;
}

/// <summary>Browser print paper profiles. A4 is required unless Letter is selected explicitly.</summary>
public enum Gs1PrintProfile
{
    A4,
    Letter
}