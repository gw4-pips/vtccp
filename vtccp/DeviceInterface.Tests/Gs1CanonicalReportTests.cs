using DeviceInterface.Reports;
using Xunit;

namespace DeviceInterface.Tests;

public sealed class Gs1CanonicalReportTests
{
    [Fact]
    public void Validation_lists_all_mandatory_missing_fields()
    {
        Gs1ReportValidationResult result = Gs1LinearHtmlReportGenerator.Validate(new Gs1ReportData());

        Assert.False(result.IsValid);
        Assert.Contains("Print method", result.MissingFields);
        Assert.Contains("GS1 symbol specification table", result.MissingFields);
    }

    [Fact]
    public void Linear_generator_uses_canonical_labels_and_makes_missing_visible()
    {
        string html = Gs1LinearHtmlReportGenerator.Generate(new Gs1ReportData());

        Assert.Contains("Testing summary of the linear barcode", html);
        Assert.Contains("ISO/IEC 15416 print quality grade", html);
        Assert.Contains("[MISSING: Print method]", html);
        Assert.Contains("Release 26.0, Ratified, Jan 26", html);
    }

    [Fact]
    public void Two_dimensional_generator_keeps_the_canonical_pages_free_of_rfid_content()
    {
        string html = Gs1TwoDimensionalHtmlReportGenerator.Generate(Complete());

        Assert.Contains("Testing summary of the two-dimensional barcode", html);
        Assert.Contains("Overall ISO/IEC 15415 print quality grade", html);
        Assert.DoesNotContain("RFID supplemental information", html);
        Assert.DoesNotContain("EPC Tag URI", html);
        Assert.Contains("@page { size:A4", html);
    }

    [Theory]
    [InlineData(true, "Pass", true)]
    [InlineData(true, "NoTag", true)]
    [InlineData(true, "Skipped", false)]
    [InlineData(false, "Pass", false)]
    [InlineData(null, null, false)]
    public void VeriWedge_addendum_requires_real_rfid_evidence(
        bool? connected, string? status, bool expected)
    {
        var record = new ExcelEngine.Models.VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidReaderConnected = connected,
            RfidStatus = status
        };

        Assert.Equal(expected, VccsPdfRenderer.HasVeriWedgeEvidence(record));
    }

    [Fact]
    public void Letter_profile_is_available_without_changing_the_a4_default()
    {
        string html = Gs1LinearHtmlReportGenerator.Generate(Complete(), Gs1PrintProfile.Letter);

        Assert.Contains("@page { size:Letter", html);
    }

    [Fact]
    public void Gs1_compliance_and_technical_values_are_only_the_literal_plugin_evidence()
    {
        var gs1 = new Dictionary<string, Gs1ParameterAssessment>
        {
            ["Data structure¹ (syntax)"] = new()
            {
                NormativeRequirement = "Required",
                ComplianceResult = "Axicon: NOT COMPLIANT",
                LiteralAssessedValue = "]C10109506000134352",
                CommentReference = "GS1 Table 5-1",
                Provenance = "Axicon export"
            }
        };
        var iso = new Dictionary<string, Gs1IsoAssessment>
        {
            ["Rmin (minimum reflectance)"] = new() { LiteralValue = "6.0%", CommentReference = "ISO/IEC 15416" },
            ["Print growth (+/- %) – Process control parameter"] = new() { LiteralValue = "-2.15%" }
        };
        string html = Gs1LinearHtmlReportGenerator.Generate(Complete(gs1Parameters: gs1, isoParameters: iso));

        Assert.Contains("Axicon: NOT COMPLIANT", html);
        Assert.Contains("]C10109506000134352", html);
        Assert.Contains("6.0%", html);
        Assert.Contains("-2.15%", html);
        Assert.DoesNotContain("<td>Within standard range</td>", html);
        Assert.DoesNotContain("<td>Compliant to standard</td>", html);
    }

    [Fact]
    public void Linear_validation_requires_the_core_linear_evidence()
    {
        Gs1ReportValidationResult result = Gs1LinearHtmlReportGenerator.Validate(Complete());

        Assert.False(result.IsValid);
        Assert.Contains("ISO/IEC evidence: Rmin (minimum reflectance)", result.MissingFields);
        Assert.Contains("GS1 compliance evidence: Data structure¹ (syntax)", result.MissingFields);
        Assert.Contains("GS1 compliance evidence: Validity of GS1 Company Prefix", result.MissingFields);
        Assert.Contains("GS1 compliance evidence: Human readable", result.MissingFields);
    }

    [Fact]
    public void Two_dimensional_validation_requires_core_2d_evidence_but_not_optional_subcomponents()
    {
        Gs1ReportValidationResult result = Gs1TwoDimensionalHtmlReportGenerator.Validate(Complete());

        Assert.False(result.IsValid);
        Assert.Contains("ISO/IEC evidence: Axial nonuniformity", result.MissingFields);
        Assert.Contains("ISO/IEC evidence: Grid nonuniformity", result.MissingFields);
        Assert.Contains("ISO/IEC evidence: Unused Error Correction", result.MissingFields);
        Assert.Contains("ISO/IEC evidence: Fixed pattern damage", result.MissingFields);
        Assert.DoesNotContain("ISO/IEC evidence: • Format information⁴", result.MissingFields);
    }

    private static Gs1ReportData Complete(
        IReadOnlyDictionary<string, Gs1ParameterAssessment>? gs1Parameters = null,
        IReadOnlyDictionary<string, Gs1IsoAssessment>? isoParameters = null) => new()
    {
        IssueDate = "1 January 2026", TestingAgencyName = "Agency",
        ProductDescription = "Product", BarcodeType = "GS1 DataMatrix",
        PrintMethod = "Thermal", NumberOfBarcodesOnProduct = "1",
        VerifierDevice = "Verifier", VerificationSoftwareVersion = "1.0",
        LastVerifierCalibrationDate = "1 January 2026", SymbolSpecificationTable = "Table 5-1",
        TestedEnvironments = "Retail POS", PlacementResult = "Complies",
        TwoDimensionalProximity = "Yes", OverallIsoIecGrade = "4.0 (A)",
        DecodedText = "(01)09506000134352",
        Gs1Parameters = gs1Parameters ?? new Dictionary<string, Gs1ParameterAssessment>(),
        IsoParameters = isoParameters ?? new Dictionary<string, Gs1IsoAssessment>()
    };
}