using DeviceInterface.Reports;
using ExcelEngine.Models;
using Xunit;

namespace DeviceInterface.Tests.Reports;

public sealed class VccsHtmlReportGeneratorTests
{
    [Fact]
    public void Generate_PrefersRealDmstFilenameOverHttpPlaceholder()
    {
        const string realFileName =
            "_F1_01006961147042882172803282009_2026-08-18_19-44-37-314.html";
        const string syntheticPath =
            @"C:\fake\2026-08-18_19-44-37-000_http.html";

        var record = new VerificationRecord
        {
            Symbology         = "GS1 DataMatrix",
            HtmlSourceFileName = realFileName,
            WebscanSourcePath = syntheticPath,
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(realFileName, report, StringComparison.Ordinal);
        Assert.DoesNotContain("_http.html", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_LabelsHttpOnlySourceAsPlaceholder()
    {
        var record = new VerificationRecord
        {
            Symbology           = "GS1 DataMatrix",
            HtmlSourceProvenance = "HTTP stream placeholder — original DMST filename unavailable",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "[HTTP stream placeholder — original DMST filename unavailable]",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("_vccs_rfid.pdf", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_LabelsLegacyHttpPathAsPlaceholder()
    {
        var record = new VerificationRecord
        {
            Symbology         = "GS1 DataMatrix",
            WebscanSourcePath = @"C:\HTTP_STREAM_PLACEHOLDER\2026-08-18_19-44-37-000_http.html",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "[HTTP STREAM PLACEHOLDER — ORIGINAL DMST FILENAME UNAVAILABLE]",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("_http.html", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_LabelsMissingDmstSourceExplicitly()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("[NO DMST HTML REPORT CORRELATED]", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RendersDataFormatCheckWhenRowsAreAvailable()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            DataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Standard = "GS1 Application Data Format",
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "AI (01) GTIN-14",
                        Data = "0069611470428",
                        Check = "PASS",
                    },
                ],
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("Data Format Check &#x2014; GS1", report, StringComparison.Ordinal);
        Assert.Contains("AI (01) GTIN-14", report, StringComparison.Ordinal);
        Assert.DoesNotContain("DATA FORMAT CHECK UNAVAILABLE", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_LabelsUnavailableDataFormatCheckInsteadOfOmittingSection()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("Data Format Check", report, StringComparison.Ordinal);
        Assert.Contains(
            "[DATA FORMAT CHECK UNAVAILABLE — NO DMST HTML REPORT CORRELATED]",
            report,
            StringComparison.Ordinal);
    }
}