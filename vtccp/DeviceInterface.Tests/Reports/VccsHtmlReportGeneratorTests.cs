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
}