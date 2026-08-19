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

        Assert.Contains(
            "TruCheck Barcode Image <span class=\"detail-separator\">|</span> Data Format Check &#x2014; GS1",
            report,
            StringComparison.Ordinal);
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

    [Fact]
    public void Generate_UsesFilenameTimeBeforeHtmlVerifiedTime()
    {
        var record = new VerificationRecord
        {
            Symbology          = "GS1 DataMatrix",
            HtmlSourceFileName = "_F1_01006961147042882172803282009_2026-08-18_20-04-21-314.html",
            HtmlVerifiedString = "Wed 19-Aug-2026 12:04:21 AM",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "<td colspan=\"2\">2026-08-18 20-04-21-314</td>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "<td colspan=\"2\">Wed 19-Aug-2026 12:04:21 AM</td>",
            report,
            StringComparison.Ordinal);
        Assert.Equal(
            "2026-08-18_20-04-21-314",
            VccsHtmlReportGenerator.GetOutputTimestamp(record));
    }

    [Fact]
    public void Generate_UsesRawHtmlVerifiedTimeWhenFilenameHasNoTimestamp()
    {
        const string verified = "Mon 18-Aug-2026 08:04:21(520ms) PM";
        var record = new VerificationRecord
        {
            Symbology          = "GS1 DataMatrix",
            HtmlSourceFileName = "barcode-report.html",
            HtmlVerifiedString = verified,
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains($"<td colspan=\"2\">{verified}</td>", report, StringComparison.Ordinal);
        Assert.Equal("2026-08-18_20-04-21", VccsHtmlReportGenerator.GetOutputTimestamp(record));
    }

    [Fact]
    public void Generate_UsesOneCombinedImageAndDfcSectionWithFullHeightDivider()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            JpegImageBase64 = "AQID",
            DataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Standard = "GS1 Application Data Format",
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "AI (21) Serial",
                        Data = "72803282009",
                        Check = "PASS",
                    },
                ],
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("<table class=\"barcode-detail-grid\"", report, StringComparison.Ordinal);
        Assert.Contains("<td class=\"barcode-image-column\"", report, StringComparison.Ordinal);
        Assert.Contains("<td class=\"barcode-dfc-column\"", report, StringComparison.Ordinal);
        Assert.Contains("border-left: 2px solid #1a3a6b", report, StringComparison.Ordinal);
        Assert.Contains("object-position: left center", report, StringComparison.Ordinal);
        Assert.Contains("See TruCheck report for additional details", report, StringComparison.Ordinal);
        Assert.DoesNotContain("display: grid", report, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("grid-template-columns", report, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(">Barcode Image</div>", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_SerialMismatch_UsesMismatchOnce()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Fail",
            RfidMismatchDetail =
                "Serial:RFID=72803282010,BC=72803282009",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "Fail &#x2014; Serial Number mismatch",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("mismatch mismatch", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_UsesUniformBlueWhiteBarcodeBanners()
    {
        string report = VccsHtmlReportGenerator.Generate(
            new VerificationRecord { Symbology = "GS1 DataMatrix" });

        Assert.Contains(
            ".barcode-sec-hdr {\n    background: #1a3a6b; color: white;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".sec-sub-hdr {\n    background: #1a3a6b; color: white;",
            report,
            StringComparison.Ordinal);
    }
}