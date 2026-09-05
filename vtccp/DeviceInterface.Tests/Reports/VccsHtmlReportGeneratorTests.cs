using DeviceInterface.Dmst;
using DeviceInterface.Reports;
using ExcelEngine.Models;
using ExcelEngine.Schema;
using ExcelEngine.Writer;
using System.Text.RegularExpressions;
using Xunit;

namespace DeviceInterface.Tests.Reports;

public sealed class VccsHtmlReportGeneratorTests
{
    [Theory]
    [InlineData(null, "DM475V", "COGNEX DataMan TruCheck Barcode Verification Results Summary", "See associated TruCheck verification report for additional details")]
    [InlineData("WEBSCAN", null, "WEBSCAN TruCheck Barcode Verification Results Summary", "See associated TruCheck verification report for additional details")]
    [InlineData("AXICON", null, "AXICON Barcode Verification Results Summary", "See associated AXICON verification report for additional details")]
    [InlineData("OMRON/LVS", null, "OMRON Barcode Verification Results Summary", "See associated OMRON verification report for additional details")]
    [InlineData("REA", null, "REA Barcode Verification Results Summary", "See associated REA verification report for additional details")]
    [InlineData("UNRECOGNIZED", null, "Barcode Verification Results Summary", "See associated verification report for additional details")]
    public void Generate_SelectsApprovedBarcodeVerificationHeader(
        string? verifierBrand,
        string? deviceModel,
        string expectedTitle,
        string expectedNote)
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            VerifierBrand = verifierBrand,
            DeviceModel = deviceModel,
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verification-report.html",
            HtmlVerifiedString = "Sat 05-Sep-2026 12:00:00 PM",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            $"{expectedTitle}<span class=\"sec-note\"> &#x2014; <em>{expectedNote}</em></span>",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RecognizesCognexFromLiveDeviceNameWhenDeviceModelIsMissing()
    {
        var record = new VerificationRecord
        {
            Symbology = "QR",
            DeviceName = "DM475-866D76",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verification-report.html",
            HtmlVerifiedString = "Sat 05-Sep-2026 04:15:41 PM",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "COGNEX DataMan TruCheck Barcode Verification Results Summary",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "See associated TruCheck verification report for additional details",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_UncorrelatedReportKeepsExplicitUnavailableHeaderWithoutBrand()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            VerifierBrand = "WEBSCAN",
            DeviceModel = "DM475V",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "[BARCODE VERIFICATION UNAVAILABLE — NO CORRELATED DMST HTML]",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "No correlated verification report is available",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "WEBSCAN TruCheck Barcode Verification Results Summary",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "COGNEX DataMan TruCheck Barcode Verification Results Summary",
            report,
            StringComparison.Ordinal);
    }

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
            HtmlVerifiedString = "Mon 18-Aug-2026 07:44:37 PM",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            WebscanSourcePath = syntheticPath,
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(realFileName, report, StringComparison.Ordinal);
        Assert.DoesNotContain("_http.html", report, StringComparison.Ordinal);
        Assert.Contains("App v", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_DoesNotTreatHttpOnlySourceAsAReport()
    {
        var record = new VerificationRecord
        {
            Symbology           = "GS1 DataMatrix",
            HtmlSourceProvenance = "HTTP stream placeholder — original DMST filename unavailable",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("[NO CORRELATED DMST HTML REPORT]", report, StringComparison.Ordinal);
        Assert.DoesNotContain("HTTP stream placeholder", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_DoesNotTreatLegacyHttpPathAsAReport()
    {
        var record = new VerificationRecord
        {
            Symbology         = "GS1 DataMatrix",
            WebscanSourcePath = @"C:\HTTP_STREAM_PLACEHOLDER\2026-08-18_19-44-37-000_http.html",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("[NO CORRELATED DMST HTML REPORT]", report, StringComparison.Ordinal);
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

        Assert.Contains("[NO CORRELATED DMST HTML REPORT]", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RendersDataFormatCheckWhenRowsAreAvailable()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlSourceFileName = "source.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlDataFormatCheck = new DataFormatCheckResult
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
            "TruCheck Barcode Image</span><span class=\"sec-note\"> | <em>Data Format Check (DFC)</em> &#x2014; GS1 Element String",
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
    public void Generate_UsesVerbatimHtmlVerifiedTimeNeverFilenameTime()
    {
        var record = new VerificationRecord
        {
            Symbology          = "GS1 DataMatrix",
            HtmlSourceFileName = "_F1_01006961147042882172803282009_2026-08-18_20-04-21-314.html",
            HtmlVerifiedString = "Wed 19-Aug-2026 12:04:21 AM",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "<td colspan=\"2\">Wed 19-Aug-2026 12:04:21 AM</td>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "<td colspan=\"2\">2026-08-18 20-04-21-314</td>",
            report,
            StringComparison.Ordinal);
        Assert.Equal(
            "2026-08-19_00-04-21",
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
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
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
            HtmlSourceFileName = "source.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlBarcodeImageBase64 = "AQID",
            HtmlDataFormatCheck = new DataFormatCheckResult
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
        Assert.DoesNotContain("Native TruCheck data and VCCS Digital Link validation are separately labelled",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(">Barcode Image</div>", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_SiblingExportImageOmitsProvenanceAndNativeStandardAnnotations()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlSourceFileName = "webscan-export.html",
            HtmlVerifiedString = "Sat 22-Aug-2026 08:31:50 AM",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlBarcodeImageBase64 = "iVBORw0KGgo=",
            HtmlBarcodeImageMimeType = "image/png",
            HtmlBarcodeImageProvenance = "SiblingExport",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Fail,
                Standard = "GS1 Application Data Format",
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "AI (21) Serial",
                        Data = "72803282009",
                        Check = "FAIL",
                    },
                ],
            },
            VeriWedgeValidationUsed = true,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.1",
                Detail = "Parsed GS1 AI data: (21)72803282009",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("data:image/png;base64,iVBORw0KGgo=", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Image1 sibling export; not embedded in the HTML", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("image referenced by the HTML export", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("Native standard: GS1 Application Data Format", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("Standard: GS1 Application Data Format", report,
            StringComparison.Ordinal);
        Assert.Contains("AI (21) Serial", report, StringComparison.Ordinal);
        Assert.Contains("OVERALL: FAIL", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_TagDetectedWithoutReaderLockResult_RendersUnknownLockStatus()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            RfidGcpTableDate = "2026-05-03",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("Yes &#x2014; Unknown", report, StringComparison.Ordinal);
        Assert.Contains("from official GS1 GCP prefix table as of 2026-05-03",
            report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_TagDetectedWithPermanentLock_RendersPermalocked()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            RfidTagLockStatus = "PermaLocked",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("Yes &#x2014; Permalocked", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Yes &#x2014; Unknown", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_NoCorrelatedHtml_NeverLeaksTransportVerifierValues()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            DecodedData = "TRANSPORT_DECODED_DATA",
            Standard = "TRANSPORT_STANDARD",
            FormalGrade = "4/A",
            Aperture = 99,
            Wavelength = 999,
            Lighting = "TRANSPORT_LIGHTING",
            OverallGrade = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS"),
            DataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Standard = "TRANSPORT_DFC_STANDARD",
                Rows = [new DataFormatCheckRow { Name = "TRANSPORT_DFC_ROW", Data = "X", Check = "PASS" }],
            },
            JpegImageBase64 = "TRANSPORT_IMAGE",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("[NO CORRELATED DMST HTML REPORT]", report, StringComparison.Ordinal);
        Assert.Contains("[UNAVAILABLE — NO CORRELATED DMST HTML]", report, StringComparison.Ordinal);
        Assert.DoesNotContain("TRANSPORT_DECODED_DATA", report, StringComparison.Ordinal);
        Assert.DoesNotContain("TRANSPORT_STANDARD", report, StringComparison.Ordinal);
        Assert.DoesNotContain("TRANSPORT_LIGHTING", report, StringComparison.Ordinal);
        Assert.DoesNotContain("TRANSPORT_DFC_STANDARD", report, StringComparison.Ordinal);
        Assert.DoesNotContain("TRANSPORT_DFC_ROW", report, StringComparison.Ordinal);
        Assert.DoesNotContain("TRANSPORT_IMAGE", report, StringComparison.Ordinal);
        Assert.DoesNotContain("4/A", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_CorrelatedHtml_UsesOnlyLiteralHtmlVerifierValues()
    {
        var record = new VerificationRecord
        {
            Symbology = "Transport Symbology",
            DecodedData = "TRANSPORT_DATA",
            Standard = "TRANSPORT_STANDARD",
            FormalGrade = "4/A",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlSymbology = "GS1 DataMatrix",
            HtmlDecodedData = "(01)00696114704283(21)72803282009",
            HtmlStandard = "ISO 15415:2024",
            HtmlOverallGradeDisplay = "4.0 (A)",
            HtmlAperture = "16",
            HtmlWavelength = "660",
            HtmlLighting = "45Q",
            HtmlFormalGrade = "4.0/16/660/45Q",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("verifier-output.html", report, StringComparison.Ordinal);
        Assert.Contains("Mon 18-Aug-2026 08:04:21 PM", report, StringComparison.Ordinal);
        Assert.DoesNotContain("GS1 Application Data Format", report, StringComparison.Ordinal);
        Assert.Contains("4.0/16/660/45Q", report, StringComparison.Ordinal);
        Assert.DoesNotContain("TRANSPORT_DATA", report, StringComparison.Ordinal);
        Assert.DoesNotContain("TRANSPORT_STANDARD", report, StringComparison.Ordinal);
        Assert.DoesNotContain("4/A", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_StandaloneEan13_UsesTheDefaultGs1ParserComparison()
    {
        var record = new VerificationRecord
        {
            Symbology = "EAN-13",
            SymbologyFamily = SymbologyFamily.Linear1D,
            IsStandaloneLinear = true,
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "ean13-verifier-output.html",
            HtmlVerifiedString = "Thu 15-Jan-2026 10:30:45 AM",
            HtmlSymbology = "EAN-13",
            HtmlDecodedData = "5901234123457",
            HtmlStandard = "ISO 15416:2024",
            HtmlOverallGradeDisplay = "4.0 (A)",
            HtmlAperture = "06",
            HtmlWavelength = "660",
            HtmlLighting = "Diffuse",
            HtmlFormalGrade = "A/06/660/Diffuse",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Standard = "GS1 Application Data Format",
                Rows =
                [
                    new() { Name = "GTIN", Data = "590123412345", Check = "PASS" },
                    new() { Name = "Check Digit", Data = "7", Check = "PASS" },
                ],
            },
            VeriWedgeValidationUsed = true,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
                Detail = "Parser state must not create a 2D comparison panel for EAN-13.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Single(
            Regex.Matches(
                report,
                "<tbody>\\s*<tr data-vccs-symbol-group=\"true\">",
                RegexOptions.CultureInvariant).Cast<Match>());
        Assert.Contains("TruCheck Barcode Image</span><span class=\"sec-note\"> | <em>Data Format Check (DFC)</em> &#x2014; GS1 GTIN", report,
            StringComparison.Ordinal);
        Assert.Contains("5901234123457", report, StringComparison.Ordinal);
        Assert.Contains("ISO 15416:2024", report, StringComparison.Ordinal);
        Assert.Contains("4.0 (A)", report, StringComparison.Ordinal);
        Assert.Contains("A/06/660/Diffuse", report, StringComparison.Ordinal);
        Assert.Contains("590123412345", report, StringComparison.Ordinal);
        Assert.Contains("Check Digit", report, StringComparison.Ordinal);
        Assert.DoesNotContain(
            "[DATA FORMAT CHECK UNAVAILABLE — NOT PRESENT IN TRUCHECK HTML]",
            report,
            StringComparison.Ordinal);
        Assert.Contains("DataMan TruCheck GS1 Parser", report, StringComparison.Ordinal);
        Assert.Contains("VeriWedge GS1 Parser", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_ThreeSymbolReportUsesIdenticalNonSplittableDfcShells()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlSymbology = "GS1 DataMatrix",
            HtmlDecodedData = "01095060001343522172803288707",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "three-symbol-report.html",
            MultiSymbolReports =
            [
                new NativeWebscanReportSummary
                {
                    Ordinal = 1, Symbology = "GS1 DataMatrix",
                    DecodedData = "01095060001343522172803288707",
                },
                new NativeWebscanReportSummary
                {
                    Ordinal = 2, Symbology = "EAN-13",
                    DecodedData = "5901234123457",
                },
                new NativeWebscanReportSummary
                {
                    Ordinal = 3, Symbology = "Code 128",
                    DecodedData = "ABC123",
                },
            ],
            VeriWedgeValidationUsed = true,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
            },
        });

        Assert.Equal(3, Regex.Matches(report, "class=\"barcode-detail-section report-block\"",
            RegexOptions.CultureInvariant).Count);
        Assert.Equal(3, Regex.Matches(report,
            "class=\"sec-sub-hdr trucheck-barcode-hdr barcode-detail-header\"",
            RegexOptions.CultureInvariant).Count);
        Assert.DoesNotContain("class=\"barcode-dual-header\"", report, StringComparison.Ordinal);
        Assert.Contains(".barcode-detail-section {\n    break-inside: avoid;",
            report, StringComparison.Ordinal);
        Assert.Contains("<div class=\"fr\"></div>", report, StringComparison.Ordinal);
        Assert.DoesNotContain("counter(page)", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Multi-Symbol Qualification", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Additional native symbols remain listed", report,
            StringComparison.Ordinal);
        Assert.Equal(3, Regex.Matches(report, "<table class=\"dfc-dual-table\">",
            RegexOptions.CultureInvariant).Count);
        Assert.Equal(3, Regex.Matches(
            report,
            "<span class=\"trucheck-header-title\">TruCheck Barcode Image</span><span class=\"sec-note\"> \\| <em>Data Format Check \\(DFC\\)</em> &#x2014; ",
            RegexOptions.CultureInvariant).Count);
        Assert.Contains("GS1 GTIN", report, StringComparison.Ordinal);
        Assert.Contains("GS1 Element String", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_MultiSymbolReportUsesOneNativeOrderAndCompleteGradeWarningPresentation()
    {
        const string gtin = "00696114704288";
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlSymbology = "GS1 DataMatrix",
            HtmlDecodedData = "(01)00696114704288(21)72803282010",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "source-order-report.html",
            HtmlVerifiedString = "Sun 23-Aug-2026 11:24:40 AM",
            RfidReaderConnected = true,
            RfidStatus = "Pass",
            RfidGtin14 = gtin,
            RfidSerial = "72803282010",
            RfidLinearGtin14Matches = true,
            RfidMatchScope = "Both",
            BarcodeSymbolAgreement = "Pass",
            LinearGtin14 = gtin,
            MultiSymbolReports =
            [
                new NativeWebscanReportSummary
                {
                    Ordinal = 2,
                    Symbology = "GS1 DataMatrix",
                    SymbologyFamily = SymbologyFamily.DataMatrix.ToString(),
                    DecodedData = "(01)00696114704288(21)72803282010",
                    Gtin14 = gtin,
                    Standard = "ISO15415:2011",
                    ApertureDisplay = "06",
                    ApertureUnit = "mm",
                },
                new NativeWebscanReportSummary
                {
                    Ordinal = 3,
                    Symbology = "UPCA",
                    SymbologyFamily = SymbologyFamily.Linear1D.ToString(),
                    DecodedData = "696114704288",
                    Gtin14 = gtin,
                    Standard = "ANSI/ISO",
                    ApertureDisplay = "06",
                    ApertureUnit = "mm",
                    Notes = "ISO15416:2016 Warning: Symbol Magnification is less than 80%",
                },
                new NativeWebscanReportSummary
                {
                    Ordinal = 1,
                    Symbology = "GS1 DataMatrix",
                    SymbologyFamily = SymbologyFamily.DataMatrix.ToString(),
                    DecodedData = "(01)00000000000000(21)72803282010",
                    Gtin14 = "00000000000000",
                    Standard = "ISO15415:2011",
                    ApertureDisplay = "06",
                    ApertureUnit = "mm",
                },
            ],
        });

        int summaryFirst = report.IndexOf(">#1 \u2013 GS1 DataMatrix</td>", StringComparison.Ordinal);
        int summarySecond = report.IndexOf(">#2 \u2013 GS1 DataMatrix</td>", StringComparison.Ordinal);
        int summaryThird = report.IndexOf(">#3 \u2013 UPCA</td>", StringComparison.Ordinal);
        Assert.True(summaryFirst >= 0 && summarySecond > summaryFirst && summaryThird > summarySecond);
        Assert.Contains("<th>Aperture (mm)</th>", report, StringComparison.Ordinal);
        Assert.Contains(">ANSI/ISO</td>", report, StringComparison.Ordinal);
        Assert.DoesNotContain(">ANSI/ISO*</td>", report, StringComparison.Ordinal);
        Assert.Contains(">ISO15416:2016*</td>", report, StringComparison.Ordinal);
        Assert.Contains("*Warning: Symbol Magnification is less than 80%", report,
            StringComparison.Ordinal);
        Assert.Equal(3, Regex.Matches(report, "<table class=\"dfc-dual-table\">",
            RegexOptions.CultureInvariant).Count);
        Assert.Contains(
            "#1 \u2013 GS1 DataMatrix RFID Validation Result",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "#2 \u2013 GS1 DataMatrix RFID Validation Result",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "#3 \u2013 UPCA RFID Validation Result",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "Fail &#x2014; EPC GTIN does not match (#1) GTIN",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "Pass &#x2014; EPC GTIN matches (#2) and (#3) GTINs; EPC Serial Number matches (#2) S/N",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "Pass &#x2014; EPC GTIN matches linear (#3) and 2D (#2) GTINs",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "Pass &#x2014; GTIN-14 00696114704288 (#2 &amp; #3)",
            report,
            StringComparison.Ordinal);
        Assert.Equal(3, Regex.Matches(
            report,
            "<span class=\"trucheck-header-title\">TruCheck Barcode Image</span><span class=\"sec-note\"> \\| <em>Data Format Check \\(DFC\\)</em> &#x2014; ",
            RegexOptions.CultureInvariant).Count);
        Assert.DoesNotContain("TruCheck Barcode (#", report, StringComparison.Ordinal);
        Assert.Equal(3, Regex.Matches(
            report,
            "<div class=\"barcode-image-symbol-label\">Symbol #\\d+</div>",
            RegexOptions.CultureInvariant).Count);
        int symbolOne = report.IndexOf(
            "<div class=\"barcode-image-symbol-label\">Symbol #1</div>",
            StringComparison.Ordinal);
        int symbolTwo = report.IndexOf(
            "<div class=\"barcode-image-symbol-label\">Symbol #2</div>",
            StringComparison.Ordinal);
        int symbolThree = report.IndexOf(
            "<div class=\"barcode-image-symbol-label\">Symbol #3</div>",
            StringComparison.Ordinal);
        Assert.True(symbolOne >= 0 && symbolTwo > symbolOne && symbolThree > symbolTwo);
        Assert.DoesNotContain(
            "[DIGITAL LINK URI NOT AVAILABLE]",
            report,
            StringComparison.Ordinal);
        Assert.Contains(".rfid-table .rfid-symbol-result td { font-size: 8pt; }",
            report, StringComparison.Ordinal);
        Assert.Contains(
            ".dfc-dual-table .dual-overall td {\n     border-top: 1px solid #aaa;",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_SummaryRendersTruCheckApplicationSettingsInline()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlSymbology = "GS1 DataMatrix",
            HtmlDecodedData = "(01)00696114704283",
            // A DFC label must never replace the canonical DMCC Application Standard.
            HtmlApplicationStandard = "GS1 Application Data Format",
            ApplicationStandardSetting = "Custom",
            DataFormatCheckSetting = "GS1",
            ApertureSettingMode = "User Set",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "<span class=\"app-settings-label\">Application Std. / Data Format Check / Aperture:</span> <span class=\"app-settings-values\">Custom / GS1 / User Set</span>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "GS1 Application Data Format<span class=\"app-settings-separator\"> / </span>GS1",
            report,
            StringComparison.Ordinal);
        Assert.Contains("table-layout: fixed;", report, StringComparison.Ordinal);
        Assert.Contains("<col style=\"width:18.2%\">", report, StringComparison.Ordinal);
        Assert.Contains("<col style=\"width:14.5%\">", report, StringComparison.Ordinal);
        Assert.Contains(
            "<th class=\"app-settings-hdr\"><span class=\"app-settings-label\">Application Std. / Data Format Check / Aperture:</span> <span class=\"app-settings-values\">Custom / GS1 / User Set</span></th>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".sum-table th:nth-child(2) { border-right: 1.5px solid #999; }",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".sum-table th.app-settings-hdr {\n    border-left: 1.5px solid #999;\n    border-bottom: 1.5px solid #999;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "<td colspan=\"2\" style=\"font-family:Consolas,monospace;\">(01)00696114704283</td>",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_ParsedDfcHeadingDoesNotOverridePerScanApplicationStandard()
    {
        const string html = """
            <html><body>
              <p>Verified: Tue 18-Aug-2026 09:51:39(375ms) PM</p>
              <table>
                <tr><td><strong>Data</strong></td><td>(01)00696114704283</td></tr>
                <tr><td><strong>Symbology</strong></td><td>GS1 DataMatrix</td></tr>
              </table>
              <table>
                <tr><th colspan="3">Data Format Check</th></tr>
                <tr><th colspan="3">GS1 Application Data Format: PASS</th></tr>
                <tr><td>GS1 Header</td><td>(01)</td><td>PASS</td></tr>
              </table>
            </body></html>
            """;

        var parsed = DmstHtmlScraper.ParseHtml(
            html,
            @"C:\fake\2026-08-18_21-51-39-000_fixture.html");
        var record = new VerificationRecord
        {
            Symbology = parsed.HtmlSymbology ?? "Unknown",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = parsed.HtmlVerifiedString,
            HtmlSymbology = parsed.HtmlSymbology,
            HtmlDecodedData = parsed.HtmlDecodedData,
            HtmlApplicationStandard = parsed.HtmlApplicationStandard,
            HtmlDataFormatCheck = parsed.ScrapedDataFormatCheck,
            ApplicationStandardSetting = "Custom",
            DataFormatCheckSetting = "GS1",
            ApertureSettingMode = "User Set",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Null(parsed.HtmlApplicationStandard);
        Assert.Equal("GS1 Application Data Format", parsed.ScrapedDataFormatCheck!.Standard);
        Assert.Contains(
            "<th class=\"app-settings-hdr\"><span class=\"app-settings-label\">Application Std. / Data Format Check / Aperture:</span> <span class=\"app-settings-values\">Custom / GS1 / User Set</span></th>",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_SummaryShowsNoneWhenTruCheckHasNoDataFormatCheck()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlSymbology = "GS1 DataMatrix",
            HtmlDecodedData = "(01)00696114704283",
            ApplicationStandardSetting = "Custom",
            DataFormatCheckSetting = "None",
            ApertureSettingMode = "User Set",
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "<th class=\"app-settings-hdr\"><span class=\"app-settings-label\">Application Std. / Data Format Check / Aperture:</span> <span class=\"app-settings-values\">Custom / None / User Set</span></th>",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_CorrelatedHtml_DoesNotUseLegacyCalculatedDfc()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            DataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Standard = "LEGACY_LOCAL_DFC",
                Rows = [new DataFormatCheckRow
                    { Name = "LEGACY_LOCAL_DFC_ROW", Data = "X", Check = "PASS" }],
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("[DATA FORMAT CHECK UNAVAILABLE — NOT PRESENT IN TRUCHECK HTML]",
            report, StringComparison.Ordinal);
        Assert.DoesNotContain("LEGACY_LOCAL_DFC", report, StringComparison.Ordinal);
        Assert.DoesNotContain("LEGACY_LOCAL_DFC_ROW", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RendersVccsDigitalLinkWithoutSourceColumnOrHeading()
    {
        var record = new VerificationRecord
        {
            Symbology = "QR Code",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlDecodedData = "https://id.gs1.org/01/09506000134352/21/72803288707",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Rows = [new DataFormatCheckRow
                    { Name = "Verifier GS1 row", Data = "vendor data", Check = "PASS" }],
            },
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.1",
                Detail = "Parsed GS1 AI data: (01)09506000134352(21)72803288707 Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("DataMan TruCheck GS1 Parser", report, StringComparison.Ordinal);
        Assert.Contains("Web URI", report, StringComparison.Ordinal);
        Assert.Contains("AI (01) GTIN", report, StringComparison.Ordinal);
        Assert.Contains("Verifier GS1 row", report, StringComparison.Ordinal);
        Assert.DoesNotContain("<th>Source</th>", report, StringComparison.Ordinal);
        Assert.DoesNotContain("VCCS / GS1 Digital Link syntax validation", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("GS1 Syntax Engine 1.4.0", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Native TruCheck Data Format Check", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Native Webscan Digital Link", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Native DataMan Digital Link", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_WebscanUsesSoftwareHeaderAndWebscanParserHeading()
    {
        var record = new VerificationRecord
        {
            VerifierBrand = "WEBSCAN",
            DeviceName = "Webscan TruCheck",
            SoftwareVersion = "3.03.74",
            Symbology = "GS1 DataMatrix",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "webscan-report.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "GS1 Header",
                        Data = "<F1>",
                        Check = "PASS",
                    },
                ],
            },
            VeriWedgeValidationUsed = true,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.1",
                Detail = "Parsed GS1 AI data: (01)00696114704283",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "<div class=\"ln\">Software: 3.03.74</div>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "<div class=\"ln\">Firmware:",
            report,
            StringComparison.Ordinal);
        Assert.Contains("Webscan TruCheck GS1 Parser", report, StringComparison.Ordinal);
        Assert.DoesNotContain("DataMan TruCheck GS1 Parser", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_NoVerifierDfcForDigitalLink_UsesVeriWedgeAlgorithmOnly()
    {
        var record = new VerificationRecord
        {
            Symbology = "QR Code",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlDecodedData = "https://id.gs1.org/01/09506000134352/21/72803288707",
            DataFormatCheckSetting = "None",
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.1",
                Detail = "Parsed GS1 AI data: (01)09506000134352(21)72803288707 Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "TruCheck Barcode Image</span><span class=\"sec-note\"> | <em>Data Format Check (DFC)</em> &#x2014; Digital Link",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("Verifier DFC", report, StringComparison.Ordinal);
        Assert.Contains("https://id.gs1.org/01/09506000134352/21/72803288707", report,
            StringComparison.Ordinal);
        Assert.Contains("AI (01) GTIN", report, StringComparison.Ordinal);
        Assert.Contains("AI (21) Serial Number", report, StringComparison.Ordinal);
        Assert.Contains(
            "AI (21) Serial Number</td><td class=\"dual-data\">72803288707</td><td class=\"dual-check pass-fg\">PASS</td>",
            report,
            StringComparison.Ordinal);
        Assert.Contains("class=\"dual-data parser-uri-data\"",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".dfc-dual-table td:nth-child(6).parser-uri-data",
            report,
            StringComparison.Ordinal);
        Assert.Contains("GS1 Element String", report, StringComparison.Ordinal);
        Assert.Contains("(01)09506000134352<wbr>(21)72803288707<wbr>", report,
            StringComparison.Ordinal);
        Assert.True(
            report.IndexOf(">Web URI<", StringComparison.Ordinal) <
            report.IndexOf("AI (01) GTIN", StringComparison.Ordinal));
        Assert.Contains("DataMan TruCheck GS1 Parser", report, StringComparison.Ordinal);
        Assert.Contains(
            "VeriWedge GS1 Parser — GS1 Barcode Syntax Engine (v. 1.4.1)",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("<th>Source</th>", report, StringComparison.Ordinal);
        Assert.DoesNotContain("VCCS / GS1 Digital Link syntax validation", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("GS1 Syntax Engine 1.4.0", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Native TruCheck Data Format Check", report, StringComparison.Ordinal);
        Assert.DoesNotContain("DATA FORMAT CHECK UNAVAILABLE", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_Gs1ElementString_RendersDynamicPageSafeDualComparison()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlDecodedData = "(01)09506000134352(21)72803288707",
            DataFormatCheckSetting = "GS1",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Rows =
                [
                    new DataFormatCheckRow { Name = "AI (01) GTIN-14", Data = "09506000134352", Check = "PASS" },
                    new DataFormatCheckRow { Name = "AI (10) LOT", Data = "LOT-12", Check = "PASS" },
                    new DataFormatCheckRow { Name = "AI (17) EXP", Data = "250101", Check = "PASS" },
                    new DataFormatCheckRow { Name = "AI (21) SN", Data = "72803288707", Check = "PASS" },
                    new DataFormatCheckRow { Name = "AI (20) VARIANT", Data = "3", Check = "PASS" },
                ],
            },
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.1",
                Detail = "Parsed GS1 AI data: (01)09506000134352(10)LOT-12(17)250101(21)72803288707(20)3 Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "TruCheck Barcode Image</span><span class=\"sec-note\"> | <em>Data Format Check (DFC)</em> &#x2014; GS1 Element String",
            report,
            StringComparison.Ordinal);
        Assert.Contains("<table class=\"dfc-dual-table\">", report, StringComparison.Ordinal);
        Assert.Contains("DataMan TruCheck GS1 Parser", report, StringComparison.Ordinal);
        Assert.Contains(
            "VeriWedge GS1 Parser — GS1 Barcode Syntax Engine (v. 1.4.1)",
            report,
            StringComparison.Ordinal);
        Assert.Contains("AI (10) LOT", report, StringComparison.Ordinal);
        Assert.Contains("AI (17) EXP", report, StringComparison.Ordinal);
        Assert.Contains("AI (21) SN", report, StringComparison.Ordinal);
        Assert.Contains("AI (20) VARIANT", report, StringComparison.Ordinal);
        Assert.Contains("OVERALL: PASS", report, StringComparison.Ordinal);
        Assert.Contains("<td colspan=\"3\" class=\"dual-overall-cell\">", report,
            StringComparison.Ordinal);
        Assert.Contains("AI (10) Batch or Lot Num.", report, StringComparison.Ordinal);
        Assert.Contains("AI (17) Expiration Date", report, StringComparison.Ordinal);
        Assert.Contains("AI (20) Variant", report, StringComparison.Ordinal);
        Assert.Contains("parser-element-string-data", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Verifier DFC", report, StringComparison.Ordinal);
        Assert.Contains("table-layout: fixed", report, StringComparison.Ordinal);
        Assert.Contains("white-space: nowrap", report, StringComparison.Ordinal);
        Assert.Contains("break-inside: avoid", report, StringComparison.Ordinal);
        Assert.Contains("page-break-inside: avoid", report, StringComparison.Ordinal);
        Assert.Contains("padding: 0.25in 0.4in 0.5in 0.4in", report, StringComparison.Ordinal);
        Assert.Contains(".barcode-image-column {\n    width: 19.48%", report, StringComparison.Ordinal);
        Assert.Contains(".barcode-dfc-column {\n    width: 80.52%", report, StringComparison.Ordinal);
        Assert.Contains(".dfc-dual-table col.dual-left-field  { width: 14%; }", report,
            StringComparison.Ordinal);
        Assert.Contains(".dfc-dual-table col.dual-left-data   { width: 21%; }", report,
            StringComparison.Ordinal);
        Assert.Contains(".dfc-dual-table col.dual-left-check  { width: 6%; }", report,
            StringComparison.Ordinal);
        Assert.Contains(".dfc-dual-table col.dual-right-field { width: 19.3%; }", report,
            StringComparison.Ordinal);
        Assert.Contains(".dfc-dual-table col.dual-right-data  { width: 30.7%; }", report,
            StringComparison.Ordinal);
        Assert.Contains(".dfc-dual-table .dual-subhead th {\n    background: #dbe5f1; font-size: 6.5pt;",
            report, StringComparison.Ordinal);
        Assert.DoesNotContain("<th>Source</th>", report, StringComparison.Ordinal);
        Assert.DoesNotContain("VCCS / GS1 Element String syntax validation", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("GS1 Syntax Engine 1.4.0", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Native TruCheck Data Format Check", report,
            StringComparison.Ordinal);
        Assert.Contains("min-height: 11in", report, StringComparison.Ordinal);
        Assert.Contains(
            "position: fixed; left: 0.4in; right: 0.4in; bottom: 0.18in;",
            report, StringComparison.Ordinal);
        Assert.DoesNotContain("counter(page)", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Verification Command &amp; Control System (VCCS)",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_MultiElementGs1DataMatrixKeepsParserFieldsAndValuesUnwrapped()
    {
        const string elementString =
            "(01)00696114704283(17)260822(10)BATCH-2026-08-22-LONG-LOT-VALUE(21)SERIAL-72803282010";
        var record = new VerificationRecord
        {
            VerifierBrand = "WEBSCAN",
            DeviceName = "Webscan TruCheck",
            SoftwareVersion = "3.03.74",
            Symbology = "GS1 DataMatrix",
            HtmlSymbology = "GS1 DataMatrix",
            HtmlDecodedData = elementString,
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "multi-element-gs1.html",
            HtmlVerifiedString = "Sat 22-Aug-2026 01:24:26 PM",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Standard = "GS1 Application Data Format",
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "AI (01) GTIN-14",
                        Data = "00696114704283",
                        Check = "PASS",
                    },
                    new DataFormatCheckRow
                    {
                        Name = "AI (17) Expiration Date",
                        Data = "260822",
                        Check = "PASS",
                    },
                    new DataFormatCheckRow
                    {
                        Name = "AI (10) Batch or Lot Number",
                        Data = "BATCH-2026-08-22-LONG-LOT-VALUE",
                        Check = "PASS",
                    },
                    new DataFormatCheckRow
                    {
                        Name = "AI (21) Serial Number",
                        Data = "SERIAL-72803282010",
                        Check = "PASS",
                    },
                ],
            },
            DataFormatCheckSetting = "GS1",
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.1",
                Detail = $"Parsed GS1 AI data: {elementString} Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(elementString, report, StringComparison.Ordinal);
        Assert.Contains("(01)00696114704283<wbr>(17)260822<wbr>(10)BATCH-2026-08-22-LONG-LOT-VALUE<wbr>(21)SERIAL-72803282010<wbr>",
            report,
            StringComparison.Ordinal);
        Assert.Contains("AI (10) Batch or Lot Number", report, StringComparison.Ordinal);
        Assert.Contains(
            "<td class=\"dual-right-field\">AI (10) Batch or Lot Num.</td>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "<td class=\"dual-right-field\">AI (10) Batch or Lot Number</td>",
            report,
            StringComparison.Ordinal);
        Assert.Contains("AI (21) Serial Number", report, StringComparison.Ordinal);
        Assert.Contains(".sum-table td[colspan=\"2\"] {\n    white-space: normal;\n    overflow-wrap: anywhere;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(".dfc-dual-table td:nth-child(1),\n   .dfc-dual-table td:nth-child(5) {\n     min-width: 0;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".dfc-dual-table td:nth-child(6).parser-element-string-data,\n   .dfc-dual-table td:nth-child(6).parser-uri-data {\n     font-size: 6.5pt; line-height: 1.05;\n     white-space: normal; overflow: visible;\n     overflow-wrap: anywhere; word-wrap: break-word; word-break: break-word;",
            report,
            StringComparison.Ordinal);
        Assert.Contains("Webscan TruCheck GS1 Parser", report, StringComparison.Ordinal);
        Assert.DoesNotContain("DataMan TruCheck GS1 Parser", report, StringComparison.Ordinal);
        Assert.Contains("Software: 3.03.74", report, StringComparison.Ordinal);
        Assert.Contains("OVERALL: PASS", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_NoTagUsesHumanReadableValidationResult()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology          = "GS1 DataMatrix",
            RfidReaderConnected = true,
            RfidStatus         = "NoTag",
        });

        Assert.Contains(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Validation Result</td>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "<td>No Tag Detected</td>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "&#x26a0; NO RFID TAG DETECTED",
            report,
            StringComparison.Ordinal);
        Assert.Contains("<table class=\"rfid-table\">", report, StringComparison.Ordinal);
        Assert.Contains("<tr><th class=\"lbl-col\">Field</th><th>Value</th></tr>", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("<td>NoTag</td>", report, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("Skipped")]
    public void Generate_InactiveRfidReader_RendersHeaderOnlyAndNoReservedTable(string? status)
    {
        const string epcSentinel = "RFID_VALUE_MUST_NOT_RENDER";
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology          = "GS1 DataMatrix",
            RfidReaderConnected = false,
            RfidStatus         = status,
            RfidEpcHex = epcSentinel,
            RfidEpcTagUri = "urn:epc:tag:sgtin-96:1.0612345.012345.1",
            RfidGtin14 = epcSentinel,
            RfidSerial = epcSentinel,
            RfidTid = epcSentinel,
        });

        Assert.Contains(
            "VCCS <em>RFID VeriWedge&#x2122; PowerPro</em> EPC RFID Validation Summary (No data: reader not activated).",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("<table class=\"rfid-table\">", report, StringComparison.Ordinal);
        Assert.DoesNotContain("<tr><th class=\"lbl-col\">Field</th><th>Value</th></tr>", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(epcSentinel, report, StringComparison.Ordinal);
        Assert.Contains("<div class=\"report-section report-section-rfid\">", report,
            StringComparison.Ordinal);
        Assert.Contains("Barcode Verification Capture Unavailable", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_ConnectedReaderWithNoData_RendersExpandedRfidValidation()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology          = "GS1 DataMatrix",
            RfidReaderConnected = true,
            RfidStatus         = "NoTag",
        });

        Assert.Contains("<table class=\"rfid-table\">", report, StringComparison.Ordinal);
        Assert.Contains(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Validation Result</td>",
            report,
            StringComparison.Ordinal);
        Assert.Contains("<td>No Tag Detected</td>", report, StringComparison.Ordinal);
        Assert.DoesNotContain("No data: reader not activated", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_DisconnectedReaderOverridesNoDataStatusAndStaysCompact()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology          = "GS1 DataMatrix",
            RfidReaderConnected = false,
            RfidStatus         = "NoTag",
        });

        Assert.DoesNotContain("<table class=\"rfid-table\">", report, StringComparison.Ordinal);
        Assert.Contains("No data: reader not activated", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RfidSectionsUseNormalFlowAndDefaultPageIntactRules()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = null,
        });

        Assert.Contains("width: 8.5in; background: white;", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("\n    height: 11in;", report, StringComparison.Ordinal);
        Assert.Contains(
            ".report-section:not(.report-section--splittable),\n  .report-block:not(.report-block--splittable) {\n    break-inside: avoid;\n    page-break-inside: avoid;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(".report-section--splittable", report, StringComparison.Ordinal);
        Assert.Contains("<div class=\"report-section report-section-summary\">", report,
            StringComparison.Ordinal);
        Assert.Contains("<div class=\"report-section report-section-rfid\">", report,
            StringComparison.Ordinal);
        Assert.Contains("<div class=\"barcode-detail-section report-block\">", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("<tbody>\n          \n        </tbody>", report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_DualParserDividerMovesWithinExistingGapAndBatchLabelCanWrap()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlDecodedData = "(01)00850010224367(17)271004(10)CU26E02",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "AI (10) Batch or Lot Number",
                        Data = "CU26E02",
                        Check = "PASS",
                    },
                ],
            },
            DataFormatCheckSetting = "GS1",
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
                Detail = "Parsed GS1 AI data: (01)00850010224367(17)271004(10)CU26E02",
            },
            VeriWedgeValidationUsed = true,
        });

        Assert.Contains("left: calc(41.56% - 1px);", report, StringComparison.Ordinal);
        Assert.Contains(".barcode-dfc-column::after", report, StringComparison.Ordinal);
        Assert.Contains("content: none;", report, StringComparison.Ordinal);
        Assert.DoesNotContain("border-left: 2px solid #1a3a6b;\n     padding: 0 !important;",
            report,
            StringComparison.Ordinal);
        Assert.Contains("<td class=\"dual-right-field\">AI (10) Batch or Lot Num.</td>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".dfc-dual-table td:nth-child(1),\n   .dfc-dual-table td:nth-child(5) {\n     min-width: 0;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".dfc-dual-table td:nth-child(6).parser-element-string-data,\n   .dfc-dual-table td:nth-child(6).parser-uri-data {\n     font-size: 6.5pt; line-height: 1.05;\n     white-space: normal; overflow: visible;\n     overflow-wrap: anywhere; word-wrap: break-word; word-break: break-word;",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_TruCheckPassUsesRegularRfidValidationLabel()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            VeriWedgeValidationUsed = false,
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Pass,
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "AI (01) GTIN-14",
                        Data = "09506000134352",
                        Check = "PASS",
                    },
                ],
            },
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
            },
        });

        Assert.Contains(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Validation Result</td>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Cross-Validation Result</td>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("<table class=\"dfc-dual-table\">", report,
            StringComparison.Ordinal);
        Assert.Contains("Native TruCheck Data Format Check", report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_KnownUnsupportedDigitalLinkFirmware_StarsNativeFailure()
    {
        var record = new VerificationRecord
        {
            Symbology = "QR Code",
            FirmwareVersion = "6.1.16_sr4",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "verifier-output.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlDecodedData = "https://id.gs1.org/01/09506000134352/21/72803288707",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Fail,
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "GS1 Format",
                        Data = "<F1> Required at beginning of data",
                        Check = "FAIL",
                    },
                ],
            },
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.1",
                Detail = "Parsed GS1 AI data: (01)09506000134352(21)72803288707 Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(">FAIL*</td>", report, StringComparison.Ordinal);
        Assert.Contains("OVERALL: FAIL*</span>", report, StringComparison.Ordinal);
        Assert.Contains(
            "Firmware 6.1.16_sr4 does not support GS1 Digital Link parsing.",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "<div class=\"dual-native-note\">Firmware 6.1.16_sr4 does not support GS1 Digital Link parsing.</div>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "<span class=\"overall-pill pill-native-limitation\">OVERALL: FAIL*</span>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".pill-native-limitation { background:#f8d7da; color:#000; border-color:#000; }",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".dfc-dual-table .dual-native-note {\n      margin-top: 2pt; color: #000;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".header-copyright {\n    font-size: 7pt; color: #000;",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("Native GS1 Digital Link Compatibility", report, StringComparison.Ordinal);
        Assert.DoesNotContain("compat-unsupported", report, StringComparison.Ordinal);
        Assert.DoesNotContain("native parser compatibility limitation", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_KnownUnsupportedDigitalLinkWebscanSoftware_ShowsConciseNativeNote()
    {
        var record = new VerificationRecord
        {
            VerifierBrand = "WEBSCAN",
            SoftwareVersion = "3.03.74",
            Symbology = "QR Code",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = "webscan-report.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlDecodedData = "https://id.gs1.org/01/09506000134352/21/72803288707",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Fail,
                Rows =
                [
                    new DataFormatCheckRow
                    {
                        Name = "GS1 Format",
                        Data = "Digital Link not recognized",
                        Check = "FAIL",
                    },
                ],
            },
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Detail = "Parsed GS1 AI data: (01)09506000134352(21)72803288707 Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "Software 3.03.74 does not support GS1 Digital Link parsing.",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "<div class=\"dual-native-note\">Software 3.03.74 does not support GS1 Digital Link parsing.</div>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(">FAIL*</td>", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RfidGcpLength_ShowsEqualsPrefix()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            RfidEpcTagUri = "urn:epc:tag:sgtin-96:1.0612345.012345.1",
            RfidGcpValid = true,
            RfidGcpLength = 7,
        });

        Assert.Contains("Valid (=7)", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Valid (7)", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RfidValidationResultUsesOneLineLabelAndAlignedDualParserColumns()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            ApplicationPass = "Pass",
            RfidStatus = "Fail",
            HtmlDataFormatCheck = new DataFormatCheckResult
            {
                Overall = OverallPassFail.Fail,
            },
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
            },
            VeriWedgeValidationUsed = true,
        });

        Assert.Contains(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Cross-Validation Result</td>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".rfid-table .rfid-result-label {\n    white-space: normal; overflow-wrap: anywhere;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "class=\"dual-right-parser-header\"",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "class=\"dual-right-field\">Field</th>",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".dfc-dual-table .dual-right-parser-header,\n   .dfc-dual-table .dual-right-field",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_TruCheckPassWithVeriWedgePanel_UsesValidationLabel()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            TruCheckValidationUsable = true,
            TruCheckValidationFailed = false,
            VeriWedgeValidationUsed = true,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
            },
        });

        Assert.Contains(
            "GS1 DataMatrix RFID Validation Result",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "GS1 DataMatrix RFID Cross-Validation Result",
            report,
            StringComparison.Ordinal);
        Assert.Contains("<table class=\"dfc-dual-table\">", report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_TruCheckOnlyPass_DoesNotUseCrossValidationLabel()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            TruCheckValidationUsable = true,
            TruCheckValidationFailed = false,
            VeriWedgeValidationUsed = false,
        });

        Assert.Contains(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Validation Result</td>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Cross-Validation Result</td>",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_TruCheckOnlyFailure_DoesNotUseCrossValidationLabel()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Fail",
            TruCheckValidationUsable = true,
            TruCheckValidationFailed = true,
            VeriWedgeValidationUsed = false,
        });

        Assert.Contains(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Validation Result</td>",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "<td class=\"rfid-result-label\">GS1 DataMatrix RFID Cross-Validation Result</td>",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_UnavailableTruCheckWithVeriWedge_UsesCrossValidationLabel()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            TruCheckValidationUsable = false,
            TruCheckValidationFailed = false,
            VeriWedgeValidationUsed = true,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Detail = "Validated by VeriWedge.",
            },
        });

        Assert.Contains("GS1 DataMatrix RFID Cross-Validation Result", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_FailedTruCheckWithVeriWedge_UsesCrossValidationLabel()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Fail",
            TruCheckValidationUsable = true,
            TruCheckValidationFailed = true,
            VeriWedgeValidationUsed = true,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Invalid,
                Detail = "VeriWedge rejected the GS1 data.",
            },
        });

        Assert.Contains("GS1 DataMatrix RFID Cross-Validation Result", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RfidGcpNotFound_IsNotReportedAsInvalid()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            RfidGcpStatus = "NotFound",
            RfidGcpLength = 7,
        });

        Assert.Contains("NOT FOUND (=7)", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Invalid (=7)", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_RfidGcpInvalid_ShowsLengthWithoutEqualsSign()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            RfidStatus = "Pass",
            RfidGcpStatus = "Invalid",
            RfidGcpLength = 8,
            RfidGcpRegisteredLength = 7,
        });

        Assert.Contains("Invalid (8); Valid = 7", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Invalid (=8)", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_DoesNotRenderVeriWedgePanelWhenItWasNotUsed()
    {
        var record = new VerificationRecord
        {
            Symbology = "EAN-13",
            VeriWedgeValidationUsed = false,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.NotApplicable,
                Detail = "Decoded verifier data is not a GS1 Digital Link URI.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.DoesNotContain("NOT APPLICABLE", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Decoded verifier data is not a GS1 Digital Link URI.", report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_MultiModeRendersOnlyTheCurrentTwoSymbolGroups()
    {
        var record = new VerificationRecord
        {
            Symbology = "QR Code",
            LinearSymbology = "EAN-13",
            LinearDecodedData = "09506000134352",
            HtmlSymbology = "QR Code",
            HtmlDecodedData = "https://id.gs1.org/01/09506000134352",
            HtmlSourceFileName = "multimode.html",
            HtmlVerifiedString = "Mon 18-Aug-2026 08:04:21 PM",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("data-vccs-symbol-group=\"true\"", report, StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_CompositeUsesItsCanonicalLegacySymbolProjection()
    {
        var record = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            LinearSymbology = "UPCA",
            LinearDecodedData = "696114704288",
            HtmlSymbology = "GS1 DataMatrix",
            HtmlDecodedData = "]d20106961147042882",
            IsWebscanComposite = true,
            CompositeOverallStatus = "Pass",
            HtmlSourceFileName = "composite.html",
            HtmlVerifiedString = "Sun 23-Aug-2026 06:27:00 AM",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        int twoDIndex = report.IndexOf(
            ">#2 \u2013 GS1 DataMatrix</td>",
            StringComparison.Ordinal);
        int linearIndex = report.IndexOf(
            ">#1 \u2013 UPCA</td>",
            StringComparison.Ordinal);

        Assert.True(linearIndex >= 0);
        Assert.True(twoDIndex > linearIndex);
    }

    [Fact]
    public void ExcelMappers_WriteExactLocalHtmlBasenameOnlyForCorrelatedReports()
    {
        const string fileName = "_F1_01006961147042882172803282009_2026-08-18_20-04-21-314.html";
        ColumnSchema schema = TruCheckCompatibleSchema.Build();

        var correlated = new VerificationRecord
        {
            Symbology = "GS1 DataMatrix",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            HtmlSourceFileName = fileName,
        };
        var uncorrelated = correlated with
        {
            HtmlReportProvenance = HtmlReportProvenance.HttpStreamOnly,
        };

        Assert.Equal(fileName, DataMatrix2DMapper.Map(correlated, schema)["HtmlSourceFileName"]);
        Assert.Equal(fileName, ISO15416Mapper.Map(correlated, schema)["HtmlSourceFileName"]);
        Assert.Null(DataMatrix2DMapper.Map(uncorrelated, schema)["HtmlSourceFileName"]);
        Assert.Null(ISO15416Mapper.Map(uncorrelated, schema)["HtmlSourceFileName"]);
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
    public void Generate_RfidFailure_UsesExplicitCrossValidationLabel()
    {
        string report = VccsHtmlReportGenerator.Generate(new VerificationRecord
        {
            Symbology = "QR Code",
            RfidStatus = "Fail",
            VeriWedgeValidationUsed = true,
            RfidMismatchDetail = "GTIN14:RFID=09506000134351,BC=09506000134352",
        });

        Assert.Contains(
            "QR Code RFID Cross-Validation Result",
            report,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Generate_UsesMediumBlueForAllTruCheckHeaders()
    {
        string report = VccsHtmlReportGenerator.Generate(
            new VerificationRecord
            {
                Symbology = "GS1 DataMatrix",
                HtmlSourceFileName = "source.html",
                HtmlVerifiedString = "Tue 18-Aug-2026 11:45:27 PM",
                HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            });

        Assert.Contains(
            ".barcode-sec-hdr {\n    background: #2c5296; color: white;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            ".trucheck-barcode-hdr {\n    background: #2c5296; color: white;\n    height: 17pt;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "<span class=\"trucheck-header-title\">TruCheck Barcode Verification Grades</span><span class=\"sec-note\"> &#x2014;",
            report,
            StringComparison.Ordinal);
        Assert.Contains(
            "<span class=\"trucheck-header-title\">TruCheck Barcode Image</span><span class=\"sec-note\"> | <em>Data Format Check (DFC)</em> &#x2014;",
            report,
            StringComparison.Ordinal);
    }
}