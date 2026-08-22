using DeviceInterface.Dmst;
using DeviceInterface.Reports;
using ExcelEngine.Models;
using ExcelEngine.Schema;
using ExcelEngine.Writer;
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
            HtmlVerifiedString = "Mon 18-Aug-2026 07:44:37 PM",
            HtmlReportProvenance = HtmlReportProvenance.CorrelatedFilesystem,
            WebscanSourcePath = syntheticPath,
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(realFileName, report, StringComparison.Ordinal);
        Assert.DoesNotContain("_http.html", report, StringComparison.Ordinal);
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
        Assert.Contains(
            "Native TruCheck data and VCCS Digital Link validation are separately labelled",
            report,
            StringComparison.Ordinal);
        Assert.DoesNotContain(">Barcode Image</div>", report, StringComparison.Ordinal);
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
        Assert.Contains("<col style=\"width:18.7%\">", report, StringComparison.Ordinal);
        Assert.Contains("<col style=\"width:13.6%\">", report, StringComparison.Ordinal);
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
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.0",
                Detail = "Parsed GS1 AI data: (01)09506000134352(21)72803288707 Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("DataMan TruCheck Parser", report, StringComparison.Ordinal);
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
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.0",
                Detail = "Parsed GS1 AI data: (01)09506000134352(21)72803288707 Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "Data Format Check (DFC) &#x2014; GS1 Digital Link",
            report,
            StringComparison.Ordinal);
        Assert.Contains("https://id.gs1.org/01/09506000134352/21/72803288707", report,
            StringComparison.Ordinal);
        Assert.Contains("AI (01) GTIN", report, StringComparison.Ordinal);
        Assert.Contains("AI (21) Serial Number", report, StringComparison.Ordinal);
        Assert.Contains(
            "AI (21) Serial Number</td><td class=\"dual-data\">72803288707</td><td class=\"dual-check pass-fg\">PASS</td>",
            report,
            StringComparison.Ordinal);
        Assert.Contains("GS1 Element String", report, StringComparison.Ordinal);
        Assert.Contains("(01)09506000134352<wbr>(21)72803288707<wbr>", report,
            StringComparison.Ordinal);
        Assert.True(
            report.IndexOf(">Web URI<", StringComparison.Ordinal) <
            report.IndexOf("AI (01) GTIN", StringComparison.Ordinal));
        Assert.Contains("DataMan TruCheck Parser", report, StringComparison.Ordinal);
        Assert.Contains("GS1 Barcode Syntax Engine (v. 1.4.0)", report, StringComparison.Ordinal);
        Assert.DoesNotContain("VeriWedge GS1 Parser", report, StringComparison.Ordinal);
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
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.0",
                Detail = "Parsed GS1 AI data: (01)09506000134352(10)LOT-12(17)250101(21)72803288707(20)3 Validated with the official GS1 Barcode Syntax Engine.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains(
            "Data Format Check (DFC) &#x2014; GS1 Element String",
            report,
            StringComparison.Ordinal);
        Assert.Contains("<table class=\"dfc-dual-table\">", report, StringComparison.Ordinal);
        Assert.Contains("DataMan TruCheck Parser", report, StringComparison.Ordinal);
        Assert.Contains("GS1 Barcode Syntax Engine (v. 1.4.0)", report, StringComparison.Ordinal);
        Assert.DoesNotContain("VeriWedge GS1 Parser", report, StringComparison.Ordinal);
        Assert.Contains("AI (10) LOT", report, StringComparison.Ordinal);
        Assert.Contains("AI (17) EXP", report, StringComparison.Ordinal);
        Assert.Contains("AI (21) SN", report, StringComparison.Ordinal);
        Assert.Contains("AI (20) VARIANT", report, StringComparison.Ordinal);
        Assert.Contains("OVERALL: PASS", report, StringComparison.Ordinal);
        Assert.Contains("<td colspan=\"3\" class=\"dual-overall-cell\">", report,
            StringComparison.Ordinal);
        Assert.Contains("AI (10) Batch or Lot Number", report, StringComparison.Ordinal);
        Assert.Contains("AI (17) Expiration Date", report, StringComparison.Ordinal);
        Assert.Contains("AI (20) Variant", report, StringComparison.Ordinal);
        Assert.Contains("parser-element-string-data", report, StringComparison.Ordinal);
        Assert.Contains("table-layout: auto", report, StringComparison.Ordinal);
        Assert.Contains("white-space: nowrap", report, StringComparison.Ordinal);
        Assert.Contains("break-inside: avoid", report, StringComparison.Ordinal);
        Assert.Contains("page-break-inside: avoid", report, StringComparison.Ordinal);
        Assert.DoesNotContain("<th>Source</th>", report, StringComparison.Ordinal);
        Assert.DoesNotContain("VCCS / GS1 Element String syntax validation", report,
            StringComparison.Ordinal);
        Assert.DoesNotContain("GS1 Syntax Engine 1.4.0", report, StringComparison.Ordinal);
        Assert.DoesNotContain("Native TruCheck Data Format Check", report,
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
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.0",
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
    public void Generate_LabelsDigitalLinkNotApplicable()
    {
        var record = new VerificationRecord
        {
            Symbology = "EAN-13",
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.NotApplicable,
                Detail = "Decoded verifier data is not a GS1 Digital Link URI.",
            },
        };

        string report = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("NOT APPLICABLE", report, StringComparison.Ordinal);
        Assert.Contains("Decoded verifier data is not a GS1 Digital Link URI.", report,
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

        Assert.Equal(2, VccsHtmlReportGenerator.MaxRenderedSymbolGroups);
        Assert.Contains("data-vccs-symbol-group=\"true\"", report, StringComparison.Ordinal);
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
            "<span class=\"trucheck-header-title\">TruCheck Barcode Image ",
            report,
            StringComparison.Ordinal);
    }
}