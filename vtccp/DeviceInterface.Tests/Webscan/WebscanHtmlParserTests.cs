namespace DeviceInterface.Tests.Webscan;

using DeviceInterface.Reports;
using DeviceInterface.Rfid;
using DeviceInterface.Validation;
using DeviceInterface.Webscan;
using ExcelEngine.Models;
using Xunit;

public sealed class WebscanHtmlParserTests
{
    [Fact]
    public void ControlledTc829Html_MapsOnlyLiteralWebscanValues()
    {
        string sourcePath = GetControlledReportPath();

        WebscanHtmlReport report = WebscanHtmlParser.ParseFile(sourcePath);
        VerificationRecord record = report.ToVerificationRecord();

        Assert.True(report.ParseSucceeded, report.ParseError);
        Assert.Equal("Sat 22-Aug-2026 08:31:50 AM", report.VerifiedDisplay);
        Assert.Equal(new DateTime(2026, 8, 22, 8, 31, 50), report.VerifiedDateTime);
        Assert.Equal("3.03.74", report.SoftwareVersion);
        Assert.Equal("TC-829-0213-021", report.DeviceSerial);
        Assert.Equal("WEBSCAN 020 CAL.", report.Data);
        Assert.Equal("DataMatrix", report.Symbology);
        Assert.Equal("GW4", report.VerifiedBy);
        Assert.Equal("VeriWedge Development", report.JobNumber);
        Assert.Equal("ISO15415:2011", report.Standard);
        Assert.Equal("A (4.0)", report.OverallGradeDisplay);
        Assert.Equal("4.0/08/660/45", report.FormalGrade);
        Assert.Equal("White on Black", report.ImagePolarity);
        Assert.Equal(19, report.QualityParameters.Count);
        Assert.Equal("DataMatrix-26-08-22_08_31_50-WEBSCAN_020_CAL.Image1_1787402227622.jpg",
            Path.GetFileName(report.SourceImagePath));
        Assert.Equal(WebscanImageProvenance.SiblingExport, report.SourceImageProvenance);
        Assert.Equal("image/jpeg", report.SourceImageMimeType);
        Assert.False(string.IsNullOrWhiteSpace(report.SourceImageBase64));
        Assert.Null(report.DataFormatCheck);

        Assert.Equal("WEBSCAN", record.VerifierBrand);
        Assert.Equal("USB", record.ConnectionMedium);
        Assert.Equal(SymbologyFamily.DataMatrix, record.SymbologyFamily);
        Assert.Equal(ImagePolarity.WhiteOnBlack, record.ImagePolarity);
        Assert.Equal(4.0m, record.OverallGrade?.NumericGrade);
        Assert.Equal(GradeLetterValue.A, record.OverallGrade?.LetterGrade);
        Assert.Equal(OverallPassFail.NotApplicable, record.OverallGrade?.PassFail);
        Assert.Equal(100m, record.UEC_Percent);
        Assert.Equal(79m, record.SC_Percent);
        Assert.Equal("Rl/Rd (83/4)", record.SC_RlRd);
        Assert.Equal(0m, record.ANU_Percent);
        Assert.Equal(1m, record.GNU_Percent);
        Assert.Equal(4.0m, record.AG_Value);
        Assert.Equal("20.0 mil", report.NominalXDim);
        Assert.Equal(20.0m, record.NominalXDim_2D);
        Assert.Null(record.ApplicationStandard);
        Assert.Null(record.FirmwareVersion);
        Assert.Equal(sourcePath, record.WebscanSourcePath);
        Assert.Equal(report.SourceImageBase64, record.HtmlBarcodeImageBase64);
        Assert.Equal("SiblingExport", record.HtmlBarcodeImageProvenance);
        Assert.Equal("image/jpeg", record.HtmlBarcodeImageMimeType);
    }

    [Fact]
    public void ControlledTc829UpcaHtml_ImportsNativeLinearReport()
    {
        string sourcePath = GetUpcaReportPath();

        WebscanHtmlReport report = WebscanHtmlParser.ParseFile(sourcePath);
        VerificationRecord record = report.ToVerificationRecord();

        Assert.True(report.ParseSucceeded, report.ParseError);
        Assert.Equal("Sat 22-Aug-2026 08:47:49 PM", report.VerifiedDisplay);
        Assert.Equal("696114704318", report.Data);
        Assert.Equal("UPCA", report.Symbology);
        Assert.Equal("ANSI/ISO", report.Standard);
        Assert.Equal("A (3.5)", report.OverallGradeDisplay);
        Assert.Equal("06", report.ApertureDisplay);
        Assert.Equal("660", report.WavelengthDisplay);
        Assert.Null(report.Lighting);
        Assert.Equal("3.5/06/660", report.FormalGrade);
        Assert.Equal(9, report.QualityParameters.Count);
        Assert.Contains(report.QualityParameters,
            parameter => parameter.Name == "Symbol Contrast (SC)");
        Assert.Contains(report.QualityParameters,
            parameter => parameter.Name == "Decodability (DEC)");
        Assert.NotNull(report.DataFormatCheck);
        Assert.Equal(OverallPassFail.Pass, report.DataFormatCheck!.Overall);
        Assert.Equal("GS1 Application Data Format", report.DataFormatCheck.Standard);
        Assert.Equal(WebscanImageProvenance.SiblingExport, report.SourceImageProvenance);
        Assert.Equal(
            "UPCA-26-08-22_20_47_49-696114704318Image1_1787446139035.jpg",
            Path.GetFileName(report.SourceImagePath));

        Assert.Equal(SymbologyFamily.Linear1D, record.SymbologyFamily);
        Assert.Equal("UPCA", record.Symbology);
        Assert.Equal(3.5m, record.OverallGrade?.NumericGrade);
        Assert.Equal(GradeLetterValue.A, record.OverallGrade?.LetterGrade);
        Assert.Same(report.DataFormatCheck, record.DataFormatCheck);
        Assert.Equal("SiblingExport", record.HtmlBarcodeImageProvenance);
    }

    [Fact]
    public void ConcatenatedLinearAndTwoDReports_ImportAsOneDualSymbologyReportByFamily()
    {
        string linearPath = GetUpcaReportPath();
        string twoDPath = GetQrExportWithoutAverageGradePath();
        string raw = File.ReadAllText(linearPath) + File.ReadAllText(twoDPath);

        WebscanHtmlMultiSymbolReport composite = WebscanHtmlParser.ParseMultiSymbol(raw, twoDPath);

        Assert.True(composite.ParseSucceeded, composite.ParseError);
        Assert.Equal(SymbologyFamily.Linear1D,
            WebscanHtmlParser.MapSymbologyFamily(composite.LinearReport!.Symbology!));
        Assert.Equal(SymbologyFamily.QRCode,
            WebscanHtmlParser.MapSymbologyFamily(composite.TwoDReport!.Symbology!));

        VerificationRecord record = composite.ToVerificationRecord();
        Assert.True(record.IsWebscanComposite);
        Assert.Equal("UPCA", record.LinearSymbology);
        Assert.Equal("QR", record.Symbology);
        Assert.Equal("696114704318", record.LinearDecodedData);
        Assert.NotEmpty(record.LinearQualityParameters);
        Assert.NotNull(record.LinearDataFormatCheck);
        Assert.NotNull(record.HtmlDataFormatCheck);
        Assert.NotEqual(record.LinearJpegImageBase64, record.HtmlBarcodeImageBase64);
        Assert.True(record.LinearTwoDMatch);
    }

    [Fact]
    public void ControlledTc829Ean8Html_PreservesLinearNotesAndUsesNotesColumn()
    {
        string sourcePath = GetAttachedReportPath(
            "EAN8-26-08-22_21_24_41-00671583_1787448450248.html");

        WebscanHtmlReport report = WebscanHtmlParser.ParseFile(sourcePath);
        VerificationRecord record = report.ToVerificationRecord();
        string html = VccsHtmlReportGenerator.Generate(record);

        Assert.True(report.ParseSucceeded, report.ParseError);
        Assert.Equal("ANSI/ISO", report.Standard);
        Assert.Equal("0.1/06/660", report.FormalGrade);
        Assert.Equal("ISO15416:2016", report.Notes);
        Assert.Null(report.Lighting);
        Assert.Equal("ISO15416:2016", record.HtmlNotes);
        Assert.Contains("<th>Notes</th>", html, StringComparison.Ordinal);
        Assert.Contains("ISO15416:2016", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<th>Lighting</th>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void TwoSymbolWebscanHtml_ProducesOneDualSymbologyRecordWithSeparateNativeReports()
    {
        string sourcePath = GetAttachedReportPath(
            "Webscan_Report--26-08-22_21_53_29_1787450204733.html");

        string rawHtml = File.ReadAllText(sourcePath);
        WebscanHtmlMultiSymbolReport composite =
            WebscanHtmlParser.ParseMultiSymbol(rawHtml, sourcePath);

        Assert.True(composite.ParseSucceeded, composite.ParseError);
        Assert.Equal("UPCA", composite.LinearReport?.Symbology);
        Assert.Equal("696114704288", composite.LinearReport?.Data);
        Assert.Equal("GS1 DataMatrix", composite.TwoDReport?.Symbology);
        Assert.StartsWith("<F1>0100696114704288",
            composite.TwoDReport?.Data ?? string.Empty);
        Assert.EndsWith("Image1_1787450204733.jpg",
            composite.LinearReport?.SourceImagePath ?? string.Empty);
        Assert.EndsWith("Image2_1787450204733.jpg",
            composite.TwoDReport?.SourceImagePath ?? string.Empty);
        Assert.Equal(9, composite.LinearReport?.QualityParameters.Count);
        Assert.Equal(19, composite.TwoDReport?.QualityParameters.Count);

        VerificationRecord record = composite.ToVerificationRecord();
        Assert.Equal("UPCA", record.LinearSymbology);
        Assert.Equal("696114704288", record.LinearDecodedData);
        Assert.Equal("GS1 DataMatrix", record.Symbology);
        Assert.Equal("00696114704288", RfidValidator.ExtractAi01(record.DecodedData));
        Assert.Equal(
            composite.LinearReport?.DataFormatCheck,
            record.LinearDataFormatCheck);
        Assert.Equal(
            composite.LinearReport?.Standard,
            record.HtmlLinearStandard);
        Assert.Equal(
            composite.LinearReport?.OverallGradeDisplay,
            record.HtmlLinearGradeDisplay);
        Assert.Equal(
            composite.LinearReport?.Notes,
            record.HtmlLinearLighting);
        Assert.NotEqual(
            composite.LinearReport?.SourceImageBase64,
            composite.TwoDReport?.SourceImageBase64);

        string reportHtml = VccsHtmlReportGenerator.Generate(record with
        {
            BarcodeSymbolAgreement = "Pass",
            BarcodeSymbolAgreementDetail = "GTIN-14: 00696114704288",
            LinearGtin14 = "00696114704288",
            RfidStatus = "Pass",
            RfidLinearGtin14Matches = true,
            RfidMatchScope = "Both",
            CompositeOverallStatus = "Pass",
        });
        Assert.Contains("MULTI-SYMBOL VERIFIED", reportHtml);
        Assert.Contains("Barcode Symbol Agreement", reportHtml);
        Assert.Contains("EPC GTIN matches both barcode symbols", reportHtml);
    }

    [Fact]
    public void ThreeSymbolWebscanHtml_PreservesEveryNativeReportAndDoesNotInventQualification()
    {
        string sourcePath = GetThreeSymbolReportPath();
        string raw = File.ReadAllText(sourcePath);

        WebscanHtmlMultiSymbolReport report =
            WebscanHtmlParser.ParseMultiSymbol(raw, sourcePath);

        Assert.True(report.ParseSucceeded, report.ParseError);
        Assert.Equal(3, report.SymbolReports.Count);
        Assert.Equal(
            ["UPCA", "GS1 DataMatrix", "UPCA"],
            report.SymbolReports.Select(symbol => symbol.Symbology));
        Assert.Equal(
            ["696114704288", "<F1>0100696114704288" +
             "2172803282010", "696114704288"],
            report.SymbolReports.Select(symbol => symbol.Data));
        Assert.Equal(
            ["ANSI/ISO", "ISO15415:2011", "ANSI/ISO"],
            report.SymbolReports.Select(symbol => symbol.Standard));
        Assert.Equal(["A (3.6)", "A (4.0)", "A (3.6)"],
            report.SymbolReports.Select(symbol => symbol.OverallGradeDisplay));
        Assert.All(report.SymbolReports, symbol =>
        {
            Assert.Equal("Sat 22-Aug-2026 09:53:29 PM", symbol.VerifiedDisplay);
            Assert.Equal("3.03.74", symbol.SoftwareVersion);
            Assert.Equal("TC-829-0213-021", symbol.DeviceSerial);
            Assert.Equal("GW4", symbol.VerifiedBy);
        });
        Assert.Equal([9, 19, 9],
            report.SymbolReports.Select(symbol => symbol.QualityParameters.Count));
        Assert.All(report.SymbolReports, symbol =>
        {
            Assert.True(symbol.ParseSucceeded, symbol.ParseError);
            Assert.Equal(OverallPassFail.Pass, symbol.DataFormatCheck?.Overall);
            Assert.Equal("GS1 Application Data Format", symbol.DataFormatCheck?.Standard);
            Assert.NotEmpty(symbol.DataFormatCheck?.Rows ?? []);
            Assert.Equal(WebscanImageProvenance.SiblingExport, symbol.SourceImageProvenance);
            Assert.Equal("image/jpeg", symbol.SourceImageMimeType);
            Assert.False(string.IsNullOrWhiteSpace(symbol.SourceImageBase64));
        });
        Assert.EndsWith(".Image1.jpg", report.SymbolReports[0].SourceImagePath);
        Assert.EndsWith(".Image2.jpg", report.SymbolReports[1].SourceImagePath);
        Assert.EndsWith(".Image3.jpg", report.SymbolReports[2].SourceImagePath);
        Assert.NotEqual(
            report.SymbolReports[0].SourceImageBase64,
            report.SymbolReports[1].SourceImageBase64);

        MultiSymbolQualification qualification = report.Qualify("00696114704288");
        Assert.Equal(MultiSymbolQualificationStatus.Qualified, qualification.Status);
        Assert.Equal([1, 2, 3], qualification.MatchingSymbols);
        Assert.Empty(qualification.MismatchingSymbols);
        Assert.Contains("all recognized symbol identities agree with RFID EPC",
            qualification.Reasons);

        MultiSymbolQualification rejected = report.Qualify("00000000000000");
        Assert.Equal(MultiSymbolQualificationStatus.Rejected, rejected.Status);
        Assert.Equal([], rejected.MatchingSymbols);
        Assert.Equal([1, 2, 3], rejected.MismatchingSymbols);
        Assert.Contains(rejected.Reasons, reason =>
            reason == "RFID GTIN mismatch: RFID 00000000000000, symbols 00696114704288");

        VerificationRecord record = report.ToVerificationRecord();
        Assert.Equal(3, record.MultiSymbolReports.Count);
        Assert.Equal([1, 2, 3],
            record.MultiSymbolReports.Select(summary => summary.Ordinal));
        Assert.Equal(["UPCA", "GS1 DataMatrix", "UPCA"],
            record.MultiSymbolReports.Select(summary => summary.Symbology));
        Assert.Equal([9, 19, 9],
            record.MultiSymbolReports.Select(summary => summary.QualityParameters.Count));
        Assert.All(record.MultiSymbolReports, summary =>
        {
            Assert.Equal(OverallPassFail.Pass, summary.DataFormatCheck?.Overall);
            Assert.NotEmpty(summary.DataFormatCheck?.Rows ?? []);
            Assert.Equal("SiblingExport", summary.SourceImageProvenance);
            Assert.False(string.IsNullOrWhiteSpace(summary.SourceImagePath));
        });
    }

    [Fact]
    public void ControlledTc829UpcaHtml_RendersDualGs1ParserAndMatchedRfidResult()
    {
        string sourcePath = GetUpcaReportPath();
        WebscanHtmlReport parsed = WebscanHtmlParser.ParseFile(sourcePath);
        VerificationRecord record = parsed.ToVerificationRecord() with
        {
            RfidStatus = "Pass",
            RfidGtin14 = "00696114704318",
            RfidSerial = "72803288694",
            RfidTagLockStatus = "PermaLocked",
            TruCheckValidationUsable = true,
            TruCheckValidationFailed = false,
            VeriWedgeValidationUsed = true,
            VccsDigitalLinkValidation = new DigitalLinkValidationResult
            {
                Status = DigitalLinkValidationStatus.Valid,
                Source = DigitalLinkValidationResult.VccsElementStringSource,
                EngineVersion = "GS1 Barcode Syntax Engine 1.4.1",
                Detail = "Parsed GS1 AI data: (01)00696114704318",
            },
        };

        string html = VccsHtmlReportGenerator.Generate(record);

        Assert.Contains("&#x2713; RFID MATCHED", html, StringComparison.Ordinal);
        Assert.Contains("UPCA RFID Validation Result", html, StringComparison.Ordinal);
        Assert.Contains("Pass &#x2014; EPC data matches barcode GTIN", html,
            StringComparison.Ordinal);
        Assert.Contains("Webscan TruCheck GS1 Parser", html, StringComparison.Ordinal);
        Assert.Contains("VeriWedge GS1 Parser", html, StringComparison.Ordinal);
        Assert.Contains("(01)00696114704318", html, StringComparison.Ordinal);
        Assert.Contains("OVERALL: PASS", html, StringComparison.Ordinal);
        Assert.DoesNotContain("RFID MISMATCH", html, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FileAdapter_ImportsWithoutChangingSourceArtifacts()
    {
        string sourcePath = GetControlledReportPath();
        string tempDirectory = Path.Combine(Path.GetTempPath(), "vtccp-webscan-" + Guid.NewGuid());
        Directory.CreateDirectory(tempDirectory);
        string copyPath = Path.Combine(tempDirectory, Path.GetFileName(sourcePath));
        File.Copy(sourcePath, copyPath);
        WebscanHtmlReport sourceReport = WebscanHtmlParser.ParseFile(sourcePath);
        Assert.NotNull(sourceReport.SourceImagePath);
        string copyImagePath = Path.Combine(
            tempDirectory,
            Path.GetFileName(sourceReport.SourceImagePath));
        File.Copy(sourceReport.SourceImagePath!, copyImagePath);
        string sourcePdfPath = Path.ChangeExtension(sourcePath, ".pdf");
        string copyPdfPath = Path.Combine(tempDirectory, Path.GetFileName(sourcePdfPath));
        File.Copy(sourcePdfPath, copyPdfPath);

        try
        {
            byte[] htmlBefore = await File.ReadAllBytesAsync(copyPath);
            byte[] imageBefore = await File.ReadAllBytesAsync(copyImagePath);
            byte[] pdfBefore = await File.ReadAllBytesAsync(copyPdfPath);
            using var adapter = new WebscanHtmlFileAdapter(tempDirectory);
            var records = new List<VerificationRecord>();
            adapter.RecordParsed += (_, record) => records.Add(record);

            VerificationRecord imported = await adapter.ImportFileAsync(copyPath);
            byte[] htmlAfter = await File.ReadAllBytesAsync(copyPath);
            byte[] imageAfter = await File.ReadAllBytesAsync(copyImagePath);
            byte[] pdfAfter = await File.ReadAllBytesAsync(copyPdfPath);

            Assert.Equal(htmlBefore, htmlAfter);
            Assert.Equal(imageBefore, imageAfter);
            Assert.Equal(pdfBefore, pdfAfter);
            Assert.Single(records);
            Assert.Same(imported, records[0]);
            Assert.Equal(Path.GetFileName(copyPath), imported.HtmlSourceFileName);
            Assert.Equal(HtmlReportProvenance.CorrelatedFilesystem, imported.HtmlReportProvenance);
            Assert.Equal("SiblingExport", imported.HtmlBarcodeImageProvenance);
        }
        finally
        {
            Directory.Delete(tempDirectory, recursive: true);
        }
    }

    [Fact]
    public void NativeDfcAndSiblingImage_AreMappedAndRenderedWithoutMutatingEvidence()
    {
        string fixturePath = GetControlledReportPath();
        string tempDirectory = Path.Combine(Path.GetTempPath(), "vtccp-webscan-" + Guid.NewGuid());
        Directory.CreateDirectory(tempDirectory);
        string sourcePath = Path.Combine(
            tempDirectory,
            "DataMatrix-26-08-22_08_31_50-WEBSCAN_020_CAL._123456789.html");
        string siblingImagePath = Path.Combine(
            tempDirectory,
            "DataMatrix-26-08-22_08_31_50-WEBSCAN_020_CAL.Image1_123456789.png");

        try
        {
            string rawHtml = File.ReadAllText(fixturePath);
            rawHtml = rawHtml.Replace(
                "<img src=\"DataMatrix-26-08-22 08_31_50-WEBSCAN 020 CAL.Image1.jpg\" alt=\"Symbol Image\" style=\"width:auto;max-height:4in;max-width:100%;\" />",
                string.Empty,
                StringComparison.Ordinal);
            rawHtml = rawHtml.Replace(
                "</body>",
                """
                    <table>
                      <tr><th colspan="3">Data Format Check</th></tr>
                      <tr><th colspan="3">GS1 Application Data Format: FAIL</th></tr>
                      <tr><td>Name</td><td>Data</td><td>Check</td></tr>
                      <tr><td>AI (01) GTIN-14</td><td>00696114704283</td><td>PASS</td></tr>
                      <tr><td>AI (21) Serial</td><td>72803282009</td><td>FAIL</td></tr>
                    </table>
                    </body>
                    """,
                StringComparison.Ordinal);
            File.WriteAllText(sourcePath, rawHtml);
            File.WriteAllBytes(siblingImagePath,
            [
                0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
                0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
            ]);

            byte[] htmlBefore = File.ReadAllBytes(sourcePath);
            byte[] imageBefore = File.ReadAllBytes(siblingImagePath);
            WebscanHtmlReport parsed = WebscanHtmlParser.ParseFile(sourcePath);
            VerificationRecord record = parsed.ToVerificationRecord();
            string rendered = VccsHtmlReportGenerator.Generate(record);

            Assert.True(parsed.ParseSucceeded, parsed.ParseError);
            Assert.Equal(siblingImagePath, parsed.SourceImagePath);
            Assert.Equal(WebscanImageProvenance.SiblingExport, parsed.SourceImageProvenance);
            Assert.Equal("image/png", parsed.SourceImageMimeType);
            Assert.NotNull(parsed.DataFormatCheck);
            Assert.Equal(OverallPassFail.Fail, parsed.DataFormatCheck!.Overall);
            Assert.Equal("GS1 Application Data Format", parsed.DataFormatCheck.Standard);
            Assert.Collection(
                parsed.DataFormatCheck.Rows,
                row =>
                {
                    Assert.Equal("AI (01) GTIN-14", row.Name);
                    Assert.Equal("00696114704283", row.Data);
                    Assert.Equal("PASS", row.Check);
                },
                row =>
                {
                    Assert.Equal("AI (21) Serial", row.Name);
                    Assert.Equal("72803282009", row.Data);
                    Assert.Equal("FAIL", row.Check);
                });
            Assert.Same(parsed.DataFormatCheck, record.DataFormatCheck);
            Assert.Same(parsed.DataFormatCheck, record.HtmlDataFormatCheck);
            Assert.Equal("SiblingExport", record.HtmlBarcodeImageProvenance);
            Assert.Contains("data:image/png;base64,", rendered, StringComparison.Ordinal);
            Assert.DoesNotContain("Image1 sibling export; not embedded in the HTML", rendered,
                StringComparison.Ordinal);
            Assert.DoesNotContain("image referenced by the HTML export", rendered,
                StringComparison.Ordinal);
            Assert.Contains("AI (01) GTIN-14", rendered, StringComparison.Ordinal);
            Assert.Contains("AI (21) Serial", rendered, StringComparison.Ordinal);
            Assert.Contains("OVERALL: FAIL", rendered, StringComparison.Ordinal);

            Assert.Equal(htmlBefore, File.ReadAllBytes(sourcePath));
            Assert.Equal(imageBefore, File.ReadAllBytes(siblingImagePath));
        }
        finally
        {
            Directory.Delete(tempDirectory, recursive: true);
        }
    }

    [Fact]
    public void NativeDfcWithoutExplicitOutcome_RemainsUnavailable()
    {
        string rawHtml = File.ReadAllText(GetControlledReportPath()).Replace(
            "</body>",
            """
                <table>
                  <tr><th colspan="3">Data Format Check</th></tr>
                  <tr><td>Name</td><td>Data</td><td>Check</td></tr>
                  <tr><td>AI (10): FAIL</td><td>00696114704283</td><td>FAIL</td></tr>
                </table>
                </body>
                """,
            StringComparison.Ordinal);

        WebscanHtmlReport parsed = WebscanHtmlParser.Parse(rawHtml, @"C:\fixture\report.html");

        Assert.True(parsed.ParseSucceeded, parsed.ParseError);
        Assert.NotNull(parsed.DataFormatCheck);
        Assert.Equal(OverallPassFail.NotApplicable, parsed.DataFormatCheck!.Overall);
        Assert.Single(parsed.DataFormatCheck.Rows);
        Assert.Equal("AI (10): FAIL", parsed.DataFormatCheck.Rows[0].Name);
        Assert.Equal("FAIL", parsed.DataFormatCheck.Rows[0].Check);
    }

    [Fact]
    public void QrExportWithoutAverageGrade_ImportsLiteralOverallGradeAndNativeDfcFailure()
    {
        string sourcePath = GetQrExportWithoutAverageGradePath();

        WebscanHtmlReport report = WebscanHtmlParser.ParseFile(sourcePath);
        VerificationRecord record = report.ToVerificationRecord();

        Assert.True(report.ParseSucceeded, report.ParseError);
        Assert.Equal(SymbologyFamily.QRCode, record.SymbologyFamily);
        Assert.Equal(4.0m, record.OverallGrade?.NumericGrade);
        Assert.Equal(GradeLetterValue.A, record.OverallGrade?.LetterGrade);
        Assert.DoesNotContain(
            report.QualityParameters,
            parameter => parameter.Name.Equals(
                "Average Grade (AG)",
                StringComparison.OrdinalIgnoreCase));
        Assert.Null(record.AG_Value);
        Assert.Null(record.AG_Grade);
        Assert.NotNull(record.DataFormatCheck);
        Assert.Equal(OverallPassFail.Fail, record.DataFormatCheck!.Overall);
        Assert.Equal("GS1 Application Data Format", record.DataFormatCheck.Standard);
    }

    [Fact]
    public void ImageReferenceOutsideReportDirectory_IsRejected()
    {
        string tempDirectory = Path.Combine(Path.GetTempPath(), "vtccp-webscan-" + Guid.NewGuid());
        string reportDirectory = Path.Combine(tempDirectory, "reports");
        Directory.CreateDirectory(reportDirectory);
        string sourcePath = Path.Combine(reportDirectory, "report.html");
        string outsideImagePath = Path.Combine(tempDirectory, "outside.png");

        try
        {
            string rawHtml = File.ReadAllText(GetControlledReportPath()).Replace(
                "DataMatrix-26-08-22 08_31_50-WEBSCAN 020 CAL.Image1.jpg",
                "../outside.png",
                StringComparison.Ordinal);
            File.WriteAllText(sourcePath, rawHtml);
            File.WriteAllBytes(outsideImagePath,
            [
                0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
                0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
            ]);

            WebscanHtmlReport parsed = WebscanHtmlParser.ParseFile(sourcePath);

            Assert.True(parsed.ParseSucceeded, parsed.ParseError);
            Assert.Null(parsed.SourceImagePath);
            Assert.Equal(WebscanImageProvenance.None, parsed.SourceImageProvenance);
            Assert.Null(parsed.SourceImageBase64);
        }
        finally
        {
            Directory.Delete(tempDirectory, recursive: true);
        }
    }

    [Fact]
    public async Task ShutdownDrain_WaitsForDelayedAcceptanceBeforeWorkbookClose()
    {
        var tracker = new WebscanAcceptanceTracker();
        tracker.BeginSession(1);

        bool sessionOpen = true;
        bool recordWritten = false;
        var acceptanceStarted = new TaskCompletionSource(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseAcceptance = new TaskCompletionSource(
            TaskCreationOptions.RunContinuationsAsynchronously);

        Assert.True(tracker.TryAdmit(
            1,
            () => sessionOpen,
            async () =>
            {
                acceptanceStarted.SetResult();
                await releaseAcceptance.Task;

                Assert.True(sessionOpen, "The workbook closed before acceptance finished.");
                recordWritten = true;
            }));

        await acceptanceStarted.Task;
        Task[] admitted = tracker.InvalidateAndCapture();

        // A callback that loses the admission race must not start another write.
        Assert.False(tracker.TryAdmit(1, () => sessionOpen, () => Task.CompletedTask));

        Task drain = Task.WhenAll(admitted);
        Assert.False(drain.IsCompleted);

        releaseAcceptance.SetResult();
        await drain;
        sessionOpen = false;

        Assert.True(recordWritten);
        Assert.False(tracker.TryAdmit(2, () => sessionOpen, () => Task.CompletedTask));
    }

    [Fact]
    public void TitleOnlyHtml_IsRejectedWithoutFabricatingARecordTimestamp()
    {
        const string rawHtml =
            "<html><head><title>Webscan TruCheck™ USB Verification Report</title></head>" +
            "<body><h1>Webscan TruCheck™ USB Verification Report</h1></body></html>";

        WebscanHtmlReport report = WebscanHtmlParser.Parse(rawHtml, @"C:\fake\partial.html");

        Assert.False(report.ParseSucceeded);
        Assert.Contains("verified timestamp", report.ParseError);
        Assert.Throws<InvalidOperationException>(() => report.ToVerificationRecord());
    }

    private static string GetControlledReportPath()
        => GetAttachedAssetPath(
            "DataMatrix-26-08-22_08_31_50-WEBSCAN_020_CAL._1787402227622.html");

    private static string GetQrExportWithoutAverageGradePath()
        => GetAttachedAssetPath(
            "QR-26-08-22_18_58_42-https___srk.my2dir.com_01_00696114704318_1787439623070.html");

    private static string GetUpcaReportPath()
        => GetAttachedAssetPath(
            "UPCA-26-08-22_20_47_49-696114704318_1787446139035.html");

    private static string GetThreeSymbolReportPath()
        => GetAttachedAssetPath(
            "Webscan_Report--26-08-22_21_53_29_three-symbol_1787450204733.html");

    private static string GetAttachedReportPath(string fileName)
        => GetAttachedAssetPath(fileName);

    private static string GetAttachedAssetPath(string fileName)
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            string candidate = Path.Combine(
                directory.FullName,
                "attached_assets",
                fileName);
            if (File.Exists(candidate))
                return candidate;
            directory = directory.Parent;
        }

        throw new FileNotFoundException($"Webscan report fixture '{fileName}' was not found.");
    }
}