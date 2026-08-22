namespace DeviceInterface.Tests.Webscan;

using DeviceInterface.Reports;
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
            Assert.Contains("Image1 sibling export; not embedded in the HTML", rendered,
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
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            string candidate = Path.Combine(
                directory.FullName,
                "attached_assets",
                "DataMatrix-26-08-22_08_31_50-WEBSCAN_020_CAL._1787402227622.html");
            if (File.Exists(candidate))
                return candidate;
            directory = directory.Parent;
        }

        throw new FileNotFoundException("Controlled TC-829 Webscan report fixture was not found.");
    }
}