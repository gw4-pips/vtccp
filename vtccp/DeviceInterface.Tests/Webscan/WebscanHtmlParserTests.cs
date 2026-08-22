namespace DeviceInterface.Tests.Webscan;

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
        Assert.Equal("DataMatrix-26-08-22 08_31_50-WEBSCAN 020 CAL.Image1.jpg",
            Path.GetFileName(report.SourceImagePath));

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
    }

    [Fact]
    public async Task FileAdapter_ImportsWithoutChangingRawHtml()
    {
        string sourcePath = GetControlledReportPath();
        string tempDirectory = Path.Combine(Path.GetTempPath(), "vtccp-webscan-" + Guid.NewGuid());
        Directory.CreateDirectory(tempDirectory);
        string copyPath = Path.Combine(tempDirectory, Path.GetFileName(sourcePath));
        File.Copy(sourcePath, copyPath);

        try
        {
            string before = await File.ReadAllTextAsync(copyPath);
            using var adapter = new WebscanHtmlFileAdapter(tempDirectory);
            var records = new List<VerificationRecord>();
            adapter.RecordParsed += (_, record) => records.Add(record);

            VerificationRecord imported = await adapter.ImportFileAsync(copyPath);
            string after = await File.ReadAllTextAsync(copyPath);

            Assert.Equal(before, after);
            Assert.Single(records);
            Assert.Same(imported, records[0]);
            Assert.Equal(Path.GetFileName(copyPath), imported.HtmlSourceFileName);
            Assert.Equal(HtmlReportProvenance.CorrelatedFilesystem, imported.HtmlReportProvenance);
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