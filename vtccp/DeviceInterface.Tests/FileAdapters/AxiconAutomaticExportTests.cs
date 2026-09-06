namespace DeviceInterface.Tests.FileAdapters;

using DeviceInterface.FileAdapters;
using ExcelEngine.Models;
using Xunit;

public sealed class AxiconAutomaticExportTests
{
    [Fact]
    public async Task ImportFileAsync_ArchivesTheExactStableBytes()
    {
        string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        string archive = Path.Combine(root, "archive");
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "scan.csv");
        byte[] expected = await File.ReadAllBytesAsync(Fixture("linear.csv"));
        await File.WriteAllBytesAsync(source, expected);
        try
        {
            using var adapter = new AxiconFileAdapter(root, archive);
            VerificationRecord record = await adapter.ImportFileAsync(source);
            Assert.NotEqual(source, record.SourceArtifactPath);
            Assert.Equal(expected, await File.ReadAllBytesAsync(record.SourceArtifactPath!));
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task Archive_UsesImmutableContentAddressedNameAndEmittedProvenance()
    {
        string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        string archive = Path.Combine(root, "archive");
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "scan.csv");
        try
        {
            using var adapter = new AxiconFileAdapter(root, archive);
            await File.WriteAllBytesAsync(source, await File.ReadAllBytesAsync(Fixture("linear.csv")));
            VerificationRecord first = await adapter.ImportFileAsync(source);
            await File.WriteAllBytesAsync(source, await File.ReadAllBytesAsync(Fixture("gs1-data-matrix.csv")));
            VerificationRecord second = await adapter.ImportFileAsync(source);

            string[] archived = Directory.GetFiles(archive, "*.csv");
            Assert.Equal(2, archived.Length);
            Assert.NotEqual(first.SourceArtifactPath, second.SourceArtifactPath);
            Assert.Contains(first.SourceArtifactPath!, archived);
            Assert.Contains(second.SourceArtifactPath!, archived);
            Assert.NotEqual(source, second.SourceArtifactPath);
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task Watcher_SuppressesLaterEventsWithIdenticalContent()
    {
        string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "scan.csv");
        var received = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        int count = 0;
        using var adapter = new AxiconFileAdapter(root);
        adapter.RecordParsed += (_, _) => { count++; received.TrySetResult(); };
        adapter.Start();
        try
        {
            string csv = await File.ReadAllTextAsync(Fixture("linear.csv"));
            await File.WriteAllTextAsync(source, csv);
            await received.Task.WaitAsync(TimeSpan.FromSeconds(5));
            await File.WriteAllTextAsync(source, csv);
            await Task.Delay(900);
            Assert.Equal(1, count);
        }
        finally { await adapter.StopAsync(); Directory.Delete(root, true); }
    }

    [Fact]
    public async Task Watcher_WaitsForTwoStageWriteBeforeParsing()
    {
        string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "scan.csv");
        string csv = await File.ReadAllTextAsync(Fixture("linear.csv"));
        int split = csv.IndexOf('\n') + 1;
        var received = new TaskCompletionSource<VerificationRecord>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var adapter = new AxiconFileAdapter(root);
        adapter.RecordParsed += (_, record) => received.TrySetResult(record);
        adapter.Start();
        try
        {
            await File.WriteAllTextAsync(source, csv[..split]);
            await Task.Delay(150);
            await File.AppendAllTextAsync(source, csv[split..]);
            VerificationRecord record = await received.Task.WaitAsync(TimeSpan.FromSeconds(6));
            Assert.Equal("012345678905", record.DecodedData);
        }
        finally { await adapter.StopAsync(); Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("linear.csv", SymbologyFamily.Linear1D)]
    [InlineData("gs1-data-matrix.csv", SymbologyFamily.GS1DataMatrix)]
    [InlineData("gs1-qr.csv", SymbologyFamily.GS1QRCode)]
    public void SuppliedTemplateContract_MapsLiteralValues(string fixture, SymbologyFamily family)
    {
        string csv = File.ReadAllText(Fixture(fixture));
        var report = AxiconAutomaticExportParser.Parse(csv, fixture, new DateTime(2026, 3, 19));

        Assert.True(report.ParseSucceeded, report.ParseError);
        VerificationRecord record = report.ToVerificationRecord();
        Assert.Equal(AxiconFileAdapter.Brand, record.VerifierBrand);
        Assert.Equal(family, record.SymbologyFamily);
        Assert.Equal(report.Message, record.DecodedData);
        Assert.Null(record.DeviceName);
        Assert.Null(record.DeviceModel);
        Assert.Equal(report.Get("Reader Serial"), record.DeviceSerial);
        Assert.Equal(report.SoftwareVersion, record.SoftwareVersion);
        Assert.Equal(new DateTime(2026, 3, 1), record.CalibrationDate);
        Assert.Equal(report.Values, record.VerifierSourceFields);
    }

    [Theory]
    [InlineData("malformed.csv")]
    [InlineData("incomplete.csv")]
    public void MalformedOrIncompleteExports_AreExplicitFailures(string fixture)
    {
        var report = AxiconAutomaticExportParser.Parse(
            File.ReadAllText(Fixture(fixture)), fixture, DateTime.UtcNow);
        Assert.False(report.ParseSucceeded);
        Assert.False(string.IsNullOrWhiteSpace(report.ParseError));
        Assert.Throws<InvalidDataException>(() => report.ToVerificationRecord());
    }

    private static string Fixture(string name) => Path.Combine(
        AppContext.BaseDirectory, "Fixtures", "Axicon", name);
}