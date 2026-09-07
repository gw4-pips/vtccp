using System.Text;
using DeviceInterface.Rfid.FlexWedge;
using DeviceInterface.Rfid.Gcp;
using Xunit;

namespace DeviceInterface.Tests.Rfid.FlexWedge;

public sealed class FlexWedgeCoreTests
{
    private const string Sgtin96 = "30342A7CC844C7D0F36A0676";

    [Theory]
    [InlineData("30342A7CC844C7D0F36A0676", "00696114704318", "72803288694")]
    [InlineData("30342A7CC844C710F36A0650", "00696114704288", "72803288656")]
    [InlineData("30342A7CC844C750F36A066F", "00696114704295", "72803288687")]
    public void Builder_MatchesOwnedFlexWedgeLiveVectors(string epc, string gtin14, string serial)
    {
        var result = new FlexWedgeReadResultBuilder().Build(epc);

        Assert.Equal(gtin14, result.Gtin14);
        Assert.Equal(serial, result.Serial);
    }

    [Fact]
    public void Builder_DecodesSgtin96WithGranularBitsAndAudit()
    {
        var audit = new FlexWedgeReaderAudit(DateTimeOffset.Parse("2026-01-02T03:04:05Z"), "reader", "USB1", -43, 0x3000, "E200", "Locked");
        var result = new FlexWedgeReadResultBuilder().Build(Sgtin96, audit);

        Assert.Equal("00696114704318", result.Gtin14);
        Assert.Equal("72803288694", result.Serial);
        Assert.Equal("urn:epc:id:sgtin:0696114.70431.72803288694", result.EpcUri);
        Assert.Equal("00110000", result.HeaderBits!.Bits);
        Assert.Equal(14, result.GcpBits!.StartBit);
        Assert.Equal(24, result.GcpBits.Length);
        Assert.Equal(audit, result.Audit);
        Assert.Equal(FlexWedgeGcpStatus.NotChecked, result.GcpStatus);
    }

    [Fact]
    public void Builder_DecodesSgtin198AsciiSerialAnd140BitRange()
    {
        // SGTIN-198: filter 1, partition 5, GCP 0696114, item reference 704318, serial ABC.
        var result = new FlexWedgeReadResultBuilder().Build("36342A7CCAAFCFA0C286000000000000000000000000000000");

        Assert.Equal(DeviceInterface.Rfid.Models.EpcScheme.Sgtin198, result.Scheme);
        Assert.Equal("ABC", result.Serial);
        Assert.Equal(140, result.SerialBits!.Length);
        Assert.Equal(58, result.SerialBits.StartBit);
    }

    [Theory]
    [InlineData("<entry prefix=\"0696114\" gcpLength=\"7\" />", FlexWedgeGcpStatus.Valid)]
    [InlineData("<entry prefix=\"0696114\" gcpLength=\"6\" />", FlexWedgeGcpStatus.Mismatch)]
    [InlineData("<entry prefix=\"1234567\" gcpLength=\"7\" />", FlexWedgeGcpStatus.NotFound)]
    public void Builder_MapsGcpStatuses(string entry, FlexWedgeGcpStatus expected)
    {
        var builder = new FlexWedgeReadResultBuilder(new GcpValidator(Table(entry)));
        Assert.Equal(expected, builder.Build(Sgtin96).GcpStatus);
    }

    [Fact]
    public void Builder_ReportsInvalidForMalformedEpcWhenValidatorWasRequested()
    {
        var result = new FlexWedgeReadResultBuilder(
            new GcpValidator(Table("<entry prefix=\"0696114\" gcpLength=\"7\" />"))).Build("ZZ");
        Assert.Equal(FlexWedgeGcpStatus.Invalid, result.GcpStatus);
        Assert.NotEmpty(result.DecodeWarnings);
    }

    [Fact]
    public void OutputAndCsv_EscapeAndDeliverFormattedText()
    {
        var result = new FlexWedgeReadResultBuilder().Build(Sgtin96);
        var profile = FlexWedgeOutputProfile.GtinAndSerial with { Prefix = "[", Suffix = "]", Delimiter = "|", AppendKey = FlexWedgeAppendKey.Tab };
        var sink = new RecordingSink();
        new FlexWedgeTextSinkAdapter(sink).Deliver(result, profile);

        Assert.Equal("[00696114704318|72803288694]\t", sink.Text);
        Assert.Equal("\"a,b\"", FlexWedgeCsvFormatter.Escape("a,b"));
        Assert.Equal("\"a\"\"b\"", FlexWedgeCsvFormatter.Escape("a\"b"));
    }

    [Fact]
    public void Session_ExportsJsonLinesAndCsv()
    {
        var log = new FlexWedgeSessionLogger();
        log.Append(new FlexWedgeReadResultBuilder().Build(Sgtin96));
        using var csv = new MemoryStream();
        using var json = new MemoryStream();
        log.ExportCsv(csv);
        log.ExportJsonLines(json);
        Assert.Contains("RawEpc", Encoding.UTF8.GetString(csv.ToArray()));
        Assert.Contains("HeaderBits", Encoding.UTF8.GetString(csv.ToArray()));
        Assert.Contains(Sgtin96, Encoding.UTF8.GetString(json.ToArray()));
    }

    [Fact]
    public void Session_RecoversValidJsonLinesAndSkipsInterruptedTail()
    {
        var original = new FlexWedgeReadResultBuilder().Build(Sgtin96);
        using var journal = new MemoryStream();
        FlexWedgeSessionLogger.AppendJsonLine(journal, original);
        journal.Write(Encoding.UTF8.GetBytes("{\"incomplete\":"));
        journal.Position = 0;

        var recovered = new FlexWedgeSessionLogger();
        Assert.Equal(1, recovered.ImportJsonLines(journal));
        Assert.Equal(Sgtin96, recovered.Entries[0].RawEpcHex);
    }

    private static GcpLengthTable Table(string entry)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes($"<GCPPrefixFormatList>{entry}</GCPPrefixFormatList>"));
        return GcpLengthTable.LoadFromStream(stream);
    }

    private sealed class RecordingSink : IFlexWedgeTextSink
    {
        public string? Text { get; private set; }
        public void WriteText(string text) => Text = text;
    }
}