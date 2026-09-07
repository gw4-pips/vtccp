using DeviceInterface.Rfid.Gcp;
using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid.FlexWedge;

/// <summary>Builds canonical FlexWedge records from EPC reader observations.</summary>
public sealed class FlexWedgeReadResultBuilder
{
    private readonly GcpValidator? _gcpValidator;

    public FlexWedgeReadResultBuilder(GcpValidator? gcpValidator = null) => _gcpValidator = gcpValidator;

    public FlexWedgeReadResult Build(EpcReadResult read, string? reader = null, string? port = null) =>
        BuildCore(EpcParser.Parse(read.EpcBytes), Convert.ToHexString(read.EpcBytes), null,
            new FlexWedgeReaderAudit(read.ReadTime, reader, port, read.Rssi, read.PcWord, read.Tid, read.LockStatus));

    public FlexWedgeReadResult Build(string? rawEpc, FlexWedgeReaderAudit? audit = null)
    {
        ParsedEpc parsed = EpcParser.ParseHex(rawEpc);
        return BuildCore(parsed, parsed.EpcHex, rawEpc,
            audit ?? new FlexWedgeReaderAudit(DateTimeOffset.UtcNow, null, null, null, null, null, null));
    }

    private FlexWedgeReadResult BuildCore(ParsedEpc parsed, string rawHex, string? rawSource, FlexWedgeReaderAudit audit)
    {
        var warnings = parsed.ParseWarning is null ? Array.Empty<string>() : [parsed.ParseWarning];
        (FlexWedgeGcpStatus status, int? registeredLength) = GetGcpStatus(parsed);
        int? m = parsed.Partition is int partition && partition is >= 0 and <= 6
            ? new[] { 40, 37, 34, 30, 27, 24, 20 }[partition] : null;
        int? n = parsed.Partition is int p && p is >= 0 and <= 6
            ? new[] { 4, 7, 10, 14, 17, 20, 24 }[p] : null;

        return new FlexWedgeReadResult
        {
            RawSource = rawSource,
            RawEpcHex = rawHex,
            Scheme = parsed.Scheme,
            Filter = parsed.Filter,
            Partition = parsed.Partition,
            CompanyPrefix = parsed.CompanyPrefix,
            ItemReference = parsed.ItemReference,
            Gtin14 = parsed.Gtin14,
            Serial = parsed.Serial,
            EpcUri = ToEpcUri(parsed),
            HeaderBits = Field(parsed.EpcBytes, "Header", 0, 8),
            FilterBits = parsed.Filter is null ? null : Field(parsed.EpcBytes, "Filter", 8, 3),
            PartitionBits = parsed.Partition is null ? null : Field(parsed.EpcBytes, "Partition", 11, 3),
            GcpBits = m is int gcpLength ? Field(parsed.EpcBytes, "Gcp", 14, gcpLength) : null,
            ItemReferenceBits = m is int gcpBits && n is int itemLength ? Field(parsed.EpcBytes, "ItemReference", 14 + gcpBits, itemLength) : null,
            SerialBits = m is int gm && n is int ni && parsed.Scheme is EpcScheme.Sgtin96 or EpcScheme.Sgtin198
                ? Field(parsed.EpcBytes, "Serial", 14 + gm + ni, parsed.Scheme == EpcScheme.Sgtin96 ? 38 : 140) : null,
            Audit = audit,
            DecodeWarnings = warnings,
            GcpStatus = status,
            GcpRegisteredLength = registeredLength,
        };
    }

    private (FlexWedgeGcpStatus, int?) GetGcpStatus(ParsedEpc parsed)
    {
        if (_gcpValidator is null) return (FlexWedgeGcpStatus.NotChecked, null);
        if (parsed.CompanyPrefix is null || parsed.Partition is null ||
            parsed.Scheme is not (EpcScheme.Sgtin96 or EpcScheme.Sgtin198))
            return (FlexWedgeGcpStatus.Invalid, null);
        if (!_gcpValidator.TryGetRegisteredLength(parsed, out int length))
            return (FlexWedgeGcpStatus.NotFound, null);
        int? encoded = GcpValidator.GetEncodedGcpLength(parsed);
        return encoded == length ? (FlexWedgeGcpStatus.Valid, length) : (FlexWedgeGcpStatus.Mismatch, length);
    }

    private static EpcBitField? Field(byte[] bytes, string name, int start, int length)
    {
        if (bytes.Length * 8 < start + length) return null;
        var bits = new char[length];
        for (int i = 0; i < length; i++)
            bits[i] = ((bytes[(start + i) / 8] >> (7 - ((start + i) % 8))) & 1) == 1 ? '1' : '0';
        return new EpcBitField(name, start, length, new string(bits));
    }

    private static string? ToEpcUri(ParsedEpc epc)
    {
        if (epc.Scheme is not (EpcScheme.Sgtin96 or EpcScheme.Sgtin198) ||
            epc.CompanyPrefix is null || epc.ItemReference is null || epc.Serial is null || epc.ItemReference.Length < 1)
            return null;
        return $"urn:epc:id:sgtin:{epc.CompanyPrefix}.{epc.ItemReference[1..]}.{Uri.EscapeDataString(epc.Serial)}";
    }
}