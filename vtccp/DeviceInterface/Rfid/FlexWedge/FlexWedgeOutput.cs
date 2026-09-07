using System.Globalization;

namespace DeviceInterface.Rfid.FlexWedge;

public enum FlexWedgeAppendKey { None, Enter, Tab }

/// <summary>Canonical fields available to a destination formatter.</summary>
public enum FlexWedgeField
{
    RawSource, RawEpc, Gtin14, Serial, EpcUri, Scheme, Filter, Partition, CompanyPrefix,
    ItemReference, GcpStatus, GcpRegisteredLength, Timestamp, Reader, Port, Rssi, Pc, Tid, LockStatus,
    DecodeWarnings, HeaderBits, FilterBits, PartitionBits, GcpBits, ItemReferenceBits, SerialBits
}

public sealed record FlexWedgeFieldMapping(FlexWedgeField Field, string Header);

/// <summary>
/// UI-independent output profile.  The same profile can feed a keyboard, clipboard,
/// focused-cell bridge, file, or API adapter.
/// </summary>
public sealed record FlexWedgeOutputProfile
{
    public string Name { get; init; } = "Default";
    public string Prefix { get; init; } = "";
    public string Suffix { get; init; } = "";
    public string Delimiter { get; init; } = "\t";
    public FlexWedgeAppendKey AppendKey { get; init; }
    public IReadOnlyList<FlexWedgeFieldMapping> Fields { get; init; } = [];

    public static FlexWedgeOutputProfile RawEpc { get; } = new()
    {
        Name = "Raw EPC", Fields = [new(FlexWedgeField.RawEpc, "EPC")]
    };
    public static FlexWedgeOutputProfile GtinAndSerial { get; } = new()
    {
        Name = "GTIN + serial", Fields = [new(FlexWedgeField.Gtin14, "GTIN-14"), new(FlexWedgeField.Serial, "Serial")]
    };
    public static FlexWedgeOutputProfile DelimitedMapped { get; } = new()
    {
        Name = "Delimited mapped fields",
        Fields = [new(FlexWedgeField.RawEpc, "EPC"), new(FlexWedgeField.Gtin14, "GTIN-14"), new(FlexWedgeField.Serial, "Serial")]
    };
    public static FlexWedgeOutputProfile CompleteDecoded { get; } = new()
    {
        Name = "Complete decoded result",
        Fields =
        [
            new(FlexWedgeField.RawEpc, "EPC"),
            new(FlexWedgeField.EpcUri, "EPC Tag URI"),
            new(FlexWedgeField.Scheme, "Scheme"),
            new(FlexWedgeField.Gtin14, "GTIN-14"),
            new(FlexWedgeField.Serial, "Serial"),
            new(FlexWedgeField.Filter, "Filter"),
            new(FlexWedgeField.Partition, "Partition"),
            new(FlexWedgeField.CompanyPrefix, "Company Prefix"),
            new(FlexWedgeField.ItemReference, "Item Reference"),
            new(FlexWedgeField.GcpStatus, "GCP Status"),
            new(FlexWedgeField.GcpRegisteredLength, "GCP Registered Length"),
            new(FlexWedgeField.Tid, "TID"),
            new(FlexWedgeField.LockStatus, "Lock Status")
        ]
    };

    public IReadOnlyList<string> Validate()
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(Name)) errors.Add("Profile name is required.");
        if (Delimiter.Length == 0) errors.Add("Delimiter is required.");
        if (Fields.Count == 0) errors.Add("At least one field mapping is required.");
        if (Fields.Any(x => string.IsNullOrWhiteSpace(x.Header))) errors.Add("Every field mapping requires a header.");
        if (Fields.GroupBy(x => x.Header, StringComparer.OrdinalIgnoreCase).Any(x => x.Count() > 1))
            errors.Add("Field mapping headers must be unique.");
        return errors;
    }
}

public static class FlexWedgeOutputFormatter
{
    public static string Format(FlexWedgeReadResult result, FlexWedgeOutputProfile profile)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(profile);
        var errors = profile.Validate();
        if (errors.Count != 0) throw new ArgumentException(string.Join(" ", errors), nameof(profile));
        string body = string.Join(profile.Delimiter, profile.Fields.Select(x => GetValue(result, x.Field)));
        return profile.Prefix + body + profile.Suffix + Append(profile.AppendKey);
    }

    public static IReadOnlyList<string> Headers(FlexWedgeOutputProfile profile)
    {
        ArgumentNullException.ThrowIfNull(profile);
        return profile.Fields.Select(x => x.Header).ToArray();
    }

    internal static string GetValue(FlexWedgeReadResult r, FlexWedgeField f) => f switch
    {
        FlexWedgeField.RawSource => r.RawSource ?? "",
        FlexWedgeField.RawEpc => r.RawEpcHex,
        FlexWedgeField.Gtin14 => r.Gtin14 ?? "",
        FlexWedgeField.Serial => r.Serial ?? "",
        FlexWedgeField.EpcUri => r.EpcUri ?? "",
        FlexWedgeField.Scheme => r.Scheme.ToString(),
        FlexWedgeField.Filter => Number(r.Filter),
        FlexWedgeField.Partition => Number(r.Partition),
        FlexWedgeField.CompanyPrefix => r.CompanyPrefix ?? "",
        FlexWedgeField.ItemReference => r.ItemReference ?? "",
        FlexWedgeField.GcpStatus => r.GcpStatus.ToString(),
        FlexWedgeField.GcpRegisteredLength => Number(r.GcpRegisteredLength),
        FlexWedgeField.Timestamp => r.Audit.Timestamp.ToString("O", CultureInfo.InvariantCulture),
        FlexWedgeField.Reader => r.Audit.Reader ?? "",
        FlexWedgeField.Port => r.Audit.Port ?? "",
        FlexWedgeField.Rssi => Number(r.Audit.Rssi),
        FlexWedgeField.Pc => r.Audit.PcWord?.ToString("X4", CultureInfo.InvariantCulture) ?? "",
        FlexWedgeField.Tid => r.Audit.Tid ?? "",
        FlexWedgeField.LockStatus => r.Audit.LockStatus ?? "",
        FlexWedgeField.DecodeWarnings => string.Join("; ", r.DecodeWarnings),
        FlexWedgeField.HeaderBits => Bits(r.HeaderBits),
        FlexWedgeField.FilterBits => Bits(r.FilterBits),
        FlexWedgeField.PartitionBits => Bits(r.PartitionBits),
        FlexWedgeField.GcpBits => Bits(r.GcpBits),
        FlexWedgeField.ItemReferenceBits => Bits(r.ItemReferenceBits),
        FlexWedgeField.SerialBits => Bits(r.SerialBits),
        _ => "",
    };

    private static string Number<T>(T? value) where T : struct, IFormattable =>
        value?.ToString(null, CultureInfo.InvariantCulture) ?? "";
    private static string Append(FlexWedgeAppendKey key) => key switch
    {
        FlexWedgeAppendKey.Enter => "\r\n", FlexWedgeAppendKey.Tab => "\t", _ => ""
    };
    private static string Bits(EpcBitField? field) =>
        field is null ? "" : $"{field.StartBit}-{field.EndBitInclusive}:{field.Bits}";
}

/// <summary>Boundary for clipboard, focused-cell, keyboard, or any text destination.</summary>
public interface IFlexWedgeTextSink { void WriteText(string text); }

public sealed class FlexWedgeTextSinkAdapter
{
    private readonly IFlexWedgeTextSink _sink;
    public FlexWedgeTextSinkAdapter(IFlexWedgeTextSink sink) => _sink = sink ?? throw new ArgumentNullException(nameof(sink));
    public void Deliver(FlexWedgeReadResult result, FlexWedgeOutputProfile profile) =>
        _sink.WriteText(FlexWedgeOutputFormatter.Format(result, profile));
}