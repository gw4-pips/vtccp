using System.Collections.ObjectModel;
using System.Text.Json.Serialization;
using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid.FlexWedge;

/// <summary>Outcome of GCP validation as exposed by FlexWedge.</summary>
public enum FlexWedgeGcpStatus
{
    NotChecked,
    Valid,
    NotFound,
    Invalid,
    /// <summary>A registered prefix was found, but its registered length differs from the EPC partition.</summary>
    Mismatch,
}

/// <summary>An EPC bit range. Bit zero is the most-significant bit of the first EPC byte.</summary>
public sealed record EpcBitField(string Name, int StartBit, int Length, string Bits)
{
    public int EndBitInclusive => StartBit + Length - 1;
}

/// <summary>Reader-supplied facts retained with a canonical FlexWedge read.</summary>
public sealed record FlexWedgeReaderAudit(
    DateTimeOffset Timestamp,
    string? Reader,
    string? Port,
    int? Rssi,
    ushort? PcWord,
    string? Tid,
    string? LockStatus);

/// <summary>
/// Destination-independent, immutable representation of one RFID observation. Raw source
/// text is retained verbatim; decoded fields never replace it.
/// </summary>
public sealed record FlexWedgeReadResult
{
    /// <summary>Raw EPC text as supplied by the caller, if it was supplied as text.</summary>
    public string? RawSource { get; init; }
    /// <summary>Normalized uppercase hexadecimal EPC value; empty for malformed input.</summary>
    public required string RawEpcHex { get; init; }
    /// <summary>Raw EPC bytes computed from the immutable normalized hexadecimal value.</summary>
    [JsonIgnore]
    public ReadOnlyCollection<byte> RawEpcBytes =>
        Array.AsReadOnly(string.IsNullOrEmpty(RawEpcHex) ? [] : Convert.FromHexString(RawEpcHex));
    public EpcScheme Scheme { get; init; }
    public int? Filter { get; init; }
    public int? Partition { get; init; }
    public string? CompanyPrefix { get; init; }
    public string? ItemReference { get; init; }
    public string? Gtin14 { get; init; }
    public string? Serial { get; init; }
    public string? EpcUri { get; init; }
    public EpcBitField? HeaderBits { get; init; }
    public EpcBitField? FilterBits { get; init; }
    public EpcBitField? PartitionBits { get; init; }
    public EpcBitField? GcpBits { get; init; }
    public EpcBitField? ItemReferenceBits { get; init; }
    public EpcBitField? SerialBits { get; init; }
    public required FlexWedgeReaderAudit Audit { get; init; }
    public IReadOnlyList<string> DecodeWarnings { get; init; } = [];
    public FlexWedgeGcpStatus GcpStatus { get; init; } = FlexWedgeGcpStatus.NotChecked;
    public int? GcpRegisteredLength { get; init; }
}