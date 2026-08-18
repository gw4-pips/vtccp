namespace DeviceInterface.Rfid.Models;

/// <summary>
/// Result of a single EPC read attempt from the RFID reader.
/// </summary>
public sealed record EpcReadResult
{
    /// <summary>
    /// Raw EPC data bytes as returned from the tag's EPC memory bank
    /// (PC word stripped — pure EPC binary, typically 12 bytes for SGTIN-96).
    /// </summary>
    public required byte[] EpcBytes { get; init; }

    /// <summary>EpcBytes expressed as an uppercase hex string, e.g. "303405E801B6A5D3..."</summary>
    public string EpcHex => Convert.ToHexString(EpcBytes);

    /// <summary>PC word from the tag response (Protocol Control; upper 5 bits = EPC word count).</summary>
    public ushort PcWord { get; init; }

    /// <summary>UTC timestamp of when this tag was observed during inventory.</summary>
    public DateTimeOffset ReadTime { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>RSSI or signal strength if available; null otherwise.</summary>
    public int? Rssi { get; init; }

    /// <summary>
    /// Tag Identifier (TID) hex string from the tag's TID memory bank, if read.
    /// Populated either automatically (when the reader delivers TID in the inventory
    /// callback) or explicitly via <c>AsReaderP35UEpcReader.ReadTidAsync</c>.
    /// Null if TID reading was not requested, timed out, or the DLL defect prevented
    /// delivery (see AsReader TID defect report).
    /// </summary>
    public string? Tid { get; init; }

    /// <summary>
    /// EPC memory bank lock status as reported by the reader's lock-check command
    /// (e.g. <c>AsReaderP35UEpcReader.ReadLockStatusAsync</c> → SDK CheckTagStatus).
    /// Values: "PermaLocked" / "Locked" / "Unlocked" / "Unknown".
    /// Null when lock status was not queried or the query was rejected.
    /// </summary>
    public string? LockStatus { get; init; }
}
