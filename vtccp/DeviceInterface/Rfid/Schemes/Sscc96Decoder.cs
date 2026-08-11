// Copyright © 2026 VCCS. All rights reserved.

using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid.Schemes;

/// <summary>
/// Placeholder decoder for SSCC-96 EPCs (96-bit, header = 0x31).
///
/// SSCC-96 bit layout (GS1 TDS 2.3):
///   bits  0– 7  : Header (8 bits) = 0x31
///   bits  8–10  : Filter (3 bits)
///   bits 11–13  : Partition (3 bits) — same table as SGTIN; L+K=17 for SSCC
///   bits 14–57  : Extension digit (1 bit) + Company Prefix (M bits) + Serial Ref (N bits)
///   bits 58–95  : Unallocated (must be zero)
///
/// Status: <b>Phase 3 stub</b> — returns <see cref="EpcScheme.Unknown"/> with a note.
/// Full SSCC-96 decode is scheduled for Phase 3 (full EPC scheme coverage).
/// </summary>
public static class Sscc96Decoder
{
    public const byte Header = 0x31;

    /// <summary>
    /// Return a stub result noting that SSCC-96 is not yet implemented.
    /// The raw EPC hex is preserved for logging; no decode fields are set.
    /// </summary>
    public static ParsedEpc DecodeStub(byte[] epcBytes) => new()
    {
        EpcBytes     = epcBytes,
        Scheme       = EpcScheme.Unknown,
        ParseWarning = $"SSCC-96 (header=0x{Header:X2}) is not yet decoded " +
                       "(Phase 3 — full scheme coverage). Raw EPC: " +
                       Convert.ToHexString(epcBytes),
    };
}
