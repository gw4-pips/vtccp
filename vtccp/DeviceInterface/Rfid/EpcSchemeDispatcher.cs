// Copyright © 2026 VCCS. All rights reserved.

using DeviceInterface.Rfid.Models;
using DeviceInterface.Rfid.Schemes;

namespace DeviceInterface.Rfid;

/// <summary>
/// Inspects the header byte of a raw EPC and dispatches to the appropriate scheme decoder.
///
/// Header byte dispatch table (GS1 TDS 2.3 active schemes):
/// <code>
/// 0x30 → SGTIN-96  (Phase 0 — implemented)
/// 0x36 → SGTIN-198 (Phase 0 — implemented)
/// 0x31 → SSCC-96   (Phase 3 — stub)
/// all others → Unknown (logged with raw hex)
/// </code>
///
/// Returns a <see cref="ParsedEpc"/> with <see cref="EpcScheme.Unknown"/> for
/// unrecognised or unsupported schemes — never throws.
/// </summary>
public static class EpcSchemeDispatcher
{
    /// <summary>
    /// Dispatch the raw EPC bytes to the appropriate scheme decoder based on the header byte.
    /// Always returns a non-null result; sets <see cref="ParsedEpc.ParseWarning"/> on error.
    /// </summary>
    public static ParsedEpc Dispatch(byte[] epcBytes)
    {
        if (epcBytes is null || epcBytes.Length == 0)
        {
            return UnknownEpc(epcBytes ?? [], "EPC data is empty.");
        }

        byte header = epcBytes[0];

        try
        {
            return header switch
            {
                Sgtin96Decoder.Header  => Sgtin96Decoder.TryDecode(epcBytes)
                                          ?? UnknownEpc(epcBytes,
                                              $"SGTIN-96 decode failed (header=0x{header:X2})."),

                Sgtin198Decoder.Header => Sgtin198Decoder.TryDecode(epcBytes)
                                          ?? UnknownEpc(epcBytes,
                                              $"SGTIN-198 decode failed (header=0x{header:X2})."),

                Sscc96Decoder.Header   => Sscc96Decoder.DecodeStub(epcBytes),

                _                      => UnknownSchemeDecoder.Decode(epcBytes),
            };
        }
        catch (Exception ex)
        {
            return new ParsedEpc
            {
                EpcBytes     = epcBytes,
                Scheme       = EpcScheme.Unknown,
                ParseWarning = $"Dispatch exception for header 0x{header:X2}: {ex.Message}",
            };
        }
    }

    private static ParsedEpc UnknownEpc(byte[] epcBytes, string warning) => new()
    {
        EpcBytes     = epcBytes,
        Scheme       = EpcScheme.Unknown,
        ParseWarning = warning,
    };
}
