// Copyright © 2026 VCCS. All rights reserved.

using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid.Schemes;

/// <summary>
/// Fallback decoder for EPC headers not recognised by <see cref="EpcSchemeDispatcher"/>.
///
/// Returns a <see cref="ParsedEpc"/> with <see cref="EpcScheme.Unknown"/>, the raw
/// EPC hex preserved, and a warning that includes the unrecognised header byte so it
/// can be logged and investigated.
///
/// All GS1 TDS 2.3 active scheme headers are dispatched before reaching this decoder.
/// Reaching this path means either:
///   - The tag is using a Phase 3+ scheme (not yet implemented), or
///   - The EPC memory bank is uninitialized / corrupt.
/// </summary>
public static class UnknownSchemeDecoder
{
    /// <summary>
    /// Decode an EPC with an unrecognised header byte.
    /// Logs the header and raw hex; never throws.
    /// </summary>
    public static ParsedEpc Decode(byte[] epcBytes)
    {
        byte header = epcBytes.Length > 0 ? epcBytes[0] : (byte)0;

        Dbg($"Unrecognised EPC header 0x{header:X2} — raw: {Convert.ToHexString(epcBytes)}");

        return new ParsedEpc
        {
            EpcBytes     = epcBytes,
            Scheme       = EpcScheme.Unknown,
            ParseWarning = $"Unrecognised EPC header 0x{header:X2}. " +
                           $"Raw EPC: {Convert.ToHexString(epcBytes)}",
        };
    }

    [System.Diagnostics.Conditional("DEBUG")]
    private static void Dbg(string msg) =>
        System.Diagnostics.Debug.WriteLine($"[UnknownScheme] {msg}");
}
