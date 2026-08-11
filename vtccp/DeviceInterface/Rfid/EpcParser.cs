using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid;

/// <summary>
/// Top-level EPC parsing entry point.
///
/// Accepts raw EPC bytes or an uppercase hex string and dispatches to the appropriate
/// scheme decoder via <see cref="EpcSchemeDispatcher"/>.
///
/// Supported schemes:
///   SGTIN-96  (0x30) — Phase 0, fully implemented
///   SGTIN-198 (0x36) — Phase 0, fully implemented
///   SSCC-96   (0x31) — Phase 3 stub (returns Unknown with note)
///   All others       — Unknown (raw hex logged via UnknownSchemeDecoder)
///
/// Never throws — returns a best-effort result with <see cref="ParsedEpc.ParseWarning"/>
/// set on any error.
/// </summary>
public static class EpcParser
{
    /// <summary>
    /// Parse a raw EPC byte array.
    /// Delegates scheme dispatch to <see cref="EpcSchemeDispatcher"/>.
    /// </summary>
    public static ParsedEpc Parse(byte[] epcBytes) =>
        EpcSchemeDispatcher.Dispatch(epcBytes);

    /// <summary>
    /// Parse an EPC supplied as an uppercase hex string (e.g. "3034257BF40C0E4000000064").
    /// Whitespace is stripped before conversion.
    /// Returns <see cref="EpcScheme.Unknown"/> with a warning on malformed input.
    /// </summary>
    public static ParsedEpc ParseHex(string? hexString)
    {
        if (string.IsNullOrWhiteSpace(hexString))
        {
            return new ParsedEpc
            {
                EpcBytes     = [],
                Scheme       = EpcScheme.Unknown,
                ParseWarning = "EPC hex string is null or empty.",
            };
        }

        string normalized = hexString.Trim().Replace(" ", "").ToUpperInvariant();

        byte[] bytes;
        try
        {
            bytes = Convert.FromHexString(normalized);
        }
        catch (FormatException)
        {
            return new ParsedEpc
            {
                EpcBytes     = [],
                Scheme       = EpcScheme.Unknown,
                ParseWarning = $"EPC hex string is not valid hexadecimal: \"{hexString}\".",
            };
        }

        return EpcSchemeDispatcher.Dispatch(bytes);
    }
}
