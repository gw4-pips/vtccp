using DeviceInterface.Rfid.Models;
using DeviceInterface.Rfid.Schemes;

namespace DeviceInterface.Rfid;

/// <summary>
/// Top-level EPC parsing entry point.
/// Inspects the header byte of a raw EPC and dispatches to the appropriate scheme decoder.
/// Supported schemes: SGTIN-96 (0x30), SGTIN-198 (0x36).
/// Unknown headers return a <see cref="ParsedEpc"/> with <see cref="EpcScheme.Unknown"/>.
/// </summary>
public static class EpcParser
{
    /// <summary>
    /// Parse a raw EPC byte array.
    /// Never throws — returns a best-effort result with <see cref="ParsedEpc.ParseWarning"/> set on error.
    /// </summary>
    public static ParsedEpc Parse(byte[] epcBytes)
    {
        if (epcBytes is null || epcBytes.Length == 0)
        {
            return new ParsedEpc
            {
                EpcBytes     = epcBytes ?? [],
                Scheme       = EpcScheme.Unknown,
                ParseWarning = "EPC data is empty.",
            };
        }

        byte header = epcBytes[0];

        try
        {
            return header switch
            {
                Sgtin96Decoder.Header  => Sgtin96Decoder.TryDecode(epcBytes)
                                          ?? UnknownEpc(epcBytes, $"SGTIN-96 decode failed (header=0x{header:X2})."),

                Sgtin198Decoder.Header => Sgtin198Decoder.TryDecode(epcBytes)
                                          ?? UnknownEpc(epcBytes, $"SGTIN-198 decode failed (header=0x{header:X2})."),

                _                      => UnknownEpc(epcBytes, $"Unrecognised EPC header 0x{header:X2}."),
            };
        }
        catch (Exception ex)
        {
            return new ParsedEpc
            {
                EpcBytes     = epcBytes,
                Scheme       = EpcScheme.Unknown,
                ParseWarning = $"Parse exception: {ex.Message}",
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
