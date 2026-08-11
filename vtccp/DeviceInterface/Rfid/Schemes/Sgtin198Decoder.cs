using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid.Schemes;

/// <summary>
/// Decodes SGTIN-198 EPCs (198-bit, header = 0x36) into their constituent fields.
/// SGTIN-198 differs from SGTIN-96 only in the serial number field:
///   - Serial (140 bits) — packed 7-bit ASCII, up to 20 printable characters.
///
/// Bit layout:
///   bits  0– 7  : Header   (8 bits) = 0x36
///   bits  8–10  : Filter   (3 bits)
///   bits 11–13  : Partition (3 bits) → same table as SGTIN-96
///   bits 14–(13+M) : Company Prefix (M bits)
///   bits (14+M)–(13+M+N) : Item Reference (N bits)
///   bits (14+M+N)–(53+M+N) : Serial (140 bits, 7-bit ASCII packed, up to 20 chars)
///
/// Total: 8+3+3+44+140 = 198 bits → 25 bytes (padded to 200 bits in EPC memory).
/// </summary>
public static class Sgtin198Decoder
{
    public const byte Header = 0x36;
    private const int ExpectedBits = 198;
    private const int ExpectedBytes = 25; // ceil(198/8)

    public static ParsedEpc? TryDecode(byte[] epcBytes)
    {
        if (epcBytes is null || epcBytes.Length < ExpectedBytes) return null;

        byte header = (byte)Sgtin96Decoder.GetBits(epcBytes, 0, 8);
        if (header != Header) return null;

        int filter    = (int)Sgtin96Decoder.GetBits(epcBytes, 8,  3);
        int partition = (int)Sgtin96Decoder.GetBits(epcBytes, 11, 3);

        if (partition > 6) return new ParsedEpc
        {
            EpcBytes     = epcBytes,
            Scheme       = EpcScheme.Sgtin198,
            ParseWarning = $"Unsupported partition value {partition}",
        };

        // Reuse partition table from Sgtin96Decoder via reflection on the same layout
        (int M, int L, int N, int K) = GetPartition(partition);

        ulong gcpRaw     = Sgtin96Decoder.GetBits(epcBytes, 14,       M);
        ulong itemRefRaw = Sgtin96Decoder.GetBits(epcBytes, 14 + M,   N);

        // Serial: 140 bits of packed 7-bit ASCII starting at bit (14+M+N)
        string serial = Extract7BitAsciiSerial(epcBytes, 14 + M + N, 140);

        string gcpStr = gcpRaw.ToString().PadLeft(L, '0');

        // GS1 GTIN-14: indicator(1) + GCP(L) + item_body(K-1) + check(1)
        ulong pow10Km1   = Pow10((ulong)(K - 1));
        int   indicator  = (int)(itemRefRaw / pow10Km1);
        string itemBody  = (itemRefRaw % pow10Km1).ToString().PadLeft(K - 1, '0');
        string payload13 = indicator.ToString() + gcpStr + itemBody;
        string gtin14    = payload13 + Sgtin96Decoder.Gs1CheckDigit(payload13);
        string itemRefStr = indicator.ToString() + itemBody;

        return new ParsedEpc
        {
            EpcBytes      = epcBytes,
            Scheme        = EpcScheme.Sgtin198,
            Filter        = filter,
            Partition     = partition,
            CompanyPrefix = gcpStr,
            ItemReference = itemRefStr,
            Serial        = serial,
            Gtin14        = gtin14,
        };
    }

    private static string Extract7BitAsciiSerial(byte[] data, int startBit, int totalBits)
    {
        int charCount = totalBits / 7; // 140 / 7 = 20
        var sb = new System.Text.StringBuilder(charCount);
        for (int c = 0; c < charCount; c++)
        {
            char ch = (char)Sgtin96Decoder.GetBits(data, startBit + c * 7, 7);
            if (ch == '\0') break; // null-terminated serial
            sb.Append(ch);
        }
        return sb.ToString();
    }

    private static ulong Pow10(ulong n) { ulong r = 1; for (ulong i = 0; i < n; i++) r *= 10; return r; }

    private static (int M, int L, int N, int K) GetPartition(int p) => p switch
    {
        0 => (40, 12,  4,  1),
        1 => (37, 11,  7,  2),
        2 => (34, 10, 10,  3),
        3 => (30,  9, 14,  4),
        4 => (27,  8, 17,  5),
        5 => (24,  7, 20,  6),
        6 => (20,  6, 24,  7),
        _ => throw new ArgumentOutOfRangeException(nameof(p)),
    };
}
