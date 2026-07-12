using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid.Schemes;

/// <summary>
/// Decodes SGTIN-96 EPCs (96-bit, header = 0x30) into their constituent fields
/// and reconstructs the GTIN-14 using GS1 EPC Tag Data Standard v1.15 definitions.
///
/// Bit layout (MSB first, bit 0 = most significant):
///   bits  0– 7  : Header   (8 bits)  = 0x30
///   bits  8–10  : Filter   (3 bits)
///   bits 11–13  : Partition (3 bits)  → determines M and N
///   bits 14–(13+M) : Company Prefix (M bits)
///   bits (14+M)–(13+M+N) : Item Reference (N bits, includes GTIN indicator)
///   bits (14+M+N)–95 : Serial (38 bits)
///
/// Partition table (GS1 TDS Table 14-1):
///   P | M  | L  | N  | K   (L = GCP decimal digits, K = Item ref decimal digits, L+K = 13)
///   0 | 40 | 12 |  4 |  1
///   1 | 37 | 11 |  7 |  2
///   2 | 34 | 10 | 10 |  3
///   3 | 30 |  9 | 14 |  4
///   4 | 27 |  8 | 17 |  5
///   5 | 24 |  7 | 20 |  6
///   6 | 20 |  6 | 24 |  7
///
/// GTIN-14 reconstruction:
///   payload13 = GCP.PadLeft(L,'0') + ItemRef.PadLeft(K,'0')  [13 digits]
///   GTIN-14   = payload13 + GS1CheckDigit(payload13)          [14 digits]
/// </summary>
public static class Sgtin96Decoder
{
    public const byte Header = 0x30;
    private const int ExpectedBytes = 12;

    // Partition table: (M, L, N, K) — L+K = 13
    private static readonly (int M, int L, int N, int K)[] Partitions =
    [
        (40, 12,  4,  1),
        (37, 11,  7,  2),
        (34, 10, 10,  3),
        (30,  9, 14,  4),
        (27,  8, 17,  5),
        (24,  7, 20,  6),
        (20,  6, 24,  7),
    ];

    /// <summary>
    /// Attempt to decode a 12-byte EPC as SGTIN-96.
    /// Returns null if the header byte does not match or the data is malformed.
    /// </summary>
    public static ParsedEpc? TryDecode(byte[] epcBytes)
    {
        if (epcBytes is null || epcBytes.Length < ExpectedBytes) return null;

        byte header = GetBits8(epcBytes, 0);
        if (header != Header) return null;

        int filter    = (int)GetBits(epcBytes, 8,  3);
        int partition = (int)GetBits(epcBytes, 11, 3);

        if (partition > 6) return new ParsedEpc
        {
            EpcBytes    = epcBytes,
            Scheme      = EpcScheme.Sgtin96,
            ParseWarning = $"Unsupported partition value {partition}",
        };

        var (M, L, N, K) = Partitions[partition];

        ulong gcpRaw     = GetBits(epcBytes, 14,       M);
        ulong itemRefRaw = GetBits(epcBytes, 14 + M,   N);
        ulong serialRaw  = GetBits(epcBytes, 14 + M + N, 38);

        string gcpStr     = gcpRaw.ToString().PadLeft(L, '0');
        string itemRefStr = itemRefRaw.ToString().PadLeft(K, '0');
        string serial     = serialRaw.ToString();

        string payload13  = gcpStr + itemRefStr;
        string gtin14     = payload13 + Gs1CheckDigit(payload13);

        return new ParsedEpc
        {
            EpcBytes      = epcBytes,
            Scheme        = EpcScheme.Sgtin96,
            Filter        = filter,
            Partition     = partition,
            CompanyPrefix = gcpStr,
            ItemReference = itemRefStr,
            Serial        = serial,
            Gtin14        = gtin14,
        };
    }

    // ── Bit extraction ─────────────────────────────────────────────────────────

    /// <summary>Extract up to 64 bits starting at <paramref name="startBit"/> (0=MSB of byte 0).</summary>
    internal static ulong GetBits(byte[] data, int startBit, int count)
    {
        ulong result = 0;
        for (int i = 0; i < count; i++)
        {
            int absoluteBit = startBit + i;
            int byteIndex   = absoluteBit / 8;
            int bitInByte   = 7 - (absoluteBit % 8); // MSB first
            if (byteIndex >= data.Length) break;
            result = (result << 1) | (byte)((data[byteIndex] >> bitInByte) & 1);
        }
        return result;
    }

    private static byte GetBits8(byte[] data, int startBit) => (byte)GetBits(data, startBit, 8);

    // ── GS1 check digit (mod-10) ──────────────────────────────────────────────

    /// <summary>
    /// Compute the GS1 check digit for a 13-digit payload string.
    /// Weights alternate 3,1 from the rightmost digit inward.
    /// Result is the single check character to append to make GTIN-14.
    /// </summary>
    internal static char Gs1CheckDigit(string digits13)
    {
        if (digits13.Length != 13)
            throw new ArgumentException($"Expected 13 digits, got {digits13.Length}.", nameof(digits13));
        int sum = 0;
        for (int i = 0; i < 13; i++)
        {
            // Index 12 (rightmost) gets weight 3; 11 gets weight 1; etc.
            int weight = ((12 - i) % 2 == 0) ? 3 : 1;
            sum += (digits13[i] - '0') * weight;
        }
        return (char)('0' + (10 - (sum % 10)) % 10);
    }
}
