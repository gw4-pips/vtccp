namespace DeviceInterface.Rfid;

/// <summary>
/// Frame constants, builder, and CRC for the GoToTags E310 UHF RFID Reader
/// (Impinj E310 chipset).
///
/// Wire format — "Common Command Communication Protocol":
///   [ 0xFF | DataLen (1B) | CmdCode (1B) | Data (DataLen bytes) | CRC_Hi | CRC_Lo ]
///
/// CRC-16/CCITT: init=0xFFFF, poly=0x1021.
/// Computed over frame bytes [1 .. end-2] — i.e. DataLen through end of Data,
/// skipping the 0xFF header and the two CRC bytes.  Stored big-endian.
///
/// Protocol source: GoToTags UHF RFID Reader Communication Protocol rev 5-30-23
/// (gitlab.com/gototags/public → UHF RFID/Readers/GoToTags/docs/).
/// </summary>
internal static class GoToTagsE310Protocol
{
    // ── Frame structure ──────────────────────────────────────────────────────

    public const byte Header = 0xFF;

    // ── APP-layer command codes ──────────────────────────────────────────────

    /// <summary>
    /// Single Tag Inventory (0x21).
    /// Scans until the first tag responds or timeout expires.
    /// Returns EPC + TagCRC in the response data field.
    /// </summary>
    public const byte CmdSingleTagInventory = 0x21;

    // ── Status codes (response Data[0..1]) ──────────────────────────────────

    /// <summary>Command executed successfully, tag present in response.</summary>
    public const ushort StatusOk = 0x0000;

    // ── Limits ──────────────────────────────────────────────────────────────

    public const int MaxTimeoutMs  = 65535;
    public const int PingTimeoutMs = 200;

    // ── Option byte ─────────────────────────────────────────────────────────

    /// <summary>
    /// Option = 0x00: Select-Option bits = 0 (no filter / no Tag Singulation),
    /// BIT4 = 0 (no Metadata Flags field).
    /// Response returns EPC + TagCRC only.
    /// </summary>
    public const byte OptionNoFilterNoMeta = 0x00;

    // ── Frame builders ───────────────────────────────────────────────────────

    /// <summary>
    /// Build a Single Tag Inventory (0x21) command frame.
    /// Data field: Timeout(2 bytes, big-endian ms) + Option(1 byte = 0x00).
    /// </summary>
    public static byte[] BuildSingleInventoryCmd(int timeoutMs)
    {
        timeoutMs = Math.Clamp(timeoutMs, 1, MaxTimeoutMs);
        Span<byte> raw = stackalloc byte[6];
        raw[0] = Header;
        raw[1] = 3;
        raw[2] = CmdSingleTagInventory;
        raw[3] = (byte)((timeoutMs >> 8) & 0xFF);
        raw[4] = (byte)(timeoutMs & 0xFF);
        raw[5] = OptionNoFilterNoMeta;
        return AppendCrc(raw);
    }

    // ── Response parsing ─────────────────────────────────────────────────────

    /// <summary>
    /// Attempt to extract an EPC from a validated response frame for command 0x21.
    ///
    /// Response Data layout (Option = 0x00):
    ///   [ Status(2B) | Option(1B) | EPC(N bytes) | TagCRC(2B) ]
    ///
    /// EPC byte count = DataLen - 5  (Status:2 + Option:1 + TagCRC:2).
    /// Returns null if status ≠ 0x0000, DataLen too small, or EPC length is zero.
    /// </summary>
    public static byte[]? TryExtractEpc(byte[] frame)
    {
        if (frame.Length < 9) return null;

        int dataLen = frame[1];
        int epcLen  = dataLen - 5;
        if (epcLen <= 0) return null;

        ushort status = (ushort)((frame[3] << 8) | frame[4]);
        if (status != StatusOk) return null;

        var epc = new byte[epcLen];
        Array.Copy(frame, 5, epc, 0, epcLen);
        return epc;
    }

    // ── Frame reading helper ─────────────────────────────────────────────────

    /// <summary>
    /// Total byte count of a complete response frame given its DataLen byte.
    /// FF(1) + DataLen(1) + CmdCode(1) + Data(DataLen) + CRC(2).
    /// </summary>
    public static int FrameSize(int dataLen) => 5 + dataLen;

    // ── CRC-16/CCITT ─────────────────────────────────────────────────────────

    /// <summary>
    /// Compute CRC-16/CCITT (init=0xFFFF, poly=0x1021) over <paramref name="count"/>
    /// bytes starting at <paramref name="offset"/> in <paramref name="buf"/>.
    /// </summary>
    public static ushort CalcCrc(byte[] buf, int offset, int count)
    {
        ushort crc = 0xFFFF;
        for (int i = offset; i < offset + count; i++)
        {
            byte b = buf[i];
            for (int bit = 7; bit >= 0; bit--)
            {
                bool xorFlag = (crc & 0x8000) != 0;
                crc <<= 1;
                if (((b >> bit) & 1) == 1) crc |= 1;
                if (xorFlag) crc ^= 0x1021;
            }
        }
        return crc;
    }

    /// <summary>
    /// Verify the CRC-16 in the last two bytes of <paramref name="frame"/>.
    /// CRC covers frame[1..frameLen-3].
    /// </summary>
    public static bool VerifyCrc(byte[] frame, int frameLen)
    {
        ushort computed = CalcCrc(frame, 1, frameLen - 3);
        ushort embedded = (ushort)((frame[frameLen - 2] << 8) | frame[frameLen - 1]);
        return computed == embedded;
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    private static byte[] AppendCrc(ReadOnlySpan<byte> raw)
    {
        var buf = raw.ToArray();
        ushort crc = CalcCrc(buf, 1, raw.Length - 1);
        var result = new byte[raw.Length + 2];
        buf.CopyTo(result, 0);
        result[raw.Length]     = (byte)((crc >> 8) & 0xFF);
        result[raw.Length + 1] = (byte)(crc & 0xFF);
        return result;
    }
}
