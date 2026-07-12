namespace DeviceInterface.Rfid;

/// <summary>
/// MTI LLCS (Linkage Layer Communication Specification) binary packet constants
/// for the RU-824-100 UHF RFID reader.
///
/// Protocol confirmed from MTI RFID Explorer v2.0.1 source (Linkage.cs, Global.cs).
///
/// HOST → READER  (16-byte command packet):
///   [0-3] "CITM" (0x43 0x49 0x54 0x4D)
///   [4]   0xFF
///   [5]   Command ID (see <see cref="Cmd"/>)
///   [6-13] Parameters (8 bytes, unused params = 0x00)
///   [14-15] ~CRC-16/CCITT over bytes [0..13], little-endian
///
/// READER → HOST  response packets:
///   4-byte sync header: TypeByte + 'I' + 'T' + 'M'
///   TypeByte = 'R' (0x52) → 16-byte common response
///              'B' (0x42) → 16-byte BEGIN notification
///              'E' (0x45) → 16-byte END notification
///              'I' (0x49) → 64-byte INVENTORY tag data
///              'A' (0x41) → 64-byte TAG_ACCESS data
///
///   Small response (16 bytes total, sync at [0-3]):
///     [4]   0xFF
///     [5]   echoed command byte
///     [6]   result code (0x00 = OK)
///     [7-13] response payload
///     [14-15] ~CRC-16/CCITT
///
///   Large response (64 bytes total, sync at [0-3]):
///     [4-6]  additional header bytes
///     [7]    command byte (0xF5=INVENTORY, 0xF6=TAG_ACCESS, 0xF1=END, 0xF0=BEGIN)
///     [8-9]  divide sequence info
///     [10-11] INFO_LENGTH (data length in 4-byte DWORDs, little-endian)
///     [12-13] additional header
///     [14..] INFO_DATA: PC_Word(2B,BE) + EPC(variable) + TagCRC(2B)
///     [62-63] packet CRC
/// </summary>
internal static class LlcsProtocol
{
    // ── Command packet header ──────────────────────────────────────────────────

    public static readonly byte[] TxHeader = [0x43, 0x49, 0x54, 0x4D, 0xFF];
    public const int TxPacketLen = 16;
    public const int TxCmdOffset = 5;
    public const int TxCrcOffset = 14;

    // ── Response packet sync bytes ─────────────────────────────────────────────

    // Second/third/fourth bytes of all response packets
    public const byte SyncByte2 = (byte)'I';  // 0x49
    public const byte SyncByte3 = (byte)'T';  // 0x54
    public const byte SyncByte4 = (byte)'M';  // 0x4D

    // First byte (type) of each packet kind
    public const byte TypeCommonResponse = (byte)'R';  // 0x52 — 16-byte response to a command
    public const byte TypeBegin          = (byte)'B';  // 0x42 — inventory started
    public const byte TypeEnd            = (byte)'E';  // 0x45 — inventory complete
    public const byte TypeInventory      = (byte)'I';  // 0x49 — tag data (64-byte packet)
    public const byte TypeTagAccess      = (byte)'A';  // 0x41 — tag access data (64-byte)

    // ── Small response (16-byte) layout ───────────────────────────────────────

    public const int SmallPktLen = 16;
    public const int SmallCmdOffset = 5;   // echoed command
    public const int SmallResultOffset = 6; // result code
    public const int SmallPayloadOffset = 7;
    public const byte ResultOk = 0x00;

    // ── Large response (64-byte) layout ───────────────────────────────────────

    public const int LargePktLen = 64;
    public const int LargeCmdOffset = 7;      // command byte (END, INVENTORY, etc.)
    public const int LargeInfoLenOffset = 10; // INFO_LENGTH (LE UInt16, in DWORDs)
    public const int LargeInfoDataOffset = 14; // start of INFO_DATA

    // ── ENUM_CMD values relevant for inventory ─────────────────────────────────

    public static class Cmd
    {
        public const byte TagInventory  = 0x40;  // l8K6C_TAG_INVENTORY
        public const byte ControlCancel = 0x50;  // CONTROL_CANCEL
        public const byte ControlAbort  = 0x51;  // CONTROL_ABORT
        public const byte MacGetDebug   = 0x61;  // MAC_GET_DEBUG (used as connectivity ping)
    }

    // ── ENUM_CMD response/notification bytes ──────────────────────────────────

    public static class Notify
    {
        public const byte Begin      = 0xF0;  // inventory started
        public const byte End        = 0xF1;  // inventory complete
        public const byte Inventory  = 0xF5;  // tag data packet
        public const byte TagAccess  = 0xF6;  // tag access data packet
        public const byte TimeOut    = 0xFE;
        public const byte Nothing    = 0xFF;
    }

    // ── CRC-16/CCITT (poly=0x1021, init=0xFFFF, output=~crc) ─────────────────

    public static ushort ComputeCrc(byte[] buf, int offset, int byteLen)
    {
        const int poly = 0x1021;
        ushort crc = 0xFFFF;
        int bits = byteLen * 8;
        int byteIdx = 0;
        ushort data = 0;

        for (int i = 0; i < bits; i++)
        {
            if (i % 8 == 0)
            {
                data = (ushort)(buf[offset + byteIdx] << 8);
                byteIdx++;
            }
            ushort val = (ushort)(crc ^ data);
            crc  = (ushort)((crc  << 1) & 0xFFFF);
            data = (ushort)((data << 1) & 0xFFFF);
            if ((val & 0x8000) != 0)
                crc ^= (ushort)poly;
        }
        return (ushort)(crc & 0xFFFF);
    }

    /// <summary>Build a 16-byte command packet ready for serial transmission.</summary>
    public static byte[] BuildCommand(byte cmd, byte p0 = 0, byte p1 = 0,
        byte p2 = 0, byte p3 = 0, byte p4 = 0,
        byte p5 = 0, byte p6 = 0, byte p7 = 0)
    {
        var pkt = new byte[TxPacketLen];
        TxHeader.CopyTo(pkt, 0);
        pkt[5] = cmd;
        pkt[6] = p0;
        pkt[7] = p1;
        pkt[8] = p2;
        pkt[9] = p3;
        pkt[10] = p4;
        pkt[11] = p5;
        pkt[12] = p6;
        pkt[13] = p7;
        ushort crc = ComputeCrc(pkt, 0, TxCrcOffset);
        crc = (ushort)~crc;
        pkt[14] = (byte)(crc & 0xFF);
        pkt[15] = (byte)((crc >> 8) & 0xFF);
        return pkt;
    }

    /// <summary>
    /// Inspect the first byte of a response packet to determine the number of
    /// remaining bytes to read after the 4-byte sync header is confirmed.
    /// Returns -1 for an unknown type.
    /// </summary>
    public static int GetRemainingBytesAfterSync(byte typeByte) => typeByte switch
    {
        TypeCommonResponse or TypeBegin or TypeEnd => SmallPktLen - 4,
        TypeInventory or TypeTagAccess             => LargePktLen - 4,
        _                                          => -1
    };
}
