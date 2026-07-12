using System.IO.Ports;
using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid;

/// <summary>
/// RFID reader implementation for the MTI RU-824-100 UHF reader using the
/// LLCS (Linkage Layer Communication Specification) binary packet protocol
/// over a USB Virtual COM Port (FTDI VCP) at 115200 8N1.
///
/// Protocol source: MTI RFID Explorer v2.0.1 Linkage.cs / Global.cs.
/// No native MTI DLLs (Transfer.dll / ftd2xx.dll) required — pure managed serial I/O.
/// </summary>
public sealed class MtiLlcsEpcReader : IEpcReader
{
    private SerialPort? _port;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private bool _disposed;

    // ── IEpcReader ─────────────────────────────────────────────────────────────

    public bool IsConnected => _port?.IsOpen == true;

    public async Task ConnectAsync(string portName, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (IsConnected)
                throw new InvalidOperationException("Reader is already connected. Call DisconnectAsync first.");

            var sp = new SerialPort(portName, 115200, Parity.None, 8, StopBits.One)
            {
                Handshake = Handshake.None,
                ReadTimeout  = 3000,
                WriteTimeout = 3000,
                DtrEnable    = false,
                RtsEnable    = false,
            };

            sp.Open();
            sp.DiscardInBuffer();
            sp.DiscardOutBuffer();
            _port = sp;

            // Ping with MAC_GET_DEBUG to confirm LLCS communication
            await PingAsync(ct).ConfigureAwait(false);
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task DisconnectAsync(CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            ClosePort();
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<IReadOnlyList<EpcReadResult>> TriggerInventoryAsync(
        TimeSpan timeout,
        CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            EnsureConnected();
            return await Task.Run(() => RunInventorySync(timeout, ct), ct)
                             .ConfigureAwait(false);
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task CancelAsync(CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (!IsConnected) return;
            await Task.Run(() => SendCancel(), ct).ConfigureAwait(false);
        }
        finally
        {
            _lock.Release();
        }
    }

    // ── Core inventory cycle ───────────────────────────────────────────────────

    private IReadOnlyList<EpcReadResult> RunInventorySync(TimeSpan timeout, CancellationToken ct)
    {
        var port = _port!;
        var results = new List<EpcReadResult>();
        var seenEpcs = new HashSet<string>(StringComparer.Ordinal);
        var deadline = DateTimeOffset.UtcNow + timeout;

        // 1. Send INVENTORY command (SelectOpsFlag=0, PostMatchFlag=0)
        byte[] cmd = LlcsProtocol.BuildCommand(LlcsProtocol.Cmd.TagInventory, 0, 0);
        port.Write(cmd, 0, cmd.Length);

        // 2. Read and verify ACK (16-byte common response, type 'R')
        if (!TryReadAck(port, LlcsProtocol.Cmd.TagInventory, out _))
            return results;

        // 3. Loop reading tag packets until END, timeout, or cancellation
        var buf = new byte[LlcsProtocol.LargePktLen];
        while (DateTimeOffset.UtcNow < deadline && !ct.IsCancellationRequested)
        {
            byte typeByte;
            if (!TryFindSyncHeader(port, deadline, ct, out typeByte))
                break;

            int remaining = LlcsProtocol.GetRemainingBytesAfterSync(typeByte);
            if (remaining < 0) continue;

            // buf[0-3] already has the sync header; read the rest
            buf[0] = typeByte;
            buf[1] = LlcsProtocol.SyncByte2;
            buf[2] = LlcsProtocol.SyncByte3;
            buf[3] = LlcsProtocol.SyncByte4;

            if (!ReadExactly(port, buf, 4, remaining, deadline, ct))
                break;

            if (typeByte == LlcsProtocol.TypeEnd)
                break; // Inventory cycle complete

            if (typeByte == LlcsProtocol.TypeInventory || typeByte == LlcsProtocol.TypeTagAccess)
            {
                var result = ExtractEpcFromLargePacket(buf);
                if (result is not null && seenEpcs.Add(result.EpcHex))
                    results.Add(result);
            }
        }

        return results;
    }

    // ── Packet helpers ─────────────────────────────────────────────────────────

    /// <summary>
    /// Scan the input stream one byte at a time until we confirm a 4-byte LLCS sync header.
    /// Returns the type byte (first sync byte) when found.
    /// </summary>
    private static bool TryFindSyncHeader(SerialPort port, DateTimeOffset deadline,
        CancellationToken ct, out byte typeByte)
    {
        typeByte = 0;
        Span<byte> window = stackalloc byte[4];
        int filled = 0;

        while (DateTimeOffset.UtcNow < deadline && !ct.IsCancellationRequested)
        {
            int b;
            try { b = port.ReadByte(); }
            catch (TimeoutException) { return false; }

            byte byt = (byte)b;

            // The sync sequence is: [type] 'I' 'T' 'M'
            // type ∈ {'R','B','E','I','A','S','F'}
            if (filled == 0)
            {
                if (IsValidTypeByte(byt)) { window[0] = byt; filled = 1; }
            }
            else if (filled == 1)
            {
                if (byt == LlcsProtocol.SyncByte2) { window[1] = byt; filled = 2; }
                else { filled = IsValidTypeByte(byt) ? 1 : 0; window[0] = byt; }
            }
            else if (filled == 2)
            {
                if (byt == LlcsProtocol.SyncByte3) { window[2] = byt; filled = 3; }
                else { filled = 0; } // restart
            }
            else // filled == 3
            {
                if (byt == LlcsProtocol.SyncByte4)
                {
                    typeByte = window[0];
                    return true;
                }
                filled = 0;
            }
        }
        return false;
    }

    private static bool IsValidTypeByte(byte b) =>
        b is (byte)'R' or (byte)'B' or (byte)'E' or (byte)'I' or (byte)'A' or (byte)'S' or (byte)'F';

    private static bool ReadExactly(SerialPort port, byte[] buf, int offset, int count,
        DateTimeOffset deadline, CancellationToken ct)
    {
        int read = 0;
        while (read < count && DateTimeOffset.UtcNow < deadline && !ct.IsCancellationRequested)
        {
            try
            {
                int n = port.Read(buf, offset + read, count - read);
                read += n;
            }
            catch (TimeoutException) { return false; }
        }
        return read == count;
    }

    /// <summary>
    /// Read a 16-byte common response and verify the echoed command and result code.
    /// Assumes the 4-byte sync header has NOT yet been consumed (we search for it).
    /// </summary>
    private static bool TryReadAck(SerialPort port, byte expectedCmd, out byte resultCode)
    {
        resultCode = 0xFF;
        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(3);
        var ct = CancellationToken.None;

        byte typeByte;
        if (!TryFindSyncHeader(port, deadline, ct, out typeByte))
            return false;

        if (typeByte != LlcsProtocol.TypeCommonResponse) return false;

        var buf = new byte[LlcsProtocol.SmallPktLen];
        buf[0] = typeByte;
        buf[1] = LlcsProtocol.SyncByte2;
        buf[2] = LlcsProtocol.SyncByte3;
        buf[3] = LlcsProtocol.SyncByte4;

        if (!ReadExactly(port, buf, 4, LlcsProtocol.SmallPktLen - 4, deadline, ct))
            return false;

        resultCode = buf[LlcsProtocol.SmallResultOffset];
        return buf[LlcsProtocol.SmallCmdOffset] == expectedCmd
            && resultCode == LlcsProtocol.ResultOk;
    }

    /// <summary>
    /// Extract EPC data from a 64-byte large response packet.
    /// INFO_DATA layout: PC_word(2B, BE) + EPC(epcWords×2 bytes) + TagCRC(2B).
    /// </summary>
    private static EpcReadResult? ExtractEpcFromLargePacket(byte[] pkt)
    {
        if (pkt.Length < LlcsProtocol.LargePktLen) return null;

        int dataOffset = LlcsProtocol.LargeInfoDataOffset;
        // PC word: big-endian UInt16 at INFO_DATA offset 0
        ushort pcWord = (ushort)((pkt[dataOffset] << 8) | pkt[dataOffset + 1]);
        int epcWords = (pcWord >> 11) & 0x1F;
        int epcBytes = epcWords * 2;

        if (epcBytes <= 0 || dataOffset + 2 + epcBytes > pkt.Length)
            return null;

        var epcData = new byte[epcBytes];
        Array.Copy(pkt, dataOffset + 2, epcData, 0, epcBytes);

        return new EpcReadResult
        {
            EpcBytes  = epcData,
            PcWord    = pcWord,
            ReadTime  = DateTimeOffset.UtcNow,
        };
    }

    // ── Ping / cancel helpers ──────────────────────────────────────────────────

    private Task PingAsync(CancellationToken ct) =>
        Task.Run(() =>
        {
            var port = _port!;
            byte[] ping = LlcsProtocol.BuildCommand(LlcsProtocol.Cmd.MacGetDebug);
            port.Write(ping, 0, ping.Length);
            // Drain ACK (best-effort; ignore result — device may not respond to MAC_GET_DEBUG exactly)
            try
            {
                var buf = new byte[LlcsProtocol.SmallPktLen];
                var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(2);
                TryFindSyncHeader(port, deadline, CancellationToken.None, out _);
            }
            catch { /* ignore ping failures */ }
        }, ct);

    private void SendCancel()
    {
        try
        {
            var port = _port!;
            byte[] cancel = LlcsProtocol.BuildCommand(LlcsProtocol.Cmd.ControlCancel);
            port.Write(cancel, 0, cancel.Length);
        }
        catch { /* swallow — best effort */ }
    }

    // ── Lifecycle ──────────────────────────────────────────────────────────────

    private void EnsureConnected()
    {
        if (!IsConnected)
            throw new InvalidOperationException("RFID reader is not connected. Call ConnectAsync first.");
    }

    private void ClosePort()
    {
        if (_port is { } sp)
        {
            try { if (sp.IsOpen) sp.Close(); } catch { /* ignore */ }
            sp.Dispose();
            _port = null;
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;
        await DisconnectAsync().ConfigureAwait(false);
        _lock.Dispose();
    }
}
