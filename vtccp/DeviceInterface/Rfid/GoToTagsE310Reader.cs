using System.IO.Ports;
using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid;

/// <summary>
/// RFID reader implementation for the GoToTags Desktop E310 UHF RFID Reader
/// (Impinj E310 chipset, Spokane WA, USA).
///
/// Interface: FTDI USB VCP → virtual COM port at 115 200 8N1 (factory default).
///   Windows driver: CDM212364_Setup.zip (gitlab.com/gototags/public →
///   UHF RFID/Readers/GoToTags/drivers/windows/).
///
/// Protocol: GoToTags UHF RFID Reader Communication Protocol rev 5-30-23.
///   "Common Command Communication Protocol" — request/reply binary framing.
///   Frame: [ 0xFF | DataLen | CmdCode | Data[DataLen] | CRC_Hi | CRC_Lo ]
///   CRC-16/CCITT (init=0xFFFF, poly=0x1021) over bytes[1..frameLen-3].
///
/// TriggerInventoryAsync strategy: issue Single Tag Inventory (0x21) commands
/// in a tight loop for the requested duration, de-duplicating EPCs by hex string.
/// Each 0x21 scan uses a short polling window so the loop stays responsive to
/// the overall timeout and the CancellationToken.
///
/// No native GoToTags DLLs required — pure managed SerialPort I/O.
/// </summary>
public sealed class GoToTagsE310Reader : IEpcReader
{
    private SerialPort? _port;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private bool _disposed;

    // ── Tuning constants ─────────────────────────────────────────────────────

    /// <summary>
    /// Duration of each individual 0x21 scan within a TriggerInventoryAsync call.
    /// Shorter = more responsive to timeout / cancellation; longer = more efficient.
    /// </summary>
    private const int ScanSliceMs = 150;

    /// <summary>
    /// Serial read timeout — must exceed ScanSliceMs so we always get the full
    /// response from a single 0x21 command (or timeout cleanly).
    /// </summary>
    private const int SerialReadTimeoutMs = ScanSliceMs + 500;

    // ── IEpcReader ───────────────────────────────────────────────────────────

    public bool IsConnected => _port?.IsOpen == true;

    public async Task ConnectAsync(string portName, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (IsConnected)
                throw new InvalidOperationException(
                    "Reader is already connected. Call DisconnectAsync first.");

            var sp = new SerialPort(portName, 115200, Parity.None, 8, StopBits.One)
            {
                Handshake   = Handshake.None,
                ReadTimeout  = SerialReadTimeoutMs,
                WriteTimeout = 3000,
                DtrEnable    = false,
                RtsEnable    = false,
            };

            sp.Open();
            sp.DiscardInBuffer();
            sp.DiscardOutBuffer();
            _port = sp;

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
            return await Task.Run(() => RunInventoryLoop(timeout, ct), ct)
                             .ConfigureAwait(false);
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task CancelAsync(CancellationToken ct = default)
    {
        // 0x21 is synchronous — the reader stops after each slice.
        // Flushing the serial buffers is sufficient to break any in-progress read.
        await _lock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (!IsConnected) return;
            _port!.DiscardInBuffer();
            _port!.DiscardOutBuffer();
        }
        finally
        {
            _lock.Release();
        }
    }

    // ── Inventory loop ───────────────────────────────────────────────────────

    private IReadOnlyList<EpcReadResult> RunInventoryLoop(
        TimeSpan timeout, CancellationToken ct)
    {
        var results  = new List<EpcReadResult>();
        var seenEpcs = new HashSet<string>(StringComparer.Ordinal);
        var deadline = DateTimeOffset.UtcNow + timeout;
        var port     = _port!;

        while (DateTimeOffset.UtcNow < deadline && !ct.IsCancellationRequested)
        {
            // Clamp the slice so we don't overshoot the deadline.
            var remaining = (int)(deadline - DateTimeOffset.UtcNow).TotalMilliseconds;
            if (remaining <= 0) break;
            int sliceMs = Math.Min(ScanSliceMs, remaining);

            byte[] cmd = GoToTagsE310Protocol.BuildSingleInventoryCmd(sliceMs);
            port.DiscardInBuffer();
            port.Write(cmd, 0, cmd.Length);

            byte[]? frame = TryReadResponseFrame(port, sliceMs + 600);
            if (frame is null) continue;

            byte[]? epcBytes = GoToTagsE310Protocol.TryExtractEpc(frame);
            if (epcBytes is null) continue;

            string epcHex = Convert.ToHexString(epcBytes);
            if (!seenEpcs.Add(epcHex)) continue;

            // PC word is not returned in no-metadata mode; derive from EPC length.
            // Per GS1 Gen2: PC[15:11] = EPC length in 16-bit words.
            int epcWords = epcBytes.Length / 2;
            ushort pcWord = (ushort)(epcWords << 11);

            results.Add(new EpcReadResult
            {
                EpcBytes = epcBytes,
                PcWord   = pcWord,
                ReadTime = DateTimeOffset.UtcNow,
            });
        }

        return results;
    }

    // ── Serial frame reader ──────────────────────────────────────────────────

    /// <summary>
    /// Read one complete response frame from the serial port.
    /// Scans for the 0xFF header byte, then reads the rest of the frame.
    /// Returns null on timeout or malformed data.
    /// </summary>
    private static byte[]? TryReadResponseFrame(SerialPort port, int windowMs)
    {
        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromMilliseconds(windowMs);

        // Scan for 0xFF header.
        byte header;
        while (true)
        {
            if (DateTimeOffset.UtcNow >= deadline) return null;
            int b;
            try { b = port.ReadByte(); }
            catch (TimeoutException) { return null; }
            if (b == GoToTagsE310Protocol.Header) { header = (byte)b; break; }
        }

        // Read DataLen byte.
        if (DateTimeOffset.UtcNow >= deadline) return null;
        int dataLen;
        try { dataLen = port.ReadByte(); }
        catch (TimeoutException) { return null; }
        if (dataLen < 0) return null;

        // Full frame: FF(1) + DataLen(1) + CmdCode(1) + Data(dataLen) + CRC(2)
        int frameSize = GoToTagsE310Protocol.FrameSize(dataLen);
        var frame = new byte[frameSize];
        frame[0] = header;
        frame[1] = (byte)dataLen;

        // Read the rest: CmdCode + Data + CRC.
        int remaining = frameSize - 2;
        int offset    = 2;
        while (remaining > 0)
        {
            if (DateTimeOffset.UtcNow >= deadline) return null;
            int read;
            try
            {
                read = port.Read(frame, offset, remaining);
            }
            catch (TimeoutException) { return null; }
            offset    += read;
            remaining -= read;
        }

        // Verify CRC.
        if (!GoToTagsE310Protocol.VerifyCrc(frame, frameSize)) return null;

        return frame;
    }

    // ── Connectivity probe ───────────────────────────────────────────────────

    /// <summary>
    /// Send a Single Tag Inventory with a very short timeout as a connectivity probe.
    /// Any response (tag found OR no-tag status) confirms the reader is alive in APP state.
    /// </summary>
    private Task PingAsync(CancellationToken ct) =>
        Task.Run(() =>
        {
            var port = _port!;
            byte[] cmd = GoToTagsE310Protocol.BuildSingleInventoryCmd(
                GoToTagsE310Protocol.PingTimeoutMs);

            port.DiscardInBuffer();
            port.Write(cmd, 0, cmd.Length);

            // Any parseable frame = reader is alive. Ignore whether a tag was found.
            try { TryReadResponseFrame(port, GoToTagsE310Protocol.PingTimeoutMs + 600); }
            catch { /* ignore — port was just opened; a non-response is acceptable */ }
        }, ct);

    // ── Lifecycle ────────────────────────────────────────────────────────────

    private void EnsureConnected()
    {
        if (!IsConnected)
            throw new InvalidOperationException(
                "RFID reader is not connected. Call ConnectAsync first.");
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
