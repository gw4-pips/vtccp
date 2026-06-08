namespace DeviceInterface.Dmcc;

using System.IO;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// Fetches a fresh camera frame per timer tick on a single TCP port-23
/// connection, operating completely independently of DMST.
///
/// Sequence:
///   1. Connect, drain welcome banner.
///   2. SET COM.DMCC-RESPONSE 0  — silent mode: TRIGGER ON fires a scan but
///      the device sends NOTHING back on this connection.  The scan result
///      travels only via the HTTP push channel (port 44444).  This is the
///      key invariant: the TCP stream stays clean for IMAGE.SEND.
///   3. TRIGGER ON — starts a TruCheck scan (TRIGGER.TYPE stays 0).
///   4. Wait <see cref="ScanWaitMs"/> for the scan to complete.
///      Typical scan time for a label in frame: 100–500 ms.
///      DECODER.TIMEOUT (worst-case no-read): 2 000 ms.
///      ScanWaitMs = 1 200 ms covers most reads; on a no-read the device
///      has already given up at 2 000 ms and IMAGE.SEND returns the
///      last good frame from the previous tick.
///   5. IMAGE.SEND — returns last acquired frame as
///      {byte_count}\r\n{binary JPEG}.
///   6. FindJpegStart strips the text preamble and returns the JPEG bytes.
///
/// TRIGGER.TYPE is never changed.
/// IMAGE.SIZE controls output resolution (default 1 = 1224×1024).
/// </summary>
public static class LiveFeedClient
{
    private const int ConnectTimeoutMs = 2_000;
    private const int ScanWaitMs       = 1_200;   // wait after TRIGGER ON
    private const int IdleGapMs        = 200;

    /// <summary>
    /// Returns a fresh JPEG frame, or null on failure.
    /// </summary>
    public static async Task<byte[]?> GetLiveImageAsync(
        string            host,
        int               totalTimeoutMs = 8_000,
        CancellationToken ct             = default)
    {
        try
        {
            using var totalCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            totalCts.CancelAfter(totalTimeoutMs);

            using var tcp = new TcpClient();
            using var connectCts =
                CancellationTokenSource.CreateLinkedTokenSource(totalCts.Token);
            connectCts.CancelAfter(ConnectTimeoutMs);
            await tcp.ConnectAsync(host, DmccCommand.RawDmccPort, connectCts.Token);

            using var stream = tcp.GetStream();

            // ── 1. Drain welcome banner ───────────────────────────────────────
            try
            {
                using var bc =
                    CancellationTokenSource.CreateLinkedTokenSource(totalCts.Token);
                bc.CancelAfter(400);
                await stream.ReadAsync(new byte[512], bc.Token);
            }
            catch { }

            // ── 2. Silent mode: no response from device on this connection ────
            //    Sends its own extended-ACK for this SET command (drain it),
            //    then all subsequent commands — including TRIGGER ON — are
            //    silent.  Scan result goes to HTTP push only.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}SET COM.DMCC-RESPONSE 0\r\n",
                400, totalCts.Token);

            // ── 3. Fire trigger ───────────────────────────────────────────────
            await stream.WriteAsync(
                Encoding.ASCII.GetBytes(
                    $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n"),
                totalCts.Token);

            System.Diagnostics.Debug.WriteLine(
                "[VTCCP-LIVEFEED] TRIGGER ON sent — waiting for scan.");

            // ── 4. Wait for scan completion ───────────────────────────────────
            //    Stream is silent; we wait a fixed beat so the device finishes
            //    the scan and has the JPEG ready before we request it.
            await Task.Delay(ScanWaitMs, totalCts.Token);

            // ── 5. Request image ──────────────────────────────────────────────
            await stream.WriteAsync(
                Encoding.ASCII.GetBytes("IMAGE.SEND\r\n"), totalCts.Token);

            byte[]? raw = await ReadUntilIdleAsync(stream, totalCts.Token);

            if (raw is null || raw.Length < 4)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] IMAGE.SEND: no data.");
                return null;
            }

            // ── Diagnostic: show first bytes so we can confirm the format ─────
            int diagLen = Math.Min(raw.Length, 32);
            var diagHex = BitConverter.ToString(raw, 0, diagLen);
            var diagAsc = Encoding.ASCII.GetString(
                raw.Take(diagLen).Select(b => b is >= 32 and <= 126 ? b : (byte)'.').ToArray());
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] IMAGE.SEND first {diagLen} B  hex: {diagHex}");
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] IMAGE.SEND first {diagLen} B  asc: {diagAsc}");

            // ── 6. Find JPEG SOI ──────────────────────────────────────────────
            int start = FindJpegStart(raw);
            if (start < 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-LIVEFEED] IMAGE.SEND: no JPEG SOI in {raw.Length} B.");
                return null;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] Frame: {raw.Length - start} B JPEG (SOI at +{start}).");
            return raw[start..];
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] GetLiveImageAsync: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static async Task<byte[]?> ReadUntilIdleAsync(
        NetworkStream stream, CancellationToken ct)
    {
        using var ms  = new MemoryStream(256 * 1024);
        byte[]    buf = new byte[16 * 1024];

        try
        {
            while (true)
            {
                using var idleCts =
                    CancellationTokenSource.CreateLinkedTokenSource(ct);
                idleCts.CancelAfter(IdleGapMs);

                int n;
                try   { n = await stream.ReadAsync(buf, idleCts.Token); }
                catch (OperationCanceledException) { break; }

                if (n <= 0) break;
                ms.Write(buf, 0, n);
            }
        }
        catch { }

        return ms.Length == 0 ? null : ms.ToArray();
    }

    private static async Task WriteAndDrainAsync(
        NetworkStream stream, string command, int drainMs, CancellationToken ct)
    {
        await stream.WriteAsync(Encoding.ASCII.GetBytes(command), ct);
        try
        {
            using var drain = CancellationTokenSource.CreateLinkedTokenSource(ct);
            drain.CancelAfter(drainMs);
            await stream.ReadAsync(new byte[256], drain.Token);
        }
        catch { }
    }

    private static int FindJpegStart(byte[] data)
    {
        for (int i = 0; i < data.Length - 1; i++)
            if (data[i] == 0xFF && data[i + 1] == 0xD8)
                return i;
        return -1;
    }
}
