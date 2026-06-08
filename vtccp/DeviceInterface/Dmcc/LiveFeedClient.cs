namespace DeviceInterface.Dmcc;

using System.IO;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// Fetches a fresh camera frame per timer tick using two dedicated TCP
/// connections on port 23, operating completely independently of DMST.
///
/// ── Phase 1 — trigger (connection A) ─────────────────────────────────────
///   Enables extended ACK (COM.DMCC-RESPONSE=2), fires TRIGGER ON, then
///   drains everything the device sends back.  With extended ACK the device
///   delivers ||:::2[0]\r\n PLUS the full ~20 KB scan-result XML on that
///   connection.  Keeping all of this data on connection A (which is closed
///   before Phase 2) is the key invariant: no scan-result bytes can
///   contaminate the IMAGE.SEND connection.
///
/// ── Phase 2 — fetch (connection B) ───────────────────────────────────────
///   Opens a fresh connection with NO extended ACK, issues IMAGE.SEND, and
///   reads the binary JPEG response.  The stream is guaranteed clean.
///
/// TRIGGER.TYPE is never changed — stays at 0 throughout.
/// IMAGE.SIZE controls output resolution (DMST default = 1 → 1224×1024).
/// </summary>
public static class LiveFeedClient
{
    private const int ConnectTimeoutMs    = 2_000;
    private const int TriggerAckTimeout  = 3_500;   // > DECODER.TIMEOUT (2 000 ms)
    private const int IdleGapMs          = 200;

    /// <summary>
    /// Fires a software trigger on connection A (drains all result data there),
    /// then retrieves the frame on fresh connection B.
    /// Returns null on failure.
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

            // Phase 1 — trigger.  All scan-result XML stays on this connection.
            await FireTriggerAndWaitAsync(host, totalCts.Token);

            // Phase 2 — image.  Fresh connection, no XML contamination possible.
            return await FetchFrameAsync(host, totalCts.Token);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] GetLiveImageAsync: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    // ── Phase 1 ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Dedicated trigger connection.  Enables extended ACK, fires TRIGGER ON,
    /// drains the complete response (ACK + scan-result XML ≈ 20 KB), then
    /// disposes the connection.  Returns when the stream has been idle for
    /// IdleGapMs or the TriggerAckTimeout expires.
    /// </summary>
    private static async Task FireTriggerAndWaitAsync(
        string host, CancellationToken ct)
    {
        try
        {
            using var tcp = new TcpClient();
            using var connectCts =
                CancellationTokenSource.CreateLinkedTokenSource(ct);
            connectCts.CancelAfter(ConnectTimeoutMs);
            await tcp.ConnectAsync(host, DmccCommand.RawDmccPort, connectCts.Token);

            using var stream = tcp.GetStream();

            // Drain welcome banner.
            try
            {
                using var bc =
                    CancellationTokenSource.CreateLinkedTokenSource(ct);
                bc.CancelAfter(300);
                await stream.ReadAsync(new byte[512], bc.Token);
            }
            catch { }

            // Extended ACK: TRIGGER ON will push ||:::2[0]\r\n + full XML here.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}{DmccCommand.SetDmccResponseExtended}\r\n",
                300, ct);

            // Fire trigger (TRIGGER.TYPE stays 0 — never changed).
            await stream.WriteAsync(
                Encoding.ASCII.GetBytes(
                    $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n"), ct);

            // Drain ALL response bytes:
            //   • First byte arrives when scan completes (up to TriggerAckTimeout)
            //   • Remainder (~20 KB XML) follows within IdleGapMs
            byte[]? resp = await DrainTriggerResponseAsync(
                stream, TriggerAckTimeout, ct);

            if (resp is null)
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] Trigger: no response (timeout / no-read).");
            else
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-LIVEFEED] Trigger OK — drained {resp.Length} B.");

            // Connection A is disposed here — scan-result bytes gone with it.
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] FireTrigger: {ex.GetType().Name}: {ex.Message}");
        }
    }

    // ── Phase 2 ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Fresh IMAGE.SEND connection.  No extended ACK — response is pure binary
    /// JPEG (no text prefix).  Returns the JPEG bytes or null.
    /// </summary>
    private static async Task<byte[]?> FetchFrameAsync(
        string host, CancellationToken ct)
    {
        try
        {
            using var tcp = new TcpClient();
            using var connectCts =
                CancellationTokenSource.CreateLinkedTokenSource(ct);
            connectCts.CancelAfter(ConnectTimeoutMs);
            await tcp.ConnectAsync(host, DmccCommand.RawDmccPort, connectCts.Token);

            using var stream = tcp.GetStream();

            // Drain welcome banner.
            try
            {
                using var bc =
                    CancellationTokenSource.CreateLinkedTokenSource(ct);
                bc.CancelAfter(300);
                await stream.ReadAsync(new byte[512], bc.Token);
            }
            catch { }

            // Request the last acquired frame.
            // No extended ACK on this connection — response is binary only.
            await stream.WriteAsync(
                Encoding.ASCII.GetBytes("IMAGE.SEND\r\n"), ct);

            byte[]? raw = await ReadUntilIdleAsync(stream, ct);
            if (raw is null || raw.Length < 4)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] IMAGE.SEND: no data.");
                return null;
            }

            int start = FindJpegStart(raw);
            if (start < 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-LIVEFEED] IMAGE.SEND: no JPEG SOI in {raw.Length} B.");
                return null;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] Frame: {raw.Length - start} B JPEG.");
            return raw[start..];
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] FetchFrame: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Waits up to <paramref name="firstByteTimeoutMs"/> for the FIRST byte
    /// (signals scan completion), then drains remaining bytes using IdleGapMs.
    /// Returns null when no data arrives within the first-byte window.
    /// </summary>
    private static async Task<byte[]?> DrainTriggerResponseAsync(
        NetworkStream     stream,
        int               firstByteTimeoutMs,
        CancellationToken ct)
    {
        using var ms    = new MemoryStream(32 * 1024);
        byte[]    buf   = new byte[16 * 1024];
        bool      first = true;

        try
        {
            while (true)
            {
                int waitMs = first ? firstByteTimeoutMs : IdleGapMs;
                using var timedCts =
                    CancellationTokenSource.CreateLinkedTokenSource(ct);
                timedCts.CancelAfter(waitMs);

                int n;
                try   { n = await stream.ReadAsync(buf, timedCts.Token); }
                catch (OperationCanceledException) { break; }

                if (n <= 0) break;
                ms.Write(buf, 0, n);
                first = false;
            }
        }
        catch { }

        return ms.Length == 0 ? null : ms.ToArray();
    }

    /// <summary>
    /// Reads until the stream has been idle for IdleGapMs or ct fires.
    /// </summary>
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
            await stream.ReadAsync(new byte[64], drain.Token);
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
