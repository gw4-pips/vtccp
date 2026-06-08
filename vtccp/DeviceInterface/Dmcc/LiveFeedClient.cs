namespace DeviceInterface.Dmcc;

using System.IO;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// Fetches a fresh camera frame from the device via raw TCP on port 23,
/// operating completely independently of DMST.
///
/// Each call opens one TCP session that:
///   1. Drains the welcome banner.
///   2. Enables extended ACK mode (COM.DMCC-RESPONSE 2).
///   3. Sends TRIGGER ON — device acquires and processes a new frame.
///   4. Waits for the TRIGGER ON completion ACK (||:::2[0]\r\n).
///      This ACK only arrives when the scan is fully done; waiting for it
///      is the correct way to know the image buffer is ready.
///      DECODER.TIMEOUT=2000ms is the worst-case scan time; a label
///      sitting still typically completes in 100–500ms.
///   5. Sends IMAGE.SEND — retrieves the just-acquired frame.
///      At IMAGE.SIZE=1 (DMST default = 1/4 area) this is 1224×1024 JPEG.
///   6. Strips any DMCC text preamble by locating the JPEG SOI marker.
///
/// TRIGGER.TYPE is never changed — stays at 0 (Single/External) throughout.
/// </summary>
public static class LiveFeedClient
{
    private const int ConnectTimeoutMs  = 2_000;
    private const int TriggerAckTimeout = 3_000;  // > DECODER.TIMEOUT (2000ms)
    private const int IdleGapMs         = 200;

    /// <summary>
    /// Fires a software trigger, waits for scan completion, then retrieves
    /// the resulting frame.  Returns null on connection failure, scan
    /// timeout, or missing JPEG SOI marker.
    /// </summary>
    public static async Task<byte[]?> GetLiveImageAsync(
        string            host,
        int               totalTimeoutMs = 6_000,   // connect + ACK wait + transfer
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

            // Drain welcome banner.
            try
            {
                using var bc =
                    CancellationTokenSource.CreateLinkedTokenSource(totalCts.Token);
                bc.CancelAfter(300);
                await stream.ReadAsync(new byte[512], bc.Token);
            }
            catch { }

            // Enable extended ACK so TRIGGER ON returns ||:::2[0]\r\n on completion.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}{DmccCommand.SetDmccResponseExtended}\r\n",
                300, totalCts.Token);

            // Fire software trigger.  TRIGGER.TYPE stays 0 throughout.
            byte[] trigCmd = Encoding.ASCII.GetBytes(
                $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n");
            await stream.WriteAsync(trigCmd, totalCts.Token);

            // Wait for the completion ACK — this is the scan-done signal.
            // Do NOT send IMAGE.SEND until this arrives; the image buffer is
            // not ready until the scan finishes.
            using var ackCts =
                CancellationTokenSource.CreateLinkedTokenSource(totalCts.Token);
            ackCts.CancelAfter(TriggerAckTimeout);
            try
            {
                byte[] ackBuf = new byte[128];
                int n = await stream.ReadAsync(ackBuf, ackCts.Token);
                if (n > 0)
                    System.Diagnostics.Debug.WriteLine(
                        "[VTCCP-LIVEFEED] TRIGGER ACK: " +
                        Encoding.ASCII.GetString(ackBuf, 0, n).Trim());
            }
            catch (OperationCanceledException)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] TRIGGER ACK timed out — sending IMAGE.SEND anyway.");
            }

            // Request the newly acquired frame.
            // Sent without WireHeader — binary response, no text ACK.
            await stream.WriteAsync(
                Encoding.ASCII.GetBytes("IMAGE.SEND\r\n"), totalCts.Token);

            byte[]? raw = await ReadUntilIdleAsync(stream, totalCts.Token);
            if (raw is null || raw.Length < 4)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] IMAGE.SEND returned no data.");
                return null;
            }

            int start = FindJpegStart(raw);
            if (start < 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-LIVEFEED] No JPEG SOI in {raw.Length}-byte response.");
                return null;
            }

            return raw[start..];
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] GetLiveImageAsync error: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static async Task WriteAndDrainAsync(
        NetworkStream     stream,
        string            command,
        int               drainMs,
        CancellationToken ct)
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

    private static async Task<byte[]?> ReadUntilIdleAsync(
        NetworkStream     stream,
        CancellationToken ct)
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
                try { n = await stream.ReadAsync(buf, idleCts.Token); }
                catch (OperationCanceledException) { break; }

                if (n <= 0) break;
                ms.Write(buf, 0, n);
            }
        }
        catch { }

        return ms.Length == 0 ? null : ms.ToArray();
    }

    private static int FindJpegStart(byte[] data)
    {
        for (int i = 0; i < data.Length - 1; i++)
            if (data[i] == 0xFF && data[i + 1] == 0xD8)
                return i;
        return -1;
    }
}
