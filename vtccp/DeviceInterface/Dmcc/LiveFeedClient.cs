namespace DeviceInterface.Dmcc;

using System.IO;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// Fetches live camera images from the device via raw TCP on port 23.
///
/// Each call opens one TCP session that:
///   1. Drains the welcome banner.
///   2. Enables extended ACK mode (COM.DMCC-RESPONSE 2) so TRIGGER ON
///      returns a text ACK.
///   3. Sends TRIGGER ON — causes the device to acquire a new frame.
///      TRIGGER.TYPE is left at 0 (Single/External) throughout; the
///      polling loop is the only thing that changes the acquisition rate.
///   4. Waits <see cref="AcquireWaitMs"/> ms for the acquisition to complete.
///   5. Sends IMAGE.SEND — retrieves the full 2448×2048 sensor frame as a
///      JPEG (IMAGE.FORMAT=1, IMAGE.QUALITY=50 at clean-state).
///      IMAGE.SIZE does not downscale IMAGE.SEND output — device-confirmed.
///   6. Strips any DMCC text preamble by locating the JPEG SOI marker.
///
/// Matches DMST Go Live behaviour: TRIGGER ON every ~400 ms, no
/// TRIGGER.TYPE or LIVEIMG.MODE changes, pull via IMAGE.SEND.
/// </summary>
public static class LiveFeedClient
{
    private const int ConnectTimeoutMs = 2_000;
    private const int AcquireWaitMs    = 150;   // wait after TRIGGER ON for sensor readout
    private const int IdleGapMs        = 200;   // stream-idle gap that terminates the read loop

    /// <summary>
    /// Fires a software trigger and retrieves the resulting image frame.
    /// Returns null if the connection fails, the device returns no data,
    /// or no valid JPEG SOI marker is found in the response.
    /// </summary>
    public static async Task<byte[]?> GetLiveImageAsync(
        string            host,
        int               totalTimeoutMs = 3_000,
        CancellationToken ct             = default)
    {
        try
        {
            using var totalCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            totalCts.CancelAfter(totalTimeoutMs);

            using var tcp = new TcpClient();
            await tcp.ConnectAsync(host, DmccCommand.RawDmccPort, totalCts.Token);
            using var stream = tcp.GetStream();

            // Drain welcome banner — typically < 10 ms on LAN.
            try
            {
                using var bannerCts =
                    CancellationTokenSource.CreateLinkedTokenSource(totalCts.Token);
                bannerCts.CancelAfter(300);
                await stream.ReadAsync(new byte[512], bannerCts.Token);
            }
            catch { }

            // Enable extended ACK so TRIGGER ON returns ||:::2[0]\r\n.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}{DmccCommand.SetDmccResponseExtended}\r\n",
                300, totalCts.Token);

            // Fire software trigger — device opens shutter and captures a frame.
            // TRIGGER.TYPE remains 0 (Single/External) — no mode change needed.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n",
                500, totalCts.Token);

            // Give the sensor time to finish readout before requesting the image.
            await Task.Delay(AcquireWaitMs, totalCts.Token);

            // IMAGE.SEND — returns {byte_count}\r\n{binary JPEG}.
            // Sent without WireHeader prefix (binary response, no text ACK).
            await stream.WriteAsync(
                Encoding.ASCII.GetBytes("IMAGE.SEND\r\n"), totalCts.Token);

            // Read until the stream goes idle.
            byte[]? raw = await ReadUntilIdleAsync(stream, totalCts.Token);
            if (raw is null || raw.Length < 4) return null;

            // FindJpegStart strips any DMCC text preamble before the SOI marker.
            int start = FindJpegStart(raw);
            return start >= 0 ? raw[start..] : null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] GetLiveImageAsync error: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Writes an ASCII command then reads (and discards) the response within
    /// <paramref name="drainMs"/> ms.  Swallows all errors — a missed ACK
    /// must never block the caller.
    /// </summary>
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
