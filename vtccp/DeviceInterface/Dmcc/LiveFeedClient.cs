namespace DeviceInterface.Dmcc;

using System.IO;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// Fetches the most recently acquired camera frame from the device via a
/// raw IMAGE.SEND command on TCP port 23.
///
/// Each call opens one TCP session that:
///   1. Drains the welcome banner.
///   2. Sends IMAGE.SEND — device returns "last camera image acquired" as
///      {byte_count}\r\n{binary JPEG data}.
///   3. Strips any DMCC text preamble by locating the JPEG SOI marker
///      (0xFF 0xD8).
///
/// NO TRIGGER ON is fired here.  The device is kept triggered by the
/// background monitoring scan loop that DMST maintains while Go Live is
/// active; IMAGE.SEND simply pulls the latest frame from that buffer.
/// Software-trigger for TruCheck verification is handled separately by
/// LiveFeedViewModel.SendTriggerAsync.
///
/// IMAGE.SIZE controls the output resolution:
///   0 = Full (2448×2048)   1 = 1/4 (1224×1024)   2 = 1/16   3 = 1/64
/// At the DMST default IMAGE.SIZE=1, IMAGE.SEND delivers 1224×1024 JPEG.
/// </summary>
public static class LiveFeedClient
{
    private const int ConnectTimeoutMs = 2_000;
    private const int IdleGapMs        = 200;   // stream-idle gap that ends the read loop

    /// <summary>
    /// Returns the last camera frame as a JPEG byte array, or null if the
    /// connection fails, the device returns no data, or no valid JPEG SOI
    /// marker is found.
    /// </summary>
    public static async Task<byte[]?> GetLiveImageAsync(
        string            host,
        int               totalTimeoutMs = 2_500,
        CancellationToken ct             = default)
    {
        try
        {
            using var connectCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            connectCts.CancelAfter(ConnectTimeoutMs);

            using var tcp = new TcpClient();
            await tcp.ConnectAsync(host, DmccCommand.RawDmccPort, connectCts.Token);
            using var stream = tcp.GetStream();

            // Drain welcome banner — typically < 10 ms on LAN.
            try
            {
                using var bannerCts =
                    CancellationTokenSource.CreateLinkedTokenSource(ct);
                bannerCts.CancelAfter(300);
                await stream.ReadAsync(new byte[512], bannerCts.Token);
            }
            catch { }

            // Request the last acquired frame.
            // Sent without WireHeader prefix — binary command, no text ACK follows.
            await stream.WriteAsync(Encoding.ASCII.GetBytes("IMAGE.SEND\r\n"), ct);

            // Read until the stream goes idle.
            using var totalCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            totalCts.CancelAfter(totalTimeoutMs);

            byte[]? raw = await ReadUntilIdleAsync(stream, totalCts.Token);
            if (raw is null || raw.Length < 4)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] IMAGE.SEND returned no data.");
                return null;
            }

            // FindJpegStart strips the DMCC text preamble ({byte_count}\r\n)
            // by scanning for the JPEG SOI marker.
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
