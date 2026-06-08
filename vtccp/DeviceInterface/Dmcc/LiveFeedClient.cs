namespace DeviceInterface.Dmcc;

using System.IO;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// Fetches live camera images from the device via raw TCP IMAGE.SEND on port 23.
/// Each call opens a fresh TCP connection, issues the command, reads the binary
/// JPEG response, strips any DMCC text header by finding the JPEG SOI marker, and
/// closes the connection.
///
/// Designed for 3-fps polling (333 ms interval): a fresh connection per frame is
/// acceptable at this rate because the device readily accepts short-lived port-23
/// sessions.
///
/// IMAGE.SEND is sent without the ||&gt; WireHeader prefix — binary commands on
/// port 23 do not require framing, and the device returns the JPEG payload directly
/// (any DMCC text preamble is stripped by locating 0xFF 0xD8).
/// </summary>
public static class LiveFeedClient
{
    private const int ConnectTimeoutMs = 2_000;
    private const int IdleGapMs        = 300;

    /// <summary>
    /// Connects to <paramref name="host"/> on port 23, sends <c>IMAGE.SEND\r\n</c>,
    /// reads the binary response until the idle gap expires, then returns the JPEG
    /// payload (everything from the first 0xFF 0xD8 SOI marker onwards).
    ///
    /// Returns null if the connection fails, the device returns no data, or no valid
    /// JPEG SOI marker is found in the response.
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

            // Drain welcome banner (~100 B on first port-23 connect).
            try
            {
                using var bannerCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                bannerCts.CancelAfter(300);
                await stream.ReadAsync(new byte[512], bannerCts.Token);
            }
            catch (OperationCanceledException) { }
            catch { }

            // Send IMAGE.SEND without WireHeader prefix (binary command — no text ACK follows).
            await stream.WriteAsync(Encoding.ASCII.GetBytes("IMAGE.SEND\r\n"), ct);

            // Read all bytes until idle gap or total timeout.
            byte[]? raw = await ReadUntilIdleAsync(stream, totalTimeoutMs, ct);
            if (raw is null || raw.Length < 4) return null;

            // Locate the JPEG SOI marker (0xFF 0xD8) — strips any DMCC text preamble.
            int start = FindJpegStart(raw);
            return start >= 0 ? raw[start..] : null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] IMAGE.SEND error: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static async Task<byte[]?> ReadUntilIdleAsync(
        NetworkStream     stream,
        int               totalTimeoutMs,
        CancellationToken ct)
    {
        using var ms  = new MemoryStream(64 * 1024);
        byte[]    buf = new byte[16 * 1024];

        using var totalCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        totalCts.CancelAfter(totalTimeoutMs);

        try
        {
            while (true)
            {
                // Inner idle-gap cancellation — breaks out when the stream goes quiet.
                // Outer totalCts fires when the overall budget is exhausted.
                using var idleCts =
                    CancellationTokenSource.CreateLinkedTokenSource(totalCts.Token);
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
