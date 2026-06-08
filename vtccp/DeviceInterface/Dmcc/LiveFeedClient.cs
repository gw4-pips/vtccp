namespace DeviceInterface.Dmcc;

using System.IO;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// Fetches a fresh camera frame per timer tick on a single TCP port-23
/// connection, operating completely independently of DMST.
///
/// ── Device behaviour on port 23 ──────────────────────────────────────────
/// After TRIGGER ON the device sends TWO separate deliveries on the same
/// TCP connection:
///   A. Command ACK: controlled by COM.DMCC-RESPONSE (0=off, 1=short, 2=full)
///   B. Scan-result XML (~21 KB): ALWAYS pushed when scan completes,
///      independent of COM.DMCC-RESPONSE.
///
/// By setting COM.DMCC-RESPONSE=0, delivery A is suppressed.  The only
/// data that arrives is delivery B — the result XML — which is our
/// scan-done signal.  We drain it completely, then IMAGE.SEND has a
/// provably clean stream.
///
/// ── Per-tick sequence ─────────────────────────────────────────────────────
///   1. Connect, drain welcome banner.
///   2. SET COM.DMCC-RESPONSE 0 — suppress command ACK.
///   3. TRIGGER ON — fires TruCheck scan (TRIGGER.TYPE stays 0).
///   4. DrainScanResultAsync — waits up to <see cref="ScanResultTimeoutMs"/>
///      for the FIRST byte (scan-done; typical: 300–500 ms), then reads
///      the full XML (~21 KB) using a 300 ms idle gap.
///   5. IMAGE.SEND — stream is clean; returns the acquired frame as
///      {byte_count}\r\n{binary JPEG}.
///   6. FindJpegStart strips the text preamble and returns JPEG bytes.
/// </summary>
public static class LiveFeedClient
{
    private const int ConnectTimeoutMs    = 2_000;
    private const int ScanResultTimeoutMs = 3_000;   // > DECODER.TIMEOUT (2 000 ms)
    private const int ScanResultIdleGapMs = 300;     // gap between XML chunks
    private const int ImageIdleGapMs      = 200;

    /// <summary>
    /// Returns a fresh JPEG frame, or null on failure.
    /// </summary>
    public static async Task<byte[]?> GetLiveImageAsync(
        string            host,
        int               totalTimeoutMs = 10_000,
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

            // ── 2. Suppress command ACK ───────────────────────────────────────
            // Default connection mode is RESPONSE=2; SET returns one extended
            // ACK for itself (drained here), then TRIGGER ON fires silently.
            await WriteAndDrainAsync(stream,
                $"{DmccCommand.WireHeader}SET COM.DMCC-RESPONSE 0\r\n",
                400, totalCts.Token);

            // ── 3. Fire trigger ───────────────────────────────────────────────
            await stream.WriteAsync(
                Encoding.ASCII.GetBytes(
                    $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n"),
                totalCts.Token);

            // ── 4. Drain scan-result XML ──────────────────────────────────────
            // The device pushes ~21 KB of XML when the scan completes.
            // First byte arrives when scan is done (300–500 ms typical).
            // We MUST drain it all before IMAGE.SEND or ReadUntilIdle will
            // return XML instead of JPEG.
            byte[]? scanResult = await DrainScanResultAsync(
                stream, ScanResultTimeoutMs, ScanResultIdleGapMs, totalCts.Token);

            if (scanResult is null)
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] Scan result: timeout (no-read or scan failed).");
            else
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-LIVEFEED] Scan result drained: {scanResult.Length} B.");

            // ── 5. Request image — stream is now clean ────────────────────────
            await stream.WriteAsync(
                Encoding.ASCII.GetBytes("IMAGE.SEND\r\n"), totalCts.Token);

            byte[]? raw = await ReadUntilIdleAsync(stream, ImageIdleGapMs, totalCts.Token);
            if (raw is null || raw.Length < 4)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-LIVEFEED] IMAGE.SEND: no data.");
                return null;
            }

            // ── Diagnostic (first 32 bytes) ───────────────────────────────────
            int diagLen = Math.Min(raw.Length, 32);
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] IMAGE.SEND first {diagLen} B  " +
                $"hex: {BitConverter.ToString(raw, 0, diagLen)}");
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] IMAGE.SEND first {diagLen} B  " +
                $"asc: {Encoding.ASCII.GetString(raw.Take(diagLen).Select(b => b is >= 32 and <= 126 ? b : (byte)'.').ToArray())}");

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

    /// <summary>
    /// Waits up to <paramref name="firstByteTimeoutMs"/> for the first byte
    /// (scan-done signal), then drains remaining bytes with
    /// <paramref name="idleGapMs"/>.  Returns null on timeout.
    /// </summary>
    private static async Task<byte[]?> DrainScanResultAsync(
        NetworkStream     stream,
        int               firstByteTimeoutMs,
        int               idleGapMs,
        CancellationToken ct)
    {
        using var ms    = new MemoryStream(32 * 1024);
        byte[]    buf   = new byte[16 * 1024];
        bool      first = true;

        try
        {
            while (true)
            {
                int waitMs = first ? firstByteTimeoutMs : idleGapMs;
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

    private static async Task<byte[]?> ReadUntilIdleAsync(
        NetworkStream stream, int idleGapMs, CancellationToken ct)
    {
        using var ms  = new MemoryStream(256 * 1024);
        byte[]    buf = new byte[16 * 1024];

        try
        {
            while (true)
            {
                using var idleCts =
                    CancellationTokenSource.CreateLinkedTokenSource(ct);
                idleCts.CancelAfter(idleGapMs);

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
