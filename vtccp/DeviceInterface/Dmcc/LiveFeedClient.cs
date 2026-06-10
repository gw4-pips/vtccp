namespace DeviceInterface.Dmcc;

using System.IO;
using System.Net.Sockets;
using System.Text;

/// <summary>
/// Fetches a fresh camera frame per timer tick using two sequential TCP
/// connections on port 23, operating completely independently of DMST.
///
/// ── Device behaviour on port 23 ──────────────────────────────────────────
/// After TRIGGER ON the device sends TWO separate deliveries:
///   A. Command ACK (controlled by COM.DMCC-RESPONSE — we suppress with 0).
///   B. Scan-result XML (~21 KB) — ALWAYS pushed when scan completes,
///      independent of COM.DMCC-RESPONSE.  After delivering the XML the
///      device closes / deactivates the session: IMAGE.SEND on the same
///      connection returns no data.
///
/// ── Per-tick sequence ────────────────────────────────────────────────────
///   Connection A — trigger + wait for scan-done:
///     1. Connect, drain welcome banner.
///     2. SET COM.DMCC-RESPONSE 0  — suppress command ACK.
///     3. TRIGGER ON               — fires TruCheck scan (TRIGGER.TYPE=0 unchanged).
///     4. DrainScanResultAsync     — waits up to <see cref="ScanResultTimeoutMs"/>
///        for the result XML (= scan-done signal), drains ~21 KB.
///     5. Close connection A.
///
///   Connection B — image fetch on a clean session:
///     6. Connect, drain welcome banner.
///     7. IMAGE.SEND               — returns last acquired frame as
///                                   {byte_count}\r\n{binary JPEG}.
///     8. FindJpegStart strips text preamble; return JPEG bytes.
///
///   Connection B is opened AFTER the scan result is confirmed received on A,
///   so no pending XML can contaminate the stream.
/// </summary>
public static class LiveFeedClient
{
    private const int ConnectTimeoutMs    = 2_000;
    private const int ScanResultTimeoutMs = 3_000;   // > DECODER.TIMEOUT (2 000 ms)
    private const int ScanResultIdleGapMs = 300;
    private const int ImageIdleGapMs      = 200;

    /// <summary>Returns a fresh JPEG frame, or null on failure.</summary>
    public static async Task<byte[]?> GetLiveImageAsync(
        string            host,
        int               totalTimeoutMs = 10_000,
        CancellationToken ct             = default)
    {
        try
        {
            using var totalCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            totalCts.CancelAfter(totalTimeoutMs);

            // ════════════════════════════════════════════════════════════════
            // CONNECTION A — trigger + drain scan-result XML (scan-done signal)
            // ════════════════════════════════════════════════════════════════
            {
                using var tcpA  = new TcpClient();
                using var conA  = CancellationTokenSource.CreateLinkedTokenSource(totalCts.Token);
                conA.CancelAfter(ConnectTimeoutMs);
                await tcpA.ConnectAsync(host, DmccCommand.RawDmccPort, conA.Token);

                using var streamA = tcpA.GetStream();

                // 1. Banner
                await DrainBannerAsync(streamA, totalCts.Token);

                // 2. Suppress command ACK (result XML still arrives regardless)
                await WriteAndDrainAsync(
                    streamA,
                    $"{DmccCommand.WireHeader}SET COM.DMCC-RESPONSE 0\r\n",
                    400, totalCts.Token);

                // 3. Trigger
                await streamA.WriteAsync(
                    Encoding.ASCII.GetBytes(
                        $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n"),
                    totalCts.Token);

                // 4. Drain scan-result XML — first byte = scan-done
                byte[]? scanResult = await DrainScanResultAsync(
                    streamA, ScanResultTimeoutMs, ScanResultIdleGapMs, totalCts.Token);

                if (scanResult is null)
                    System.Diagnostics.Debug.WriteLine(
                        "[VTCCP-LIVEFEED] Scan result: timeout — device may not have scanned.");
                else
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-LIVEFEED] Scan result drained: {scanResult.Length} B.");

                // 5. Connection A closes here — device session ends cleanly.
            }

            // ════════════════════════════════════════════════════════════════
            // CONNECTION B — IMAGE.SEND on a fresh, clean session
            // ════════════════════════════════════════════════════════════════
            {
                using var tcpB  = new TcpClient();
                using var conB  = CancellationTokenSource.CreateLinkedTokenSource(totalCts.Token);
                conB.CancelAfter(ConnectTimeoutMs);
                await tcpB.ConnectAsync(host, DmccCommand.RawDmccPort, conB.Token);

                using var streamB = tcpB.GetStream();

                // 6. Banner
                await DrainBannerAsync(streamB, totalCts.Token);

                // 7. Switch to extended ACK mode so IMAGE.SEND returns data.
                //    Default COM.DMCC-RESPONSE on a fresh connection is 0 (silent).
                //    The SET command itself has no response (was in mode 0 when sent);
                //    no drain needed.
                await streamB.WriteAsync(
                    Encoding.ASCII.GetBytes(
                        $"{DmccCommand.WireHeader}SET COM.DMCC-RESPONSE 2\r\n"),
                    totalCts.Token);

                // 8. Request image — MUST include the "||>" wire header.
                //    Bare "IMAGE.SEND\r\n" is silently ignored by the device.
                await streamB.WriteAsync(
                    Encoding.ASCII.GetBytes(
                        $"{DmccCommand.WireHeader}{DmccCommand.ImageSend}\r\n"),
                    totalCts.Token);

                byte[]? raw = await ReadUntilIdleAsync(streamB, ImageIdleGapMs, totalCts.Token);
                if (raw is null || raw.Length < 4)
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[VTCCP-LIVEFEED] IMAGE.SEND: no data.");
                    return null;
                }

                // Diagnostic: show first 32 bytes in hex + ASCII
                int diagLen = Math.Min(raw.Length, 32);
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-LIVEFEED] IMAGE.SEND first {diagLen} B  " +
                    $"hex: {BitConverter.ToString(raw, 0, diagLen)}");
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-LIVEFEED] IMAGE.SEND first {diagLen} B  " +
                    $"asc: {Encoding.ASCII.GetString(raw.Take(diagLen).Select(b => b is >= 32 and <= 126 ? b : (byte)'.').ToArray())}");

                // 8. Find JPEG SOI (strips any text preamble / DMCC ACK header)
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
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] GetLiveImageAsync: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Fires a software trigger and returns the resulting fresh sensor frame.
    /// Use for live-view polling — every frame comes from a real acquisition.
    ///
    /// Faster than GetLiveImageAsync:
    ///   • No WriteAndDrainAsync for COM.DMCC-RESPONSE — the device defaults to
    ///     mode 0 on a fresh connection; the SET has no response to drain.
    ///   • Drains only the FIRST BYTE of the scan-result XML (scan-done signal)
    ///     rather than the full ~21 KB payload.
    ///   • Connection B (IMAGE.SEND) is opened after scan-done, not serialised
    ///     behind the full XML drain.
    ///
    /// Estimated round-trip: ~650 ms (Grade F) to ~400 ms (Grade A).
    /// The _fetchInProgress guard in LiveFeedViewModel prevents concurrent calls.
    /// </summary>
    public static async Task<byte[]?> GetFreshFrameAsync(
        string            host,
        int               totalTimeoutMs = 4_000,
        CancellationToken ct             = default)
    {
        try
        {
            using var total = CancellationTokenSource.CreateLinkedTokenSource(ct);
            total.CancelAfter(totalTimeoutMs);

            // ════════════════════════════════════════════════════════════════
            // CONNECTION A — trigger + first-byte drain (scan-done signal)
            // ════════════════════════════════════════════════════════════════
            {
                using var tcpA = new TcpClient();
                using var conA = CancellationTokenSource.CreateLinkedTokenSource(total.Token);
                conA.CancelAfter(ConnectTimeoutMs);
                await tcpA.ConnectAsync(host, DmccCommand.RawDmccPort, conA.Token);
                using var streamA = tcpA.GetStream();

                await DrainBannerAsync(streamA, total.Token);

                // Default COM.DMCC-RESPONSE on a fresh connection is already 0
                // (silent — no command ACK).  Skip the SET + drain that costs
                // 400 ms in GetLiveImageAsync.

                // Fire trigger — result XML will arrive on this stream when done.
                await streamA.WriteAsync(
                    Encoding.ASCII.GetBytes(
                        $"{DmccCommand.WireHeader}{DmccCommand.TriggerOn}\r\n"),
                    total.Token);

                // Wait for first byte of result XML = scan-done signal.
                // Not draining the full payload — that goes to the push listener.
                try
                {
                    using var scanDone = CancellationTokenSource
                        .CreateLinkedTokenSource(total.Token);
                    scanDone.CancelAfter(ScanResultTimeoutMs);
                    await streamA.ReadAsync(new byte[1], scanDone.Token);
                }
                catch { /* timeout — proceed; IMAGE.SEND may return previous frame */ }

                // Connection A closes here; device discards remaining XML on its end.
            }

            // ════════════════════════════════════════════════════════════════
            // CONNECTION B — IMAGE.SEND on a clean session after scan is done
            // ════════════════════════════════════════════════════════════════
            {
                using var tcpB = new TcpClient();
                using var conB = CancellationTokenSource.CreateLinkedTokenSource(total.Token);
                conB.CancelAfter(ConnectTimeoutMs);
                await tcpB.ConnectAsync(host, DmccCommand.RawDmccPort, conB.Token);
                using var streamB = tcpB.GetStream();

                await DrainBannerAsync(streamB, total.Token);

                // Switch to extended mode — IMAGE.SEND requires it to return data.
                await streamB.WriteAsync(
                    Encoding.ASCII.GetBytes(
                        $"{DmccCommand.WireHeader}SET COM.DMCC-RESPONSE 2\r\n"),
                    total.Token);

                await streamB.WriteAsync(
                    Encoding.ASCII.GetBytes(
                        $"{DmccCommand.WireHeader}{DmccCommand.ImageSend}\r\n"),
                    total.Token);

                byte[]? raw = await ReadUntilIdleAsync(streamB, 100, total.Token);
                if (raw is null || raw.Length < 4)
                    return null;

                int start = FindJpegStart(raw);
                if (start < 0)
                    return null;

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-LIVEFEED] GetFreshFrameAsync: {raw.Length - start} B JPEG.");
                return raw[start..];
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] GetFreshFrameAsync: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Fetches the latest buffered frame via IMAGE.SEND with no trigger.
    /// Returns a stale frame (whatever the device last acquired) — use only
    /// when the calling code has already triggered the device separately.
    /// </summary>
    public static async Task<byte[]?> GetFrameAsync(
        string            host,
        int               totalTimeoutMs = 2_000,
        CancellationToken ct             = default)
    {
        try
        {
            using var total = CancellationTokenSource.CreateLinkedTokenSource(ct);
            total.CancelAfter(totalTimeoutMs);

            using var tcp = new TcpClient();
            using var con = CancellationTokenSource.CreateLinkedTokenSource(total.Token);
            con.CancelAfter(ConnectTimeoutMs);
            await tcp.ConnectAsync(host, DmccCommand.RawDmccPort, con.Token);

            using var stream = tcp.GetStream();
            await DrainBannerAsync(stream, total.Token);

            await stream.WriteAsync(
                Encoding.ASCII.GetBytes(
                    $"{DmccCommand.WireHeader}SET COM.DMCC-RESPONSE 2\r\n"),
                total.Token);

            await stream.WriteAsync(
                Encoding.ASCII.GetBytes(
                    $"{DmccCommand.WireHeader}{DmccCommand.ImageSend}\r\n"),
                total.Token);

            byte[]? raw = await ReadUntilIdleAsync(stream, 100, total.Token);
            if (raw is null || raw.Length < 4)
                return null;

            int start = FindJpegStart(raw);
            return start >= 0 ? raw[start..] : null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-LIVEFEED] GetFrameAsync: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static async Task DrainBannerAsync(NetworkStream stream, CancellationToken ct)
    {
        try
        {
            using var bc = CancellationTokenSource.CreateLinkedTokenSource(ct);
            bc.CancelAfter(400);
            await stream.ReadAsync(new byte[512], bc.Token);
        }
        catch { }
    }

    /// <summary>
    /// Waits up to <paramref name="firstByteTimeoutMs"/> for the first byte
    /// (scan-done signal), then drains the rest with
    /// <paramref name="idleGapMs"/> idle gap.  Returns null on timeout.
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
                using var timed = CancellationTokenSource.CreateLinkedTokenSource(ct);
                timed.CancelAfter(waitMs);

                int n;
                try   { n = await stream.ReadAsync(buf, timed.Token); }
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
                using var idle = CancellationTokenSource.CreateLinkedTokenSource(ct);
                idle.CancelAfter(idleGapMs);

                int n;
                try   { n = await stream.ReadAsync(buf, idle.Token); }
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
