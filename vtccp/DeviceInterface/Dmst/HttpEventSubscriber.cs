namespace DeviceInterface.Dmst;

using System.Net.Sockets;
using System.Text;
using System.Xml.Linq;
using ExcelEngine.Models;

/// <summary>
/// Subscribes to the device's HTTP event stream on port 44444 and receives
/// push results directly — no DMST installation, no DataMan Network Client,
/// no filesystem dependency.
///
/// ── Protocol (confirmed from Wireshark 2026-05-25) ───────────────────────
///
/// Both DMCC and this HTTP channel share port 44444.  The device distinguishes
/// connections by their opening bytes: raw XML = DMCC; HTTP GET = event stream.
/// All traffic runs on a SINGLE Keep-Alive TCP connection:
///
///   VTCCP → device :  GET /events?enable HTTP/1.1   (subscribe)
///   device → VTCCP :  HTTP/1.1 204 No Content        (ack, Content-Length: 0)
///   device → VTCCP :  PUT /status.xml     (~1/sec, telemetry — ignored)
///   device → VTCCP :  PUT /vs.cfg         (AES-encrypted config — ignored)
///   device → VTCCP :  PUT /pcm_report.html (per verification scan — HTML report)
///   device → VTCCP :  PUT /codes.xml      (per scan — push XML in base64 general block)
///
/// pcm_report.html always arrives BEFORE codes.xml for the same scan (confirmed
/// from stream ordering). This eliminates the timestamp-correlation complexity
/// in <see cref="DmstHtmlScraper.TryMergeAsync"/>; the two are paired sequentially.
///
/// ── Data flow ────────────────────────────────────────────────────────────
///
///   PUT /pcm_report.html body
///     → DmstHtmlScraper.ParseHtml()
///     → buffered as _pendingHtml
///
///   PUT /codes.xml (origin="common" only; "monitor" scans ignored)
///     → XDocument.Parse()
///     → extract &lt;full_string encoding="base64"&gt; from &lt;general&gt;
///     → Convert.FromBase64String() → UTF-8 push XML
///     → DmstResultParser.Parse()
///     → DmstReportValidator.MergeAndValidate(record, _pendingHtml)
///     → resultCallback
///
/// ── Replaces ─────────────────────────────────────────────────────────────
///
/// Fully replaces both <see cref="DmstListener"/> (push XML delivery) and
/// <see cref="DmstHtmlScraper"/> (HTML supplemental fields).  When this
/// subscriber is running, DMST does not need to be installed or open.
///
/// ── Lifecycle ────────────────────────────────────────────────────────────
///
///   1. var sub = new HttpEventSubscriber(host, port, map, ctx, callback);
///   2. await sub.StartAsync(ct);       // connects, subscribes, starts loop
///   3. [results arrive via callback]
///   4. await sub.StopAsync();          // or await sub.DisposeAsync()
/// </summary>
public sealed class HttpEventSubscriber : IAsyncDisposable
{
    private readonly string                     _host;
    private readonly int                        _port;
    private readonly VerificationXmlMap         _map;
    private readonly VerificationRecord         _context;
    private readonly Action<VerificationRecord> _callback;

    private TcpClient?               _tcp;
    private CancellationTokenSource? _cts;
    private Task?                    _receiveTask;

    /// <summary>
    /// Most recently parsed HTML report from PUT /pcm_report.html.
    /// pcm_report.html always precedes codes.xml for the same scan on the stream.
    /// Consumed (set null) when codes.xml arrives.
    /// </summary>
    private DmstHtmlReport? _pendingHtml;

    public bool IsRunning => _receiveTask is { IsCompleted: false };

    public HttpEventSubscriber(
        string                     host,
        int                        port,
        VerificationXmlMap         map,
        VerificationRecord         sessionContext,
        Action<VerificationRecord> resultCallback)
    {
        _host     = host;
        _port     = port;
        _map      = map;
        _context  = sessionContext;
        _callback = resultCallback;
    }

    // ── Lifecycle ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Opens a raw TCP connection to <c>host:port</c>, drains the device welcome
    /// banner, sends <c>GET /events?enable HTTP/1.1</c>, discards the 204 response,
    /// then starts the receive loop on a thread-pool task.
    /// </summary>
    public async Task StartAsync(CancellationToken ct = default)
    {
        if (IsRunning) return;

        _tcp = new TcpClient();
        await _tcp.ConnectAsync(_host, _port, ct);
        var stream = _tcp.GetStream();

        // Drain the device welcome banner (~100 B sent on raw TCP connect).
        try
        {
            using var bannerCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            bannerCts.CancelAfter(400);
            await stream.ReadAsync(new byte[512], bannerCts.Token);
        }
        catch (OperationCanceledException) { /* no banner or timeout — OK */ }
        catch { }

        // Send the HTTP subscribe request.
        // X-Peer: 0 — DMST sends a non-zero peer ID; 0 is accepted by the device
        // (the device uses the TCP source address for callback, not X-Peer).
        byte[] req = Encoding.ASCII.GetBytes(
            "GET /events?enable HTTP/1.1\r\n" +
            $"Date: {DateTime.UtcNow:R}\r\n" +
            "X-Peer: 0\r\n" +
            "\r\n");
        await stream.WriteAsync(req, ct);

        // Read and discard the HTTP/1.1 204 No Content response headers.
        // Content-Length: 0 — no body to skip.
        var (_, subscribeOk) = await ReadHeadersAsync(stream, new byte[1], ct);
        if (!subscribeOk)
            throw new InvalidOperationException(
                "HttpEventSubscriber: did not receive 204 ack from device.");

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-HTTP-SUB] Subscribed to {_host}:{_port}/events. Receive loop starting.");

        _cts         = CancellationTokenSource.CreateLinkedTokenSource(ct);
        _receiveTask = Task.Run(() => ReceiveLoopAsync(stream, _cts.Token), _cts.Token);
    }

    /// <summary>Cancels the receive loop and closes the TCP connection.</summary>
    public async Task StopAsync()
    {
        _cts?.Cancel();
        _tcp?.Close();      // unblocks any pending ReadAsync

        if (_receiveTask is not null)
            try { await _receiveTask; } catch { /* expected on cancel */ }

        _receiveTask = null;
        _cts?.Dispose();
        _cts         = null;
        _tcp?.Dispose();
        _tcp         = null;
        _pendingHtml = null;

        System.Diagnostics.Debug.WriteLine("[VTCCP-HTTP-SUB] Stopped.");
    }

    public async ValueTask DisposeAsync() => await StopAsync();

    // ── Receive loop ──────────────────────────────────────────────────────────

    private async Task ReceiveLoopAsync(NetworkStream stream, CancellationToken ct)
    {
        // bodyPool is reused across messages; grown if a message exceeds its size.
        // codes.xml with JPEG can be ~202 KB; 256 KB gives comfortable headroom.
        var bodyPool = new byte[256 * 1024];
        var oneByte  = new byte[1];

        while (!ct.IsCancellationRequested)
        {
            // Read HTTP request/response headers (byte-by-byte until \r\n\r\n).
            // Headers are short (~100–300 B); the bulk cost is in body reads below.
            var (headers, ok) = await ReadHeadersAsync(stream, oneByte, ct);
            if (!ok) break;

            string[] lines         = headers.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) continue;

            string  requestLine   = lines[0];                   // e.g. "PUT /codes.xml HTTP/1.1"
            string  path          = ParsePath(requestLine);     // e.g. "/codes.xml"
            int     contentLength = ParseContentLength(lines);
            string? dateHeader    = ParseHeaderValue(lines, "Date:");

            // Grow pool on-demand (rare — only if a codes.xml is larger than expected).
            if (contentLength > bodyPool.Length)
                bodyPool = new byte[contentLength + 4096];

            // Read exactly Content-Length bytes (bulk async — efficient for large bodies).
            if (!await ReadExactAsync(stream, bodyPool, contentLength, ct)) break;

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-HTTP-SUB] {requestLine} ({contentLength} B)");

            // Dispatch — all other paths (/status.xml, /vs.cfg, HTTP responses) ignored.
            if (path == "/pcm_report.html")
                HandleHtml(Encoding.UTF8.GetString(bodyPool, 0, contentLength), dateHeader);
            else if (path == "/codes.xml")
                HandleCodesXml(Encoding.UTF8.GetString(bodyPool, 0, contentLength));
        }

        System.Diagnostics.Debug.WriteLine("[VTCCP-HTTP-SUB] Receive loop ended.");
    }

    // ── /pcm_report.html handler ──────────────────────────────────────────────

    private void HandleHtml(string html, string? dateHeader)
    {
        // Build a synthetic path solely so ParseHtml can extract a correlation
        // timestamp. The HTTP stream transports report content, not the original
        // DMST filesystem filename, so clear file provenance before the parsed
        // report is merged into the record. If DMST also writes its local HTML
        // report, DmstHtmlScraper later supplies the genuine filename.
        string syntheticPath = MakeSyntheticCorrelationPath(dateHeader);
        _pendingHtml = DmstHtmlScraper.ParseHtml(
            html, syntheticPath, hasSyntheticSourcePath: true);

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-HTTP-SUB] pcm_report.html parsed: ok={_pendingHtml.ParseSucceeded} " +
            $"polarity={_pendingHtml.ImagePolarity ?? "null"} " +
            $"ECLevel={_pendingHtml.ECLevel ?? "null"} " +
            $"DataCW={_pendingHtml.DataCodewords?.ToString() ?? "null"}");
    }

    /// <summary>
    /// Constructs a clearly labelled synthetic local path whose filename matches the
    /// DMST timestamp format
    /// expected by <see cref="DmstHtmlScraper.ParseHtml"/>:
    /// <c>yyyy-MM-dd_HH-mm-ss-mmm_suffix.html</c>
    /// </summary>
    private static string MakeSyntheticCorrelationPath(string? dateHeader)
    {
        DateTime dt = DateTime.Now;
        if (dateHeader is not null)
            DateTime.TryParse(dateHeader, out dt);   // HTTP Date is always UTC
        // This path is synthetic correlation metadata only. Never convert it to
        // local time; report time comes from HTML Verified: or the real filename.
        return Path.Combine("C:", "HTTP_STREAM_PLACEHOLDER",
            $"{dt:yyyy-MM-dd_HH-mm-ss}-000_http.html");
    }

    // ── /codes.xml handler ────────────────────────────────────────────────────

    private void HandleCodesXml(string xml)
    {
        try
        {
            var doc  = XDocument.Parse(xml);
            var root = doc.Root;   // <result id="N" origin="common|monitor" ...>

            // Skip background monitoring scans — origin="monitor" has no TruCheck
            // data and no preceding pcm_report.html.
            string origin = root?.Attribute("origin")?.Value ?? "";
            if (!string.Equals(origin, "common", StringComparison.OrdinalIgnoreCase))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-HTTP-SUB] codes.xml origin='{origin}' — skipping.");
                return;
            }

            // Extract base64-encoded push XML.
            // Location: <result><general><full_string encoding="base64">...</full_string>
            string? base64 = root?
                .Element("general")?
                .Elements()
                .FirstOrDefault(e =>
                    e.Name.LocalName == "full_string" &&
                    string.Equals(e.Attribute("encoding")?.Value, "base64",
                                  StringComparison.OrdinalIgnoreCase))?
                .Value?.Trim();

            if (string.IsNullOrWhiteSpace(base64))
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-HTTP-SUB] codes.xml: <full_string encoding=\"base64\"> not found.");
                return;
            }

            // Decode to UTF-8 push XML document and parse.
            string pushXml = Encoding.UTF8.GetString(Convert.FromBase64String(base64));
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-HTTP-SUB] Push XML decoded: {pushXml.Length} chars.");

            var record = DmstResultParser.Parse(pushXml, _map, _context);

            // Merge with the buffered HTML report (always arrives first on the stream).
            // If _pendingHtml is null (e.g. connection joined mid-session), skip merge.
            var html    = _pendingHtml;
            _pendingHtml = null;
            if (html is not null)
                record = DmstReportValidator.MergeAndValidate(record, html);

            try { _callback(record); } catch { /* caller exception isolation */ }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-HTTP-SUB] HandleCodesXml: {ex.GetType().Name}: {ex.Message}");
        }
    }

    // ── Low-level HTTP stream helpers ─────────────────────────────────────────

    /// <summary>
    /// Reads bytes from <paramref name="stream"/> one at a time until the
    /// double-CRLF header terminator (<c>\r\n\r\n</c>) is found.
    /// Returns the accumulated header text and <c>ok=true</c>, or
    /// <c>(string.Empty, false)</c> on EOF or cancellation.
    /// </summary>
    private static async Task<(string headers, bool ok)> ReadHeadersAsync(
        NetworkStream stream, byte[] oneByte, CancellationToken ct)
    {
        var buf = new List<byte>(512);

        while (!ct.IsCancellationRequested)
        {
            int n;
            try   { n = await stream.ReadAsync(oneByte.AsMemory(), ct); }
            catch { return (string.Empty, false); }

            if (n <= 0) return (string.Empty, false);

            buf.Add(oneByte[0]);
            int c = buf.Count;
            if (c >= 4 &&
                buf[c-4] == '\r' && buf[c-3] == '\n' &&
                buf[c-2] == '\r' && buf[c-1] == '\n')
                return (Encoding.ASCII.GetString(buf.ToArray()), true);
        }

        return (string.Empty, false);
    }

    /// <summary>
    /// Reads exactly <paramref name="count"/> bytes from <paramref name="stream"/>
    /// into <paramref name="buf"/> starting at offset 0.
    /// Returns false on EOF or cancellation.
    /// </summary>
    private static async Task<bool> ReadExactAsync(
        NetworkStream stream, byte[] buf, int count, CancellationToken ct)
    {
        int read = 0;
        while (read < count && !ct.IsCancellationRequested)
        {
            int n;
            try   { n = await stream.ReadAsync(buf.AsMemory(read, count - read), ct); }
            catch { return false; }
            if (n <= 0) return false;
            read += n;
        }
        return true;
    }

    // ── Header parse helpers ──────────────────────────────────────────────────

    private static string ParsePath(string requestLine)
    {
        // "PUT /codes.xml HTTP/1.1"  →  "/codes.xml"
        // "HTTP/1.1 204 No Content"  →  ""   (response lines ignored)
        var parts = requestLine.Split(' ');
        return parts.Length >= 2 && !requestLine.StartsWith("HTTP")
            ? parts[1]
            : string.Empty;
    }

    private static int ParseContentLength(string[] lines)
    {
        foreach (var line in lines.Skip(1))
        {
            if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(line["Content-Length:".Length..].Trim(), out int len))
                    return len;
            }
        }
        return 0;
    }

    private static string? ParseHeaderValue(string[] lines, string prefix)
    {
        foreach (var line in lines.Skip(1))
            if (line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return line[prefix.Length..].Trim();
        return null;
    }
}
