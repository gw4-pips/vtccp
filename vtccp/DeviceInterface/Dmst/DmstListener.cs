namespace DeviceInterface.Dmst;

using System.Net;
using System.Net.Sockets;
using System.Text;
using ExcelEngine.Models;

/// <summary>
/// TCP listener that receives pushed results from the DataMan device and fires a
/// callback with the parsed <see cref="VerificationRecord"/>.
///
/// Two wire formats are supported automatically:
///
///   XML mode  — device sends a <DMCCResponse> (or other ResponseRoot) XML document.
///               Detected when the first non-whitespace byte is '&lt;'.
///               Used when Format Data scripting outputs a full XML blob.
///
///   Text mode — device sends plain-text lines (one decoded-string per CR/LF).
///               Used when Format Data is in Basic/Standard mode and the device
///               is configured with the Network Client push to our port.
///               Each non-empty line becomes a minimal VerificationRecord
///               (DecodedData populated; quality grades all null).
///
/// The device must be configured to connect to host_ip:port after each scan
/// (DataMan Setup Tool → Communications → Network Client → Enabled, Host, Port).
/// </summary>
public sealed class DmstListener
{
    private readonly int                       _port;
    private readonly VerificationXmlMap        _map;
    private readonly VerificationRecord        _context;
    private readonly Action<VerificationRecord> _callback;
    private TcpListener?              _server;
    private CancellationTokenSource?  _cts;
    private Task?                     _listenTask;

    public bool IsRunning => _listenTask is { IsCompleted: false };

    public DmstListener(
        int                         port,
        VerificationXmlMap          map,
        VerificationRecord          sessionContext,
        Action<VerificationRecord>  resultCallback)
    {
        _port     = port;
        _map      = map;
        _context  = sessionContext;
        _callback = resultCallback;
    }

    // ── Lifecycle ─────────────────────────────────────────────────────────────

    public Task StartAsync(CancellationToken ct = default)
    {
        if (IsRunning) return Task.CompletedTask;

        _cts    = CancellationTokenSource.CreateLinkedTokenSource(ct);
        _server = new TcpListener(IPAddress.Any, _port);
        _server.Start();
        _listenTask = Task.Run(() => AcceptLoopAsync(_cts.Token), _cts.Token);
        return Task.CompletedTask;
    }

    public async Task StopAsync()
    {
        _cts?.Cancel();
        _server?.Stop();
        if (_listenTask is not null)
        {
            try { await _listenTask; } catch { /* expected on cancel */ }
        }
        _listenTask = null;
        _cts?.Dispose();
        _cts = null;
    }

    // ── Accept loop ───────────────────────────────────────────────────────────

    private async Task AcceptLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            TcpClient client;
            try   { client = await _server!.AcceptTcpClientAsync(ct); }
            catch { break; }

            _ = Task.Run(() => HandleClientAsync(client, ct), ct);
        }
    }

    // ── Per-client handler ────────────────────────────────────────────────────

    private async Task HandleClientAsync(TcpClient client, CancellationToken ct)
    {
        using (client)
        using (var stream = client.GetStream())
        {
            var sb  = new StringBuilder();
            var buf = new byte[8192];
            bool? isXml = null;   // null = not yet determined

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-DMST] Client connected from {client.Client.RemoteEndPoint}");

            while (!ct.IsCancellationRequested)
            {
                int n;
                try   { n = await stream.ReadAsync(buf, ct); }
                catch (Exception ex) {
                    System.Diagnostics.Debug.WriteLine($"[VTCCP-DMST] ReadAsync exception: {ex.Message}");
                    break;
                }

                if (n <= 0) {
                    System.Diagnostics.Debug.WriteLine("[VTCCP-DMST] Connection closed by device (n<=0).");
                    break;
                }

                string chunk = Encoding.UTF8.GetString(buf, 0, n);
                sb.Append(chunk);

                // Determine format on first non-whitespace byte received.
                if (isXml is null)
                {
                    isXml = sb.ToString().TrimStart().StartsWith('<');
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-DMST] First chunk ({n} bytes), isXml={isXml}. " +
                        $"Preview: {chunk[..Math.Min(120, chunk.Length)].Replace("\r","\\r").Replace("\n","\\n")}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[VTCCP-DMST] Subsequent chunk ({n} bytes).");
                }

                if (isXml == true)
                    ProcessXmlBuffer(sb);
                else
                    ProcessTextBuffer(sb);
            }

            // Drain any remaining text (connection closed without trailing newline).
            if (isXml == false)
            {
                string remaining = sb.ToString().Trim();
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-DMST] Draining text buffer ({remaining.Length} chars): {remaining[..Math.Min(80, remaining.Length)]}");
                if (remaining.Length > 0)
                    FireText(remaining);
            }
            else if (isXml is null)
            {
                System.Diagnostics.Debug.WriteLine("[VTCCP-DMST] Connection closed with no data received.");
            }

            System.Diagnostics.Debug.WriteLine("[VTCCP-DMST] HandleClientAsync exiting.");
        }
    }

    // ── XML buffer processing ─────────────────────────────────────────────────

    /// <summary>
    /// Fires once a complete XML document is accumulated (detected by closing root tag).
    /// Leaves any data after the closing tag in the buffer (rare, but defensive).
    /// </summary>
    private void ProcessXmlBuffer(StringBuilder sb)
    {
        string accumulated = sb.ToString();
        string trimmed     = accumulated.TrimEnd();

        string closingTag  = $"</{_map.ResponseRoot}>";
        int    tagIdx      = trimmed.LastIndexOf(closingTag, StringComparison.OrdinalIgnoreCase);

        if (tagIdx < 0)
        {
            // Also accept a bare XML declaration end (<?xml …?>) as a single-line document.
            if (trimmed.EndsWith("?>", StringComparison.Ordinal))
                tagIdx = 0;
        }

        if (tagIdx < 0)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-DMST] XML closing tag '{closingTag}' not yet found (buffered {accumulated.Length} chars). Buffering.");
            return;   // incomplete — keep buffering
        }

        int endPos = tagIdx + closingTag.Length;
        string xmlDoc = accumulated[..endPos];
        string leftover = accumulated[endPos..];

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-DMST] Complete XML document found ({xmlDoc.Length} chars). Parsing. " +
            $"XML preview: {xmlDoc[..Math.Min(600, xmlDoc.Length)].Replace("\r","").Replace("\n"," ")}");

        sb.Clear();
        sb.Append(leftover);

        var record = DmstResultParser.Parse(xmlDoc, _map, _context);
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-DMST] Parsed record — DecodedData='{record.DecodedData}', Grade='{record.OverallGrade?.LetterGradeString}'. Firing callback.");
        Fire(record);
    }

    // ── Plain-text buffer processing ──────────────────────────────────────────

    /// <summary>
    /// Splits the buffer on newlines and fires a record for each complete line.
    /// The last (potentially incomplete) fragment is left in the buffer.
    /// </summary>
    private void ProcessTextBuffer(StringBuilder sb)
    {
        string data = sb.ToString();
        int lastNewline = data.LastIndexOfAny(['\n', '\r']);

        if (lastNewline < 0)
        {
            // No newline found — Cognex DataMan Network Client in persistent-connection
            // mode sends each scan result as one TCP write with no trailing delimiter.
            // Fire the whole buffer as a complete record immediately.
            string all = data.Trim();
            if (all.Length > 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-DMST] Text mode, no newline — firing as complete record: '{all}'");
                sb.Clear();
                FireText(all);
            }
            return;
        }

        string complete = data[..(lastNewline + 1)];
        string leftover = data[(lastNewline + 1)..];

        sb.Clear();
        sb.Append(leftover);

        foreach (string raw in complete.Split('\n'))
        {
            string line = raw.Trim('\r', '\n', ' ');
            if (line.Length > 0)
                FireText(line);
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private void Fire(VerificationRecord r)
    {
        try { _callback(r); } catch { /* caller exception isolation */ }
    }

    private void FireText(string decodedLine)
    {
        var record = DmstResultParser.ParseText(decodedLine, _context);
        Fire(record);
    }
}
