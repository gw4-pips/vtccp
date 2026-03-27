namespace DeviceInterface.Dmst;

using System.Net;
using System.Net.Sockets;
using System.Text;
using ExcelEngine.Models;

/// <summary>
/// TCP listener that receives pushed DMST XML results from the DataMan device
/// and fires a callback with the parsed <see cref="VerificationRecord"/>.
///
/// The device must be configured to send DMST output to host_ip:port after each scan.
/// This listener accepts one connection (the device), receives XML push payloads,
/// parses each one, and raises the result callback.
/// </summary>
public sealed class DmstListener
{
    private readonly int                _port;
    private readonly VerificationXmlMap _map;
    private readonly VerificationRecord _context;
    private readonly Action<VerificationRecord> _callback;
    private TcpListener?       _server;
    private CancellationTokenSource? _cts;
    private Task?              _listenTask;

    public bool IsRunning => _listenTask is { IsCompleted: false };

    public DmstListener(
        int                 port,
        VerificationXmlMap  map,
        VerificationRecord  sessionContext,
        Action<VerificationRecord> resultCallback)
    {
        _port     = port;
        _map      = map;
        _context  = sessionContext;
        _callback = resultCallback;
    }

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
            try { await _listenTask; } catch { /* expected cancellation */ }
        }
        _listenTask = null;
        _cts?.Dispose();
        _cts = null;
    }

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

    private async Task HandleClientAsync(TcpClient client, CancellationToken ct)
    {
        using (client)
        using (var stream = client.GetStream())
        {
            var sb  = new StringBuilder();
            var buf = new byte[8192];

            while (!ct.IsCancellationRequested)
            {
                int n;
                try   { n = await stream.ReadAsync(buf, ct); }
                catch { break; }

                if (n <= 0) break;
                sb.Append(Encoding.UTF8.GetString(buf, 0, n));

                // Each DMST push is a complete XML document; detect end of doc.
                string accumulated = sb.ToString();
                string trimmed = accumulated.TrimEnd();
                if (trimmed.EndsWith($"</{_map.ResponseRoot}>", StringComparison.OrdinalIgnoreCase)
                 || trimmed.EndsWith("?>", StringComparison.OrdinalIgnoreCase))
                {
                    // Parse and fire callback.
                    var record = DmstResultParser.Parse(accumulated, _map, _context);
                    try { _callback(record); } catch { /* caller exception isolation */ }
                    sb.Clear();
                }
            }
        }
    }
}
