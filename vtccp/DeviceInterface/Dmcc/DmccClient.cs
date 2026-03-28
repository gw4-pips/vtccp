namespace DeviceInterface.Dmcc;

using System.Net.Sockets;
using System.Text;
using System.Threading;

/// <summary>
/// Async TCP client for the Cognex DataMan DMCC (DataMan Command and Control) protocol.
///
/// Lifecycle:
///   1. Construct with DeviceConfig.
///   2. await ConnectAsync() — opens TCP socket, reads welcome banner.
///   3. await SendAsync(command) — sends a DMCC command, returns parsed DmccResponse.
///   4. await DisconnectAsync() / Dispose() — closes the socket.
///
/// Thread safety: one command at a time per instance. For concurrent access wrap
/// calls in a SemaphoreSlim.
/// </summary>
public sealed class DmccClient : IAsyncDisposable
{
    private readonly DeviceConfig _cfg;
    private TcpClient?     _tcp;
    private NetworkStream? _stream;
    private bool           _disposed;

    public bool IsConnected => _tcp?.Connected == true && !_disposed;

    /// <summary>Banner string sent by the device immediately after TCP accept.</summary>
    public string? WelcomeBanner { get; private set; }

    public DmccClient(DeviceConfig config)
    {
        _cfg = config ?? throw new ArgumentNullException(nameof(config));
    }

    // ── Connection ────────────────────────────────────────────────────────────

    /// <summary>Opens a TCP connection to the device and reads the welcome banner.</summary>
    public async Task ConnectAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (IsConnected) return;

        _tcp = new TcpClient();
        _tcp.ReceiveTimeout = _cfg.ResponseTimeoutMs;
        _tcp.SendTimeout    = _cfg.ResponseTimeoutMs;

        using var connectCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        connectCts.CancelAfter(_cfg.ConnectTimeoutMs);

        await _tcp.ConnectAsync(_cfg.Host, _cfg.Port, connectCts.Token);
        _stream = _tcp.GetStream();

        // Read welcome banner with a short timeout — many devices send nothing
        // and we must not stall for the full ResponseTimeoutMs here.
        WelcomeBanner = await ReadUntilIdleAsync(ct, _cfg.BannerTimeoutMs);

        // ── DMCC handshake ────────────────────────────────────────────────────
        // DataMan firmware requires the client to send "DMCC\r\n" immediately
        // after connecting before it will respond to any DMCC commands.
        // Without this, the device accepts the TCP connection but silently
        // ignores every subsequent command (returns nothing → code -2).
        // DataMan Setup Tool (DMST) sends this handshake automatically.
        byte[] handshake = Encoding.ASCII.GetBytes("DMCC\r\n");
        await _stream.WriteAsync(handshake, ct);
        await _stream.FlushAsync(ct);

        // Read the handshake acknowledgement (typically "\r\n0\r\n").
        // We don't check the status — some firmware sends 0, others send nothing.
        string handshakeAck = await ReadUntilIdleAsync(ct, _cfg.BannerTimeoutMs);
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-DMCC] Handshake ack ({handshakeAck.Length} chars): " +
            $"'{handshakeAck.Replace("\r", "\\r").Replace("\n", "\\n")}'");
    }

    /// <summary>Gracefully closes the connection.</summary>
    public async Task DisconnectAsync()
    {
        if (_stream is not null) { await _stream.FlushAsync(); _stream.Close(); }
        _tcp?.Close();
        _stream = null;
        _tcp    = null;
    }

    // ── Command exchange ──────────────────────────────────────────────────────

    /// <summary>
    /// Sends a DMCC command and returns the parsed response.
    /// Appends CRLF automatically; caller should NOT include it.
    /// </summary>
    public async Task<DmccResponse> SendAsync(string command, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (!IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        byte[] cmd = Encoding.ASCII.GetBytes(command + "\r\n");
        await _stream!.WriteAsync(cmd, ct);
        await _stream.FlushAsync(ct);

        string raw = await ReadUntilIdleAsync(ct);
        return DmccResponse.Parse(raw);
    }

    // ── Read helper ───────────────────────────────────────────────────────────

    /// <summary>
    /// Reads bytes until no new bytes arrive for <see cref="DeviceConfig.IdleGapMs"/>
    /// milliseconds, or the overall timeout elapses.
    ///
    /// <paramref name="overrideTimeoutMs"/> overrides <see cref="DeviceConfig.ResponseTimeoutMs"/>
    /// for this call only — used for the welcome-banner read which should time out quickly.
    ///
    /// Implementation note: NetworkStream.ReadAsync with a CancellationToken disposes the
    /// underlying socket on Linux when the token fires, destroying the connection.  To avoid
    /// this we run a synchronous Socket.Receive loop on a thread-pool thread; synchronous
    /// socket reads respect ReceiveTimeout without touching the socket lifetime.
    /// </summary>
    private Task<string> ReadUntilIdleAsync(CancellationToken ct, int? overrideTimeoutMs = null)
    {
        var socket = _tcp!.Client;
        int idleGap    = _cfg.IdleGapMs;
        int responseMs = overrideTimeoutMs ?? _cfg.ResponseTimeoutMs;

        return Task.Run(() =>
        {
            var sb       = new StringBuilder();
            var buf      = new byte[4096];
            var deadline = DateTime.UtcNow.AddMilliseconds(responseMs);

            socket.ReceiveTimeout = idleGap;   // governs synchronous Receive()

            while (DateTime.UtcNow < deadline)
            {
                ct.ThrowIfCancellationRequested();

                int n;
                try
                {
                    n = socket.Receive(buf);
                }
                catch (SocketException ex) when (ex.SocketErrorCode == SocketError.TimedOut)
                {
                    // Idle gap elapsed — no bytes arrived.
                    if (sb.Length > 0) break;              // end of response
                    if (DateTime.UtcNow >= deadline) break; // overall timeout
                    continue;                              // keep waiting for first byte
                }

                if (n <= 0) break;  // connection closed by remote side
                sb.Append(Encoding.ASCII.GetString(buf, 0, n));
            }

            socket.ReceiveTimeout = 0;   // restore blocking mode
            return sb.ToString();
        }, ct);
    }

    // ── IAsyncDisposable ──────────────────────────────────────────────────────

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;
        await DisconnectAsync();
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DmccClient));
    }
}
