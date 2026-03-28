namespace DeviceInterface.Dmcc;

// Alias avoids name collision: the SDK also defines DmccResponse.
using CognexSdk = Cognex.DataMan.SDK;

/// <summary>
/// DMCC client backed by the Cognex DataMan SDK (Cognex.DataMan.SDK.PC.dll).
///
/// Replaces the hand-rolled ASCII TCP client (DmccClient).  The SDK speaks the
/// binary proprietary protocol used on port 44444 (EthSystemConnector default).
///
/// API is intentionally identical to DmccClient so DeviceSession needs only a
/// one-line type change to switch between implementations.
/// </summary>
public sealed class DataManSdkClient : IAsyncDisposable
{
    private readonly DeviceConfig         _cfg;
    private CognexSdk.EthSystemConnector? _connector;
    private CognexSdk.DataManSystem?      _system;
    private bool                          _disposed;

    // ── Public surface (mirrors DmccClient) ──────────────────────────────────

    /// <summary>True after a successful ConnectAsync and before DisconnectAsync.</summary>
    public bool IsConnected => _system?.IsConnected ?? false;

    /// <summary>Not applicable for the SDK path — always null.</summary>
    public string? WelcomeBanner => null;

    public DataManSdkClient(DeviceConfig config)
    {
        _cfg = config ?? throw new ArgumentNullException(nameof(config));
    }

    // ── Connection ────────────────────────────────────────────────────────────

    /// <summary>
    /// Opens an SDK connection to the device.
    /// EthSystemConnector uses port 44444 by default (the DataMan SDK protocol).
    /// </summary>
    public async Task ConnectAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (IsConnected) return;

        await Task.Run(() =>
        {
            _connector = new CognexSdk.EthSystemConnector(_cfg.Host);
            _system    = new CognexSdk.DataManSystem(_connector);
            _system.Connect();

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SDK] Connected to {_cfg.Host}.  IsConnected={_system.IsConnected}");
        }, ct);
    }

    /// <summary>Closes the SDK connection.</summary>
    public async Task DisconnectAsync()
    {
        await Task.Run(() =>
        {
            try
            {
                _system?.Disconnect();
                System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] Disconnected.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VTCCP-SDK] Disconnect error: {ex.Message}");
            }
        });

        _system    = null;
        _connector = null;
    }

    // ── Command exchange ──────────────────────────────────────────────────────

    /// <summary>
    /// Sends a DMCC command via the SDK and returns a parsed DmccResponse.
    /// SendCommand returns the raw DMCC response string (status + body);
    /// we feed it to DmccResponse.Parse() to produce a strongly-typed result.
    /// </summary>
    public async Task<DmccResponse> SendAsync(string command, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (!IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        return await Task.Run(() =>
        {
            string? raw = _system!.SendCommand(command);

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SDK] CMD '{command}' → " +
                $"'{(raw ?? "<null>").Replace("\r", "\\r").Replace("\n", "\\n")}'");

            return DmccResponse.Parse(raw ?? string.Empty);
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
        if (_disposed) throw new ObjectDisposedException(nameof(DataManSdkClient));
    }
}
