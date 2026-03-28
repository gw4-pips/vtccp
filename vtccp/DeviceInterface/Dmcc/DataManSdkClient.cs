namespace DeviceInterface.Dmcc;

using Cognex.DataMan.SDK;
using Cognex.DataMan.SDK.Utils;

/// <summary>
/// DMCC client backed by the Cognex DataMan SDK (Cognex.DataMan.SDK.PC.dll).
///
/// Replaces the hand-rolled ASCII TCP client (DmccClient).  The SDK speaks the
/// binary proprietary protocol that the DMV475 uses on port 44444, which our
/// raw ASCII client could never reach.
///
/// API is intentionally identical to DmccClient so DeviceSession needs only a
/// one-line type change to switch between implementations.
/// </summary>
public sealed class DataManSdkClient : IAsyncDisposable
{
    private readonly DeviceConfig         _cfg;
    private EthernetDataManSystem?        _eth;
    private DataManSystem?                _system;
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
    /// EthernetDataManSystem uses port 44444 by default (the proprietary DataMan
    /// SDK protocol port), which is what the DMV475 expects.
    /// </summary>
    public async Task ConnectAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (IsConnected) return;

        await Task.Run(() =>
        {
            _eth    = new EthernetDataManSystem(_cfg.Host);
            _system = new DataManSystem(_eth);
            _system.Connect();

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SDK] Connected to {_cfg.Host} via DataMan SDK. " +
                $"IsConnected={_system.IsConnected}");
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

        _system = null;
        _eth    = null;
    }

    // ── Command exchange ──────────────────────────────────────────────────────

    /// <summary>
    /// Sends a DMCC command via the SDK and returns a parsed DmccResponse.
    /// The SDK handles framing and protocol details; we receive the raw
    /// response string and parse it with DmccResponse.Parse().
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
