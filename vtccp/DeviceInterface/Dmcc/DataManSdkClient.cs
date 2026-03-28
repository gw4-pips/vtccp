namespace DeviceInterface.Dmcc;

using System.Net;

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
    private readonly DeviceConfig          _cfg;
    private CognexSdk.EthSystemConnector?  _connector;
    private CognexSdk.DataManSystem?       _system;
    private bool                           _isConnected;
    private bool                           _disposed;

    // ── Public surface (mirrors DmccClient) ──────────────────────────────────

    /// <summary>True after a successful ConnectAsync and before DisconnectAsync.</summary>
    public bool IsConnected => _isConnected && _system != null;

    /// <summary>Not applicable for the SDK path — always null.</summary>
    public string? WelcomeBanner => null;

    public DataManSdkClient(DeviceConfig config)
    {
        _cfg = config ?? throw new ArgumentNullException(nameof(config));
    }

    // ── Connection ────────────────────────────────────────────────────────────

    /// <summary>
    /// Opens an SDK connection to the device.
    /// EthSystemConnector uses port 44444 by default (DataMan SDK protocol).
    /// Sets result types via the SDK's native API (avoids InvalidCommandException
    /// from sending raw SET DMCC.RESULT-FORMAT via SendCommand).
    /// </summary>
    public async Task ConnectAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (IsConnected) return;

        await Task.Run(() =>
        {
            var ip     = IPAddress.Parse(_cfg.Host);
            _connector = new CognexSdk.EthSystemConnector(ip);
            _system    = new CognexSdk.DataManSystem(_connector);
            _system.Connect();
            _isConnected = true;

            // Request XML results via the SDK's native API instead of the
            // raw "SET DMCC.RESULT-FORMAT FULL" string (which the SDK rejects
            // with InvalidCommandException on some firmware versions).
            try
            {
                _system.SetResultTypes(CognexSdk.ResultTypes.ReadXml);
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] Connected to {_cfg.Host}. SetResultTypes(XmlResult) OK.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] Connected to {_cfg.Host}. SetResultTypes failed: {ex.Message}");
            }
        }, ct);
    }

    /// <summary>Closes the SDK connection.</summary>
    public async Task DisconnectAsync()
    {
        _isConnected = false;

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
    ///
    /// SDK.SendCommand() returns Cognex.DataMan.SDK.DmccResponse; we call
    /// ToString() to obtain the wire-format string and parse it with our
    /// DmccResponse.Parse().  Every raw response is logged as [VTCCP-SDK].
    ///
    /// InvalidCommandException (SDK rejected the command before sending) and
    /// all other exceptions are caught and returned as NoResponse / ParseError
    /// rather than propagating — callers treat code -2 as "not supported".
    /// </summary>
    public async Task<DmccResponse> SendAsync(string command, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (!IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        return await Task.Run(() =>
        {
            try
            {
                var sdkResp = _system!.SendCommand(command);
                string raw  = sdkResp?.ToString() ?? string.Empty;

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] CMD '{command}' → " +
                    $"'{raw.Replace("\r", "\\r").Replace("\n", "\\n")}'");

                return DmccResponse.Parse(raw);
            }
            catch (CognexSdk.InvalidCommandException ex)
            {
                // SDK rejected the command string before it was sent to the device.
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] CMD '{command}' — InvalidCommandException: {ex.Message}");
                return DmccResponse.Parse(string.Empty); // code -2 NoResponse
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] CMD '{command}' — Exception ({ex.GetType().Name}): {ex.Message}");
                return DmccResponse.Parse(string.Empty); // code -2 NoResponse
            }
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
