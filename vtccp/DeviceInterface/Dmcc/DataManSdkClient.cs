namespace DeviceInterface.Dmcc;

using System.Net;

// Alias avoids name collision: the SDK also defines DmccResponse.
using CognexSdk = Cognex.DataMan.SDK;

/// <summary>
/// DMCC client backed by the Cognex DataMan SDK (Cognex.DataMan.SDK.PC.dll).
///
/// Key SDK behaviours discovered via reflection / runtime testing:
///   - EthSystemConnector takes IPAddress, not string.
///   - DataManSystem has no IsConnected; use local bool.
///   - SendCommand() returns Cognex.DataMan.SDK.DmccResponse whose ToString()
///     returns the class name — StatusCode / Body must be read via reflection.
///   - "GET FIRMWARE.VER" → InvalidCommandException; use _system.FirmwareVersion.
///   - "TRIGGER" → InvalidParameterException; "TRIGGER 1" is the correct form.
///   - "GET SYMBOL.RESULT" → InvalidCommandException; use XmlResultArrived event.
///   - SetResultTypes() uses ResultTypes.ReadXml (= 2), not XmlResult.
/// </summary>
public sealed class DataManSdkClient : IAsyncDisposable
{
    private readonly DeviceConfig          _cfg;
    private CognexSdk.EthSystemConnector?  _connector;
    private CognexSdk.DataManSystem?       _system;
    private bool                           _isConnected;
    private bool                           _disposed;

    // ── Public surface (mirrors DmccClient) ──────────────────────────────────

    public bool    IsConnected     => _isConnected && _system != null;
    public string? WelcomeBanner   => null;

    /// <summary>Firmware version read directly from the SDK property after Connect().</summary>
    public string? FirmwareVersion => _system?.FirmwareVersion;

    public DataManSdkClient(DeviceConfig config)
    {
        _cfg = config ?? throw new ArgumentNullException(nameof(config));
    }

    // ── Connection ────────────────────────────────────────────────────────────

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

            try
            {
                _system.SetResultTypes(CognexSdk.ResultTypes.ReadXml);
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] Connected to {_cfg.Host}.  " +
                    $"FW={_system.FirmwareVersion}  SetResultTypes(ReadXml) OK.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] Connected to {_cfg.Host}.  SetResultTypes failed: {ex.Message}");
            }
        }, ct);
    }

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
    /// Sends a DMCC command and returns a parsed DmccResponse.
    /// Uses reflection to extract StatusCode and body from the SDK's DmccResponse
    /// (whose ToString() unhelpfully returns the class name).
    /// InvalidCommandException / InvalidParameterException → code -2 (NoResponse).
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

                int    code = TryGetIntProp(sdkResp, "StatusCode", "Code") ?? -2;
                string body = TryGetStrProp(sdkResp, "Body", "Value", "Message", "Result")
                              ?? string.Empty;

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] CMD '{command}' → code={code}  " +
                    $"body='{(body.Length > 120 ? body[..120] + "…" : body)}'");

                // First call: dump all available DmccResponse properties for diagnostics.
                if (command == DmccCommand.GetDeviceType)
                    DumpProps("[VTCCP-SDK] DmccResponse", sdkResp);

                string raw = $"\r\n{code}\r\n\r\n{body}";
                return DmccResponse.Parse(raw);
            }
            catch (CognexSdk.InvalidCommandException ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] CMD '{command}' — InvalidCommandException: {ex.Message}");
                return DmccResponse.Parse(string.Empty);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] CMD '{command}' — {ex.GetType().Name}: {ex.Message}");
                return DmccResponse.Parse(string.Empty);
            }
        }, ct);
    }

    // ── Trigger + event-based result collection ───────────────────────────────

    /// <summary>
    /// Fires a software trigger via the SDK and waits for the device to deliver
    /// an XML verification result through the XmlResultArrived event.
    ///
    /// Returns the raw XML string, or null on timeout / no read.
    ///
    /// Why event-based: "GET SYMBOL.RESULT" throws InvalidCommandException — the
    /// SDK does not expose that command. Results must be consumed via event.
    /// </summary>
    public async Task<string?> TriggerAndWaitForXmlAsync(
        int               timeoutMs = 10_000,
        CancellationToken ct        = default)
    {
        ThrowIfDisposed();
        if (!IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        var tcs = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);

        CognexSdk.XmlResultArrivedHandler xmlHandler = (_, args) =>
        {
            // XmlResultArrivedEventArgs.XmlResult — name confirmed by XmlResultArrivedEventArgs type.
            var prop = args.GetType().GetProperty("XmlResult")
                    ?? args.GetType().GetProperty("Xml")
                    ?? args.GetType().GetProperty("Result");

            if (prop is null)
            {
                // Fallback: dump all properties so we can find the right name.
                DumpProps("[VTCCP-SDK] XmlResultArrivedEventArgs", args);
            }

            string? xml = prop?.GetValue(args)?.ToString();
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SDK] XmlResultArrived: {xml?.Length ?? 0} chars");
            tcs.TrySetResult(xml);
        };

        _system!.XmlResultArrived += xmlHandler;
        try
        {
            // Fire the trigger.  Try each form in turn; every exception is caught
            // so the enclosing await never throws and the timeout path stays clean.
            bool triggered = false;
            await Task.Run(() =>
            {
                // Attempt 1: TRIGGER 1
                try
                {
                    _system!.SendCommand("TRIGGER 1");
                    System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] TRIGGER 1 sent OK.");
                    triggered = true;
                    return;
                }
                catch (Exception ex1)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SDK] TRIGGER 1 failed ({ex1.GetType().Name}: {ex1.Message}), trying plain TRIGGER...");
                }

                // Attempt 2: plain TRIGGER
                try
                {
                    _system!.SendCommand("TRIGGER");
                    System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] TRIGGER sent OK.");
                    triggered = true;
                }
                catch (Exception ex2)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SDK] TRIGGER also failed ({ex2.GetType().Name}: {ex2.Message}).");
                }
            }, ct);

            if (!triggered)
            {
                tcs.TrySetResult(null);
            }

            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            cts.CancelAfter(timeoutMs);
            try
            {
                return await tcs.Task.WaitAsync(cts.Token);
            }
            catch (OperationCanceledException)
            {
                System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] TriggerAndWaitForXml: timed out.");
                return null;
            }
        }
        finally
        {
            _system.XmlResultArrived -= xmlHandler;
        }
    }

    // ── Reflection helpers ────────────────────────────────────────────────────

    private static int? TryGetIntProp(object obj, params string[] names)
    {
        foreach (var name in names)
        {
            var p = obj.GetType().GetProperty(name);
            if (p?.GetValue(obj) is int v) return v;
        }
        return null;
    }

    private static string? TryGetStrProp(object obj, params string[] names)
    {
        foreach (var name in names)
        {
            var p = obj.GetType().GetProperty(name);
            if (p != null) return p.GetValue(obj)?.ToString();
        }
        return null;
    }

    private static void DumpProps(string label, object obj)
    {
        foreach (var p in obj.GetType().GetProperties())
        {
            try
            {
                var v = p.GetValue(obj);
                System.Diagnostics.Debug.WriteLine($"{label}.{p.Name} = {v}");
            }
            catch { }
        }
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
