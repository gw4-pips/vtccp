namespace DeviceInterface.Dmcc;

using System.Net;
using System.Net.Sockets;
using System.Text;

// Alias avoids name collision: the SDK also defines DmccResponse.
using CognexSdk = Cognex.DataMan.SDK;

/// <summary>
/// DMCC client backed by the Cognex DataMan SDK (Cognex.DataMan.SDK.PC.dll).
///
/// Key SDK behaviours discovered via reflection / runtime testing:
///   - EthSystemConnector takes IPAddress, not string.
///   - DataManSystem has no IsConnected; use local bool.
///   - SendCommand() returns Cognex.DataMan.SDK.DmccResponse whose body is in
///     the "PayLoad" property; SDK throws exceptions on failure (code is always 0).
///   - "GET FIRMWARE.VER" → InvalidCommandException; use _system.FirmwareVersion.
///   - "TRIGGER" / "TRIGGER 1" → InvalidParameterException from SDK's own
///     validation layer (firmware 6.1.16_sr4 / SDK v25 mismatch). Bypassed via
///     raw TCP on _cfg.Port; SDK connection kept alive for XmlResultArrived.
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

                // SDK throws an exception on failure; reaching here means success (code 0).
                // The body is in the "PayLoad" property (confirmed by reflection dump).
                // ResponseId is an SDK-internal tracker, NOT the DMCC status code.
                const int code = 0;
                string body = TryGetStrProp(sdkResp, "PayLoad", "Body", "Value", "Message", "Result")
                              ?? string.Empty;

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] CMD '{command}' → code={code}  " +
                    $"body='{(body.Length > 120 ? body[..120] + "…" : body)}'");

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
                // Both SDK TRIGGER forms were rejected by the SDK's own parameter
                // validation layer (not the device) — firmware 6.1.16_sr4 / SDK v25
                // version mismatch.  Bypass the SDK entirely: open a second raw TCP
                // connection to the DMCC port and send "TRIGGER\r\n" directly.
                // The SDK connection stays alive to deliver XmlResultArrived.
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-SDK] SDK rejected TRIGGER — trying raw TCP bypass...");
                try
                {
                    using var tcp = new TcpClient();
                    await tcp.ConnectAsync(_cfg.Host, _cfg.Port, ct);
                    using var stream = tcp.GetStream();

                    // Drain any welcome banner the device may send on connect.
                    try
                    {
                        using var bannerCts = new CancellationTokenSource(400);
                        byte[] buf = new byte[512];
                        await stream.ReadAsync(buf, bannerCts.Token);
                    }
                    catch (OperationCanceledException) { /* no banner or timeout — OK */ }
                    catch { }

                    // Send the bare TRIGGER command.
                    await stream.WriteAsync(Encoding.ASCII.GetBytes("TRIGGER\r\n"), ct);
                    System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] TRIGGER sent via raw TCP.");
                    triggered = true;
                }
                catch (Exception tcpEx)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SDK] Raw TCP TRIGGER failed: {tcpEx.GetType().Name}: {tcpEx.Message}");
                }
            }

            if (!triggered)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-SDK] All TRIGGER attempts exhausted — aborting.");
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

    // ── D4 Image Load — LoadImage + IMAGE.REPLAY ─────────────────────────────

    /// <summary>
    /// Loads a locally-stored symbol image into the device's image buffer, then
    /// sends <c>IMAGE.REPLAY</c> to fire TruCheck verification on the loaded image,
    /// and waits for the result to arrive via <c>XmlResultArrived</c>.
    ///
    /// Image load strategy: reads the file as raw bytes and passes them to
    /// <c>DataManSystem.SendCommand("IMAGE.LOAD", bytes)</c> via reflection,
    /// avoiding a compile-time dependency on the Windows-only Cognex SDK on
    /// Linux build agents.  The SDK's SendCommand(String, Byte[]) overload is
    /// the correct mechanism — no LoadImage method exists on DataManSystem.
    ///
    /// D4 result discriminator (confirmed scan #13, 2026-05-24):
    ///   • ContrastUniformity = -1, MRD = -1  →  OpticsSource = "LoadedImage"
    ///   • r.image.exposureTime = 0, autoExposure = false  (secondary indicator)
    ///
    /// Returns the raw XML string, or null on timeout / load failure.
    /// </summary>
    public async Task<string?> LoadAndReplayImageAsync(
        string            filePath,
        int               timeoutMs = 30_000,
        CancellationToken ct        = default)
    {
        ThrowIfDisposed();
        if (!IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        if (!File.Exists(filePath))
        {
            System.Diagnostics.Debug.WriteLine($"[VTCCP-D4] File not found: '{filePath}'");
            return null;
        }

        var tcs = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);

        CognexSdk.XmlResultArrivedHandler xmlHandler = (_, args) =>
        {
            var prop = args.GetType().GetProperty("XmlResult")
                    ?? args.GetType().GetProperty("Xml")
                    ?? args.GetType().GetProperty("Result");
            if (prop is null) DumpProps("[VTCCP-D4] XmlResultArrivedEventArgs", args);
            string? xml = prop?.GetValue(args)?.ToString();
            System.Diagnostics.Debug.WriteLine($"[VTCCP-D4] XmlResultArrived: {xml?.Length ?? 0} chars");
            tcs.TrySetResult(xml);
        };

        _system!.XmlResultArrived += xmlHandler;
        try
        {
            // Load image into device buffer: read file as bytes, send via SDK.
            bool loaded = await Task.Run(() => TrySendImageViaCommand(filePath), ct);

            // Give the device firmware time to ingest the image before IMAGE.REPLAY.
            if (loaded)
                await Task.Delay(400, ct);

            // Send IMAGE.REPLAY to fire TruCheck on the loaded image.
            // Result arrives asynchronously via XmlResultArrived — not the SendAsync body.
            var replayResp = await SendAsync(DmccCommand.ImageReplay, ct);

            // Abort only when both steps failed — if load succeeded, always wait for the event.
            if (!loaded && replayResp.StatusCode != 0)
            {
                Console.Error.WriteLine(
                    $"[VTCCP-D4] IMAGE.LOAD + IMAGE.REPLAY both failed (replay code={replayResp.StatusCode}).");
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
                System.Diagnostics.Debug.WriteLine("[VTCCP-D4] LoadAndReplayImage: timed out.");
                return null;
            }
        }
        finally
        {
            _system!.XmlResultArrived -= xmlHandler;
        }
    }

    // ── Replay-only (image already loaded) ───────────────────────────────────

    /// <summary>
    /// Sends IMAGE.REPLAY to re-grade the currently loaded image buffer and waits
    /// for the result via XmlResultArrived.  The caller is responsible for ensuring
    /// an image is already present in the device buffer — no LoadImage call is made.
    ///
    /// Used by Repeatability Analysis to loop IMAGE.REPLAY N times on a fixed image.
    /// Returns the raw XML string, or null on timeout / REPLAY rejection.
    /// </summary>
    public async Task<string?> ReplayAndWaitForXmlAsync(
        int               timeoutMs = 15_000,
        CancellationToken ct        = default)
    {
        ThrowIfDisposed();
        if (!IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        var tcs = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);

        CognexSdk.XmlResultArrivedHandler xmlHandler = (_, args) =>
        {
            var prop = args.GetType().GetProperty("XmlResult")
                    ?? args.GetType().GetProperty("Xml")
                    ?? args.GetType().GetProperty("Result");
            if (prop is null) DumpProps("[VTCCP-REPLAY] XmlResultArrivedEventArgs", args);
            string? xml = prop?.GetValue(args)?.ToString();
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-REPLAY] XmlResultArrived: {xml?.Length ?? 0} chars");
            tcs.TrySetResult(xml);
        };

        _system!.XmlResultArrived += xmlHandler;
        try
        {
            var replayResp = await SendAsync(DmccCommand.ImageReplay, ct);
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-REPLAY] IMAGE.REPLAY → code={replayResp.StatusCode} body='{replayResp.Body}'");

            if (replayResp.StatusCode != 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-REPLAY] IMAGE.REPLAY rejected by device — aborting wait.");
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
                System.Diagnostics.Debug.WriteLine("[VTCCP-REPLAY] ReplayAndWaitForXml: timed out.");
                return null;
            }
        }
        finally
        {
            _system!.XmlResultArrived -= xmlHandler;
        }
    }

    /// <summary>
    /// Reads <paramref name="filePath"/> as raw bytes and sends them to the device
    /// via <c>DataManSystem.SendCommand("IMAGE.LOAD", bytes)</c>.
    ///
    /// Confirmed 2026-05-25: DataManSystem has no LoadImage method.
    /// SendCommand(String, Byte[]) is the correct SDK mechanism for binary image upload.
    ///
    /// Returns true if the call succeeded without throwing; false otherwise.
    /// </summary>
    private bool TrySendImageViaCommand(string filePath)
    {
        byte[] imageBytes;
        try
        {
            imageBytes = File.ReadAllBytes(filePath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[VTCCP-D4] File read failed: {ex.Message}");
            return false;
        }

        // Resolve SendCommand(String, Byte[]) via reflection (no compile-time SDK dep on Linux).
        var sendMethod = _system!.GetType()
            .GetMethod("SendCommand", [typeof(string), typeof(byte[])]);

        if (sendMethod is null)
        {
            Console.Error.WriteLine("[VTCCP-D4] SendCommand(String, Byte[]) not found on DataManSystem.");
            return false;
        }

        try
        {
            sendMethod.Invoke(_system, ["IMAGE.LOAD", imageBytes]);
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(
                $"[VTCCP-D4] SendCommand(IMAGE.LOAD) failed: {ex.InnerException?.Message ?? ex.Message}");
            return false;
        }
    }

    // ── Diagnostic: raw SYMBOL.RESULT probe ──────────────────────────────────

    /// <summary>
    /// Diagnostic probe: retrieves the raw DMCC SYMBOL.RESULT response by bypassing
    /// the SDK (which throws InvalidCommandException on GET SYMBOL.RESULT) via a
    /// parallel raw TCP connection — the same technique used for TRIGGER.
    ///
    /// Sequence:
    ///   1. Open raw TCP to DMCC port
    ///   2. Drain welcome banner
    ///   3. SET DMCC.RESULT-FORMAT FULL  (requests the firmware's full internal format)
    ///   4. GET SYMBOL.RESULT            (retrieves the last scan result)
    ///   5. Return raw response string
    ///
    /// Call immediately after a scan completes. Used to determine whether FULL format
    /// exposes ECLevel, DataMaskPattern, ECI, or ImagePolarity beyond what the push
    /// script XML provides. If the firmware's FULL format is richer than the JS-generated
    /// push XML, these fields can be sourced via DMCC rather than HTML scraping.
    ///
    /// Returns the raw response string (may be large XML), or null on failure.
    /// Full response is written to Debug output.
    /// </summary>
    public async Task<string?> GetRawSymbolResultAsync(
        int               timeoutMs = 5_000,
        CancellationToken ct        = default)
    {
        ThrowIfDisposed();
        try
        {
            using var tcp = new TcpClient();
            await tcp.ConnectAsync(_cfg.Host, _cfg.Port, ct);
            using var stream = tcp.GetStream();

            // Drain welcome banner (device sends ~100 bytes on connect).
            try
            {
                using var bannerCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                bannerCts.CancelAfter(400);
                byte[] bannerBuf = new byte[512];
                await stream.ReadAsync(bannerBuf, bannerCts.Token);
            }
            catch (OperationCanceledException) { /* no banner or timeout — OK */ }
            catch { }

            // SET DMCC.RESULT-FORMAT FULL — ask firmware for its full internal result format.
            // This is different from the push script XML which is JS-generated and filtered.
            // Drain the SET response (short status line) before sending GET.
            await stream.WriteAsync(Encoding.ASCII.GetBytes("SET DMCC.RESULT-FORMAT FULL\r\n"), ct);
            string setResp = await ReadRawDmccResponseAsync(stream, 1_000, ct);
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-PROBE] SET DMCC.RESULT-FORMAT FULL → '{setResp.Trim()}'");

            // GET SYMBOL.RESULT — retrieve the last scan's result in the FULL format.
            await stream.WriteAsync(Encoding.ASCII.GetBytes("GET SYMBOL.RESULT\r\n"), ct);
            string result = await ReadRawDmccResponseAsync(stream, timeoutMs, ct);

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-PROBE] GET SYMBOL.RESULT FULL: {result.Length} chars\n" +
                (result.Length <= 8000 ? result : result[..8000] + "\n…[truncated]"));

            return result.Length > 0 ? result : null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-PROBE] GetRawSymbolResult failed: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Reads bytes from a raw DMCC TCP stream until the connection closes or the
    /// per-call timeout elapses. Returns whatever was accumulated.
    /// </summary>
    private static async Task<string> ReadRawDmccResponseAsync(
        NetworkStream     stream,
        int               timeoutMs,
        CancellationToken ct)
    {
        var sb  = new StringBuilder();
        var buf = new byte[65536];
        using var readCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        readCts.CancelAfter(timeoutMs);
        try
        {
            while (true)
            {
                int n = await stream.ReadAsync(buf, readCts.Token);
                if (n <= 0) break;
                sb.Append(Encoding.UTF8.GetString(buf, 0, n));
            }
        }
        catch (OperationCanceledException) { /* timeout — return what we have */ }
        return sb.ToString();
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
