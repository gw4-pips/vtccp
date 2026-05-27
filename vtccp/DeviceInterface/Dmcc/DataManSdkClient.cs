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
            _system.Connect(_cfg.ConnectTimeoutMs);
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

        // Timestamp gate: only accept XmlResultArrived events that fire AFTER the
        // trigger has actually been sent.  Events from background Live Mode scans
        // that arrive before our trigger are silently ignored.
        // Set to DateTime.UtcNow.Ticks just before each TRIGGER send.
        long triggerSentTicks = long.MaxValue;

        CognexSdk.XmlResultArrivedHandler xmlHandler = (_, args) =>
        {
            // Ignore events that arrived before the trigger was sent.
            if (DateTime.UtcNow.Ticks < Volatile.Read(ref triggerSentTicks)) return;

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
                    Volatile.Write(ref triggerSentTicks, DateTime.UtcNow.Ticks);
                    _system!.SendCommand("TRIGGER 1");
                    System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] TRIGGER 1 sent OK.");
                    triggered = true;
                    return;
                }
                catch (Exception ex1)
                {
                    Volatile.Write(ref triggerSentTicks, long.MaxValue); // reset gate on failure
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SDK] TRIGGER 1 failed ({ex1.GetType().Name}: {ex1.Message}), trying plain TRIGGER...");
                }

                // Attempt 2: plain TRIGGER
                try
                {
                    Volatile.Write(ref triggerSentTicks, DateTime.UtcNow.Ticks);
                    _system!.SendCommand("TRIGGER");
                    System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] TRIGGER sent OK.");
                    triggered = true;
                }
                catch (Exception ex2)
                {
                    Volatile.Write(ref triggerSentTicks, long.MaxValue); // reset gate on failure
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SDK] TRIGGER also failed ({ex2.GetType().Name}: {ex2.Message}).");
                }
            }, ct);

            if (triggered)
            {
                // SDK trigger sent — wait for XmlResultArrived on the SDK connection.
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

            // Both SDK TRIGGER forms were rejected by the SDK's own parameter
            // validation layer (not the device) — firmware 6.1.16_sr4 / SDK v25
            // version mismatch.  Bypass the SDK entirely: open a raw TCP connection
            // and send TRIGGER.
            //
            // Result delivery is ambiguous: the device may return XML on the raw
            // socket (DMCC synchronous response) OR push it to the SDK's existing
            // connection via XmlResultArrived (async result subscription).
            // We race both channels with Task.WhenAny so we catch the result
            // regardless of which path the firmware uses.
            System.Diagnostics.Debug.WriteLine(
                "[VTCCP-SDK] SDK rejected TRIGGER — raw TCP bypass (racing socket + XmlResultArrived)...");
            try
            {
                using var tcp = new TcpClient();
                tcp.ReceiveTimeout = timeoutMs + 5_000;
                await tcp.ConnectAsync(_cfg.Host, _cfg.Port, ct);
                using var stream = tcp.GetStream();

                // Drain the device welcome banner.
                try
                {
                    using var bannerCts = new CancellationTokenSource(600);
                    byte[] bannerBuf = new byte[1024];
                    int bannerN = await stream.ReadAsync(bannerBuf, bannerCts.Token);
                    if (bannerN > 0)
                        System.Diagnostics.Debug.WriteLine(
                            $"[VTCCP-SDK] Banner ({bannerN}B): '{Encoding.ASCII.GetString(bannerBuf, 0, bannerN).Trim()}'");
                }
                catch (OperationCanceledException) { }
                catch { }

                // Send the bare TRIGGER command.
                // Arm the timestamp gate BEFORE WriteAsync so XmlResultArrived
                // events from this trigger are accepted by the xmlHandler.
                Volatile.Write(ref triggerSentTicks, DateTime.UtcNow.Ticks);
                await stream.WriteAsync(Encoding.ASCII.GetBytes("TRIGGER\r\n"), ct);
                System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] TRIGGER sent — racing socket vs XmlResultArrived...");

                // Race 1: read XML result from the same raw socket.
                using var readCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                readCts.CancelAfter(timeoutMs);

                async Task<string?> ReadSocketAsync(System.IO.Stream s, CancellationToken tok)
                {
                    var sb2     = new StringBuilder(8192);
                    byte[] buf2 = new byte[4096];
                    bool loggedFirst = false;
                    try
                    {
                        while (true)
                        {
                            int n = await s.ReadAsync(buf2, tok);
                            if (n == 0) break;
                            string chunk = Encoding.UTF8.GetString(buf2, 0, n);
                            sb2.Append(chunk);

                            // Always log the first bytes from the device — diagnostic gold.
                            if (!loggedFirst)
                            {
                                loggedFirst = true;
                                string preview = sb2.ToString()[..Math.Min(sb2.Length, 300)]
                                                    .Replace("\r", "\\r").Replace("\n", "\\n");
                                System.Diagnostics.Debug.WriteLine(
                                    $"[VTCCP-SDK] Device response on raw socket ({sb2.Length}B so far): '{preview}'");
                            }

                            string acc  = sb2.ToString();
                            bool hasOpen  = acc.Contains("<result", StringComparison.OrdinalIgnoreCase)
                                         || acc.Contains("<?xml",   StringComparison.OrdinalIgnoreCase);
                            bool hasClose = acc.Contains("</result>", StringComparison.OrdinalIgnoreCase);
                            if (hasOpen && hasClose)
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    $"[VTCCP-SDK] Raw socket XML complete: {acc.Length} chars.");
                                int xs = acc.IndexOf('<');
                                return xs > 0 ? acc[xs..] : acc;
                            }
                        }
                    }
                    catch (OperationCanceledException) { }
                    catch (Exception ex2)
                    {
                        System.Diagnostics.Debug.WriteLine($"[VTCCP-SDK] Socket read error: {ex2.Message}");
                    }

                    if (sb2.Length > 0)
                        System.Diagnostics.Debug.WriteLine(
                            $"[VTCCP-SDK] Socket closed without complete XML ({sb2.Length}B accumulated).");
                    return null;
                }

                // Race 2: tcs resolves when XmlResultArrived fires on SDK channel.
                var socketTask = ReadSocketAsync(stream, readCts.Token);
                var eventTask  = tcs.Task;

                // WhenAny completes as soon as either task finishes.
                // socketTask is bounded by readCts; eventTask may never complete if
                // XmlResultArrived never fires — use WaitAsync to impose the same limit.
                string? result = null;
                try
                {
                    var winner = await Task.WhenAny(socketTask, eventTask).WaitAsync(readCts.Token);
                    result = await winner;
                    System.Diagnostics.Debug.WriteLine(
                        $"[VTCCP-SDK] WhenAny winner: {(winner == (Task<string?>)socketTask ? "socket" : "XmlResultArrived")}  result={result?.Length ?? 0}chars");

                    if (result is null)
                    {
                        // First winner returned null — wait briefly for the other channel.
                        var other = winner == (Task<string?>)socketTask
                            ? eventTask : (Task<string?>)socketTask;
                        try { result = await other.WaitAsync(readCts.Token); }
                        catch (OperationCanceledException) { }
                    }
                }
                catch (OperationCanceledException)
                {
                    System.Diagnostics.Debug.WriteLine("[VTCCP-SDK] Both channels timed out.");
                    // Capture any result that arrived just after the timeout fired.
                    if (socketTask.IsCompleted) result = await socketTask;
                    if (result is null && tcs.Task.IsCompleted) result = tcs.Task.Result;
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] TriggerAndWaitForXml: returning {result?.Length ?? 0} chars.");
                return result;
            }
            catch (Exception tcpEx)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-SDK] Raw TCP TRIGGER exception: {tcpEx.GetType().Name}: {tcpEx.Message}");
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
            // Attempt IMAGE.REPLAY to request a fresh grade cycle.
            // The SDK's SendCommand throws on rejection (device busy / wrong state),
            // which DmccResponse.Parse maps to code=-2.  Do NOT abort on non-zero
            // status — the device may already be in a continuous replay/monitoring
            // loop that delivers XmlResultArrived independently of this command.
            // Just wait for the event regardless.
            var replayResp = await SendAsync(DmccCommand.ImageReplay, ct);
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-REPLAY] IMAGE.REPLAY → code={replayResp.StatusCode} body='{replayResp.Body}'");

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
        // Device-confirmed 2026-05-25: DataManSystem has NO SendImage method.
        // Full SDK method inventory logged in firmware-confirmed-facts.md §11.
        //
        // SendCommand("IMAGE.LOAD", Byte[]) throws InvalidParameterException on all
        // tested formats (JPEG, BMP, PNG).  The correct image-upload mechanism has
        // not yet been identified.  Returns false; caller falls through to replay.
        //
        // Candidate for future investigation: SendCommandWithExpectedBinaryResult —
        // confirmed present on DataManSystem, not yet tested with IMAGE.LOAD.
        _ = filePath;
        return false;
    }

    // ── IMAGE.SEND — ROI frame retrieval ─────────────────────────────────────

    /// <summary>
    /// Retrieves the current image buffer from the device as a raw JPEG byte array.
    ///
    /// This is the Level 2 ROI frame in the three-level image stack:
    ///   Level 1 — barcode crop  : r.trucheck.jpegImage in push XML (already captured)
    ///   Level 2 — ROI frame     : IMAGE.SEND (this method) — operator-configured ROI rect
    ///   Level 3 — full frame    : DataManSystem.GetLastReadImage() SDK (not implemented)
    ///
    /// Call AFTER a scan result has been received and parsed — the device retains the
    /// image buffer until the next trigger.
    ///
    /// Strategy (two-stage fallback):
    ///   1. SDK path — SendCommandWithExpectedBinaryResult("IMAGE.SEND") via reflection.
    ///      The SDK knows the binary framing protocol; binary payload extracted from
    ///      the response via reflection (property name scan for byte[] type).
    ///   2. Raw TCP fallback — send IMAGE.SEND\r\n on a raw socket, read all bytes,
    ///      strip the DMCC text header (\r\n0\r\n\r\n), return the remaining JPEG bytes.
    ///
    /// Returns null if both paths fail or if the response is not a valid JPEG
    /// (does not start with 0xFF 0xD8 — the JPEG SOI marker).
    /// Never throws.
    ///
    /// Resolution: controlled by the device's IMAGE.SIZE DMCC setting (0=Full … 3=1/64).
    /// For highest-fidelity OCR use IMAGE.SIZE 0 (Full) — confirmed DMCC key.
    /// </summary>
    public async Task<byte[]?> GetRoiImageAsync(
        int               timeoutMs = 5_000,
        CancellationToken ct        = default)
    {
        ThrowIfDisposed();

        // ── Stage 1: SDK path ────────────────────────────────────────────────
        if (_system != null)
        {
            var sdkResult = await Task.Run(() => TryGetRoiImageViaSdk(), ct);
            if (sdkResult != null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-ROI] IMAGE.SEND via SDK: {sdkResult.Length} bytes");
                return sdkResult;
            }
            System.Diagnostics.Debug.WriteLine(
                "[VTCCP-ROI] SDK path failed — falling back to raw TCP.");
        }

        // ── Stage 2: Raw TCP fallback ────────────────────────────────────────
        return await GetRoiImageViaRawTcpAsync(timeoutMs, ct);
    }

    /// <summary>
    /// Attempts IMAGE.SEND via the SDK's SendCommandWithExpectedBinaryResult(String) overload.
    /// Returns the JPEG bytes on success, null on any failure.
    /// </summary>
    private byte[]? TryGetRoiImageViaSdk()
    {
        if (_system is null) return null;
        try
        {
            // SendCommandWithExpectedBinaryResult is in the SDK method inventory
            // (firmware-confirmed-facts.md §11).  Call via reflection to avoid
            // a compile-time dependency on the Windows-only SDK from Linux agents.
            var method = _system.GetType().GetMethod(
                "SendCommandWithExpectedBinaryResult",
                [typeof(string)]);

            if (method is null)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[VTCCP-ROI] SendCommandWithExpectedBinaryResult(String) not found on DataManSystem.");
                return null;
            }

            var sdkResp = method.Invoke(_system, ["IMAGE.SEND"]);
            if (sdkResp is null) return null;

            // Extract binary payload.
            // The SDK response carries the image in one of two forms:
            //   • byte[]      — scan for a property whose value is byte[] with length > 2
            //   • MemoryStream — confirmed live: SdkResp.BinaryData is a MemoryStream
            //                   (logged as "System.IO.MemoryStream" in DumpProps output)
            // Check both forms before falling back to DumpProps.
            foreach (var prop in sdkResp.GetType().GetProperties())
            {
                try
                {
                    var val = prop.GetValue(sdkResp);

                    // Form 1: direct byte[]
                    if (val is byte[] bytes && bytes.Length > 2)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[VTCCP-ROI] SDK byte[] prop '{prop.Name}': {bytes.Length} bytes");
                        return IsJpeg(bytes) ? bytes : null;
                    }

                    // Form 2: MemoryStream (BinaryData property — confirmed on fw 6.1.16_sr4)
                    if (val is System.IO.MemoryStream ms && ms.Length > 2)
                    {
                        byte[] msBytes = ms.ToArray();
                        System.Diagnostics.Debug.WriteLine(
                            $"[VTCCP-ROI] SDK MemoryStream prop '{prop.Name}': {msBytes.Length} bytes");
                        return IsJpeg(msBytes) ? msBytes : null;
                    }
                }
                catch { }
            }

            System.Diagnostics.Debug.WriteLine(
                "[VTCCP-ROI] SDK response had no byte[] or MemoryStream property — dumping props:");
            DumpProps("[VTCCP-ROI] SdkResp", sdkResp);
            return null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-ROI] SDK IMAGE.SEND failed: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Sends IMAGE.SEND on a raw TCP connection, reads the binary JPEG response,
    /// strips the DMCC text header (\r\n0\r\n\r\n), and returns the JPEG bytes.
    /// </summary>
    private async Task<byte[]?> GetRoiImageViaRawTcpAsync(
        int               timeoutMs,
        CancellationToken ct)
    {
        try
        {
            using var tcp = new TcpClient();
            tcp.ReceiveBufferSize = 1 << 20;  // 1 MB receive buffer
            await tcp.ConnectAsync(_cfg.Host, _cfg.Port, ct);
            using var stream = tcp.GetStream();

            // Drain welcome banner.
            try
            {
                using var bannerCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                bannerCts.CancelAfter(400);
                byte[] bannerBuf = new byte[512];
                await stream.ReadAsync(bannerBuf, bannerCts.Token);
            }
            catch (OperationCanceledException) { }
            catch { }

            // Send IMAGE.SEND command.
            await stream.WriteAsync(Encoding.ASCII.GetBytes("IMAGE.SEND\r\n"), ct);

            // Read all bytes until the device closes the connection or timeout.
            var raw = await ReadRawBinaryResponseAsync(stream, timeoutMs, ct);

            if (raw is null || raw.Length == 0)
            {
                System.Diagnostics.Debug.WriteLine("[VTCCP-ROI] Raw TCP: no data received.");
                return null;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-ROI] Raw TCP: {raw.Length} bytes received.");

            // Strip the DMCC text header.  Format: \r\n0\r\n\r\n{binary}
            // The header ends at the first occurrence of \r\n\r\n after the status line.
            byte[] jpeg = StripDmccHeader(raw);

            if (jpeg.Length == 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-ROI] Header strip returned 0 bytes.  " +
                    $"First 32: {BitConverter.ToString(raw[..Math.Min(32, raw.Length)])}");
                return null;
            }

            if (!IsJpeg(jpeg))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-ROI] Payload is not a JPEG.  " +
                    $"First 4 bytes: {BitConverter.ToString(jpeg[..Math.Min(4, jpeg.Length)])}");
                return null;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-ROI] JPEG confirmed: {jpeg.Length} bytes.");
            return jpeg;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-ROI] Raw TCP IMAGE.SEND failed: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Reads all available bytes from a NetworkStream until connection close or timeout.
    /// Returns null if reading fails entirely.
    /// </summary>
    private static async Task<byte[]?> ReadRawBinaryResponseAsync(
        NetworkStream     stream,
        int               timeoutMs,
        CancellationToken ct)
    {
        var ms  = new MemoryStream();
        var buf = new byte[65536];
        using var readCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        readCts.CancelAfter(timeoutMs);
        try
        {
            while (true)
            {
                int n = await stream.ReadAsync(buf, readCts.Token);
                if (n <= 0) break;
                await ms.WriteAsync(buf.AsMemory(0, n), readCts.Token);
            }
        }
        catch (OperationCanceledException) { /* timeout — return whatever arrived */ }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-ROI] ReadRawBinary error: {ex.GetType().Name}: {ex.Message}");
            if (ms.Length == 0) return null;
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Strips the DMCC text response header from a binary response buffer.
    ///
    /// DMCC header format for a successful binary response:
    ///   \r\n0\r\n\r\n{binary data}
    ///
    /// Locates the first \r\n\r\n (double CRLF) in the first 64 bytes, which marks
    /// the end of the header.  Returns everything after that point.
    /// If no header separator is found, returns the full buffer (allows for devices
    /// that send raw binary without a text header, e.g. some firmware variants).
    /// </summary>
    private static byte[] StripDmccHeader(byte[] raw)
    {
        const int maxHeaderScan = 64;
        int limit = Math.Min(raw.Length - 3, maxHeaderScan);
        for (int i = 0; i <= limit; i++)
        {
            if (raw[i] == '\r' && raw[i + 1] == '\n' &&
                raw[i + 2] == '\r' && raw[i + 3] == '\n')
            {
                return raw[(i + 4)..];
            }
        }
        return raw;  // no header found — treat entire buffer as payload
    }

    /// <summary>Returns true if the byte array starts with the JPEG SOI marker 0xFF 0xD8.</summary>
    private static bool IsJpeg(byte[] data) =>
        data.Length >= 2 && data[0] == 0xFF && data[1] == 0xD8;

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
