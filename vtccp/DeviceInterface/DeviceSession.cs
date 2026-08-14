namespace DeviceInterface;

using DeviceInterface.Dmcc;
using DeviceInterface.Dmst;
using ExcelEngine.Models;

/// <summary>
/// High-level orchestration of a live Cognex DataMan DMV device session.
///
/// Lifecycle:
///   1. Construct with <see cref="DeviceConfig"/> and optional <see cref="VerificationXmlMap"/>.
///   2. await <see cref="ConnectAsync"/> — opens DMCC connection, queries device info.
///   3. Read <see cref="DeviceInfo"/> to populate <see cref="SessionState"/> fields.
///   4. Loop: await <see cref="TriggerAndGetResultAsync"/> (Poll mode) for each scan.
///      Or: attach <see cref="ResultReceived"/> event and call <see cref="StartPushListenerAsync"/>
///      for Push (DMST) mode.
///   5. await <see cref="DisconnectAsync"/> / DisposeAsync() when done.
///
/// All async methods accept a <see cref="CancellationToken"/> for timeout / cancellation.
/// </summary>
public sealed class DeviceSession : IAsyncDisposable
{
    private readonly DeviceConfig       _cfg;
    private readonly VerificationXmlMap _map;
    private readonly DataManSdkClient   _client;
    private DmstListener?               _listener;
    private DmstHtmlScraper?            _scraper;
    private HttpEventSubscriber?        _httpSubscriber;
    private bool                        _disposed;

    /// <summary>
    /// Original trigger type read from the device in ConnectAsync.
    /// Restored in DisconnectAsync so VTCCP does not permanently alter device settings.
    /// Null if GET TRIGGER.TYPE was not supported by this firmware.
    /// </summary>
    private string? _originalTriggerType;

    /// <summary>
    /// The TRIGGER.TYPE value that was on the device when VTCCP connected — i.e. the
    /// value DMST had left it at.  Exposed so the UI can display it to the operator.
    /// Null until ConnectAsync has completed.
    /// </summary>
    public string? OriginalTriggerType => _originalTriggerType;

    /// <summary>
    /// Source file path of the Webscan HTML report that was most recently correlated
    /// by the internal <see cref="DmstHtmlScraper"/>.  Set each time
    /// <see cref="DmstHtmlScraper.TryMergeAsync"/> finds a match.
    ///
    /// Used by <c>SessionViewModel</c> in Replace mode: after
    /// <see cref="TriggerAndGetResultAsync"/> returns, this property holds the path
    /// of the original Webscan HTML (already deleted by the scraper) so the caller
    /// can write the hybrid report to the same location.
    ///
    /// Null when no scraper is running (e.g. DMST extension is .pdf, not .html).
    /// </summary>
    public string? LastMatchedSourcePath => _scraper?.LastMatchedSourcePath;

    /// <summary>
    /// Registers <paramref name="path"/> with the internal <see cref="DmstHtmlScraper"/> so
    /// that the FileSystemWatcher suppresses the Created event triggered by VTCCP's own write
    /// of a hybrid report to the CodeQuality folder (Replace mode).
    ///
    /// Call this immediately before writing the hybrid file to prevent the scraper from
    /// re-parsing and deleting the newly written report.  No-op when no scraper is running.
    /// </summary>
    public void RegisterOwnedHybridPath(string path) => _scraper?.RegisterOwnedPath(path);

    /// <summary>
    /// Configures whether the internal <see cref="DmstHtmlScraper"/> deletes Webscan HTML
    /// files from the CodeQuality folder after parsing them.
    ///
    /// Pass <c>true</c> (Replace mode) to delete after parse — the hybrid report will be
    /// written back to the same path.  Pass <c>false</c> (Alongside mode) to leave the
    /// original Webscan HTML on disk so both files coexist.
    ///
    /// Must be called after <see cref="ConnectAsync"/> but before the first scan.
    /// No-op when no scraper is running.
    /// </summary>
    public void ConfigureScraperDeletion(bool deleteAfterParse)
    {
        if (_scraper is not null)
            _scraper.DeleteAfterParse = deleteAfterParse;
    }

    /// <summary>
    /// Device information queried during <see cref="ConnectAsync"/>.
    /// Populated fields: Type, Serial, Name, FirmwareVersion, CalibrationDate.
    /// Use these to pre-fill <see cref="SessionState"/> before opening an Excel session.
    /// </summary>
    public DeviceInfo DeviceInfo { get; private set; } = new();

    /// <summary>Raised after each result is received in Push mode.</summary>
    public event EventHandler<VerificationRecord>? ResultReceived;

    public DeviceSession(DeviceConfig config, VerificationXmlMap? map = null)
    {
        _cfg    = config ?? throw new ArgumentNullException(nameof(config));
        _map    = map ?? new VerificationXmlMap();
        _client = new DataManSdkClient(_cfg);
    }

    // ── Connection ────────────────────────────────────────────────────────────

    /// <summary>
    /// Opens the DMCC TCP connection and queries device metadata.
    /// Throws <see cref="DeviceConnectionException"/> if the connection cannot be established.
    /// </summary>
    public async Task ConnectAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        System.Diagnostics.Debug.WriteLine("[VTCCP-BUILD] DeviceSession 2026-05-28 — SET TRIGGER.TYPE disabled. HTTP subscriber activated.");

        for (int attempt = 1; ; attempt++)
        {
            try
            {
                await _client.ConnectAsync(ct);
                break;
            }
            catch (Exception ex) when (attempt < _cfg.MaxReconnectAttempts)
            {
                await Task.Delay(_cfg.ReconnectDelayMs, ct);
                _ = ex; // suppress unused warning
            }
            catch (Exception ex)
            {
                throw new DeviceConnectionException(
                    $"Cannot connect to {_cfg.Host}:{_cfg.Port} after {_cfg.MaxReconnectAttempts} attempts.", ex);
            }
        }

        // SetResultTypes() is intentionally NOT called (see DataManSdkClient class comment).
        // VTCCP receives all results via HttpEventSubscriber — not the DMCC result channel.

        // Query device identity info.
        // FirmwareVersion is read from the SDK's native property first (avoids
        // latency on the hot path); DMCC fallback uses the confirmed key
        // DEVICE.FIRMWARE-VER (verified against DMCC Reference 6.1.16_sr4).
        // Pre-fetch device type first so we can drive the sensor lookup table
        // without a second DMCC round-trip inside the DeviceInfo initializer.
        var devType      = (await _client.SendAsync(DmccCommand.GetDeviceType,  ct)).Body;
        var imageSizeRaw = (await _client.SendAsync(DmccCommand.GetImageSize,   ct)).Body;
        var sensorSpec   = DeviceSensorSpecs.TryGet(devType);

        DeviceInfo = new DeviceInfo
        {
            Type               = devType,
            FirmwareVersion    = _client.FirmwareVersion
                              ?? (await _client.SendAsync(DmccCommand.GetFirmwareVer,        ct)).Body,
            Name               = (await _client.SendAsync(DmccCommand.GetDeviceName,         ct)).Body,
            Serial             = (await _client.SendAsync(DmccCommand.GetDeviceSerialNumber, ct)).Body,
            SensorWidthPx      = sensorSpec?.WidthPx,
            SensorHeightPx     = sensorSpec?.HeightPx,
            SensorPixelPitchUm = sensorSpec?.PixelPitchUm,
            ImageSizeSetting   = imageSizeRaw switch {
                "0" => "Full", "1" => "1/4", "2" => "1/16", "3" => "1/64",
                _   => imageSizeRaw,   // preserve raw value for unknown / future firmware
            },
        };

        // TRIGGER.TYPE confirmed values from DMCC Reference fw 6.1.16_sr4:
        //   0=Single(external)  1=Presentation(internal)  2=Manual(button)
        //   3=Burst(external)   4=Self(internal)           5=Continuous(external)
        //
        // GET is kept as a diagnostic read so the UI and logs can show the
        // device's current trigger mode on connect.
        //
        // SET TRIGGER.TYPE is intentionally NOT issued here.
        // Rationale: the raw TCP TRIGGER command fires a single acquisition
        // regardless of TRIGGER.TYPE, and DMST itself never changes TRIGGER.TYPE
        // during normal Go-Live / Verify operation (it stays at Single/external).
        // The prior SET TRIGGER.TYPE 1 was based on a wrong assumption and
        // may have caused undesirable post-scan reader behaviour — under probe.
        var trigResp = await _client.SendAsync(DmccCommand.GetTriggerType, ct);
        _originalTriggerType = string.IsNullOrWhiteSpace(trigResp.Body)
            ? null
            : trigResp.Body.Trim();
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-DMCC] GET TRIGGER.TYPE: code={trigResp.StatusCode}  value='{_originalTriggerType}'");

        // SET TRIGGER.TYPE 1 — COMMENTED OUT (probe: was this causing post-scan looping?)
        // var setResp = await _client.SendAsync(DmccCommand.SetTriggerTypeSingle, ct);
        // System.Diagnostics.Debug.WriteLine(
        //     $"[VTCCP-DMCC] SET TRIGGER.TYPE 1 (Single): code={setResp.StatusCode}");

        // ── HTML report scraper ───────────────────────────────────────────────
        // Watches the DMST CodeQuality directory for HTML reports written by DMST
        // after each verification scan.  These reports supply supplemental fields
        // absent from the push XML on fw 6.1.16_sr4:
        //   • EncodedCharacters (correct value — push XML is wrong)
        //   • DataCodewords, ErrorCorrectionBudget
        //   • ImagePolarity
        //   • ECLevel, DataMaskPattern, ECI  (QR only)
        //
        // Prerequisite: DMST Options → Data Logging → Reporting →
        //   "Preferred Quality Report File Extension" must be set to ".html".
        //
        // Safe no-op when DMST is not running — TryMergeAsync times out in 4 s
        // and returns the push-XML record unmodified, with no exception thrown.
        if (DeviceInfo.Name is { Length: > 0 } deviceName)
        {
            var reportPath = DmstHtmlScraper.BuildReportPath(deviceName);
            _scraper = new DmstHtmlScraper(reportPath);
            _scraper.Start();
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-SCRAPER] Started watching '{reportPath}'.");
        }
    }

    /// <summary>Closes the DMCC connection and stops any active push listener.</summary>
    public async Task DisconnectAsync()
    {
        // TRIGGER.TYPE restore — COMMENTED OUT (SET on connect is also commented out;
        // nothing to restore.  Re-enable both together if SET is reinstated.)
        // if (_originalTriggerType is not null && _client.IsConnected)
        // {
        //     try
        //     {
        //         var restoreCmd = $"SET TRIGGER.TYPE {_originalTriggerType}";
        //         await _client.SendAsync(restoreCmd);
        //         System.Diagnostics.Debug.WriteLine(
        //             $"[VTCCP-DMCC] Restored trigger type to '{_originalTriggerType}'.");
        //     }
        //     catch (Exception ex)
        //     {
        //         System.Diagnostics.Debug.WriteLine(
        //             $"[VTCCP-DMCC] Could not restore trigger type: {ex.Message}");
        //     }
        // }

        if (_httpSubscriber is not null)
        {
            await _httpSubscriber.StopAsync();
            _httpSubscriber = null;
        }

        _scraper?.Stop();
        _scraper = null;

        if (_listener is not null) await _listener.StopAsync();
        await _client.DisconnectAsync();
    }

    /// <summary>
    /// Sends REBOOT to the device, then disconnects cleanly.
    /// The device will be unavailable for ~30–60 s while it restarts.
    /// After it comes back up DMST and other clients can reconnect normally.
    /// </summary>
    public async Task RebootAndDisconnectAsync()
    {
        // TRIGGER.TYPE restore — COMMENTED OUT (SET on connect is also commented out;
        // nothing to restore.  Re-enable both together if SET is reinstated.)
        // if (_originalTriggerType is not null && _client.IsConnected)
        // {
        //     try
        //     {
        //         await _client.SendAsync($"SET TRIGGER.TYPE {_originalTriggerType}");
        //         System.Diagnostics.Debug.WriteLine(
        //             $"[VTCCP-DMCC] Restored trigger type to '{_originalTriggerType}'.");
        //     }
        //     catch (Exception ex)
        //     {
        //         System.Diagnostics.Debug.WriteLine(
        //             $"[VTCCP-DMCC] Could not restore trigger type: {ex.Message}");
        //     }
        // }

        // Stop any active subscribers/listeners before the device disappears.
        if (_httpSubscriber is not null) { await _httpSubscriber.StopAsync(); _httpSubscriber = null; }
        _scraper?.Stop();
        _scraper = null;
        if (_listener is not null) await _listener.StopAsync();

        // Issue REBOOT — device starts rebooting immediately; TCP connection will drop.
        if (_client.IsConnected)
        {
            try
            {
                await _client.SendAsync(DmccCommand.Reboot);
                System.Diagnostics.Debug.WriteLine("[VTCCP-DMCC] REBOOT command sent.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-DMCC] REBOOT send error (expected if device closed TCP immediately): {ex.Message}");
            }
        }

        // SDK disconnect — may throw if device already closed the connection; swallow it.
        try { await _client.DisconnectAsync(); }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-DMCC] Disconnect after reboot (benign): {ex.Message}");
        }
    }

    // ── Device configuration (Code Properties) ───────────────────────────────

    /// <summary>
    /// Reads the current UPC/EAN supplemental (add-on) mode from the device.
    /// Returns 0–5 matching the DMST Code Details dropdown, or -1 if the command
    /// is not supported / key name differs on this firmware.
    /// Integer mapping: 0=Ignore, 1=Parse, 2=Required, 3=Required 2-digit,
    ///                  4=Required 5-digit, 5=Not Required.
    /// </summary>
    public async Task<int> GetUpcEanSupplementalAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        var resp = await _client.SendAsync(DmccCommand.GetUpcEanSupplemental, ct);
        return resp.StatusCode == DmccStatus.Ok
               && int.TryParse(resp.Body.Trim(), out int mode)
               && mode is >= 0 and <= 5
            ? mode
            : -1;
    }

    /// <summary>
    /// Writes the UPC/EAN supplemental mode to firmware (persistent — no explicit SAVE needed).
    /// Returns true on success.
    /// </summary>
    public async Task<bool> SetUpcEanSupplementalAsync(int mode, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        var resp = await _client.SendAsync(DmccCommand.SetUpcEanSupplemental(mode), ct);
        return resp.StatusCode == DmccStatus.Ok;
    }

    // ── Poll mode ─────────────────────────────────────────────────────────────

    /// <summary>
    /// (Poll mode) Sends a software TRIGGER, waits for the device to complete
    /// verification, then polls <c>GET SYMBOL.RESULT</c> and returns the parsed record.
    ///
    /// Pass a <paramref name="sessionContext"/> record carrying session-level fields
    /// (OperatorId, JobName, etc.) — they are copied into the returned record.
    ///
    /// Returns null if the device reports "no read" (status 6).
    /// </summary>
    public async Task<VerificationRecord?> TriggerAndGetResultAsync(
        VerificationRecord? sessionContext = null,
        CancellationToken   ct            = default)
    {
        ThrowIfDisposed();
        if (!_client.IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        // Use the SDK's XmlResultArrived event-based path.
        // "GET SYMBOL.RESULT" throws InvalidCommandException in the SDK; results
        // must arrive via event after the trigger fires.
        string? xml = await _client.TriggerAndWaitForXmlAsync(ct: ct);
        if (string.IsNullOrWhiteSpace(xml)) return null;

        var record = DmstResultParser.Parse(xml, _map, sessionContext ?? ContextFromDeviceInfo());
        record = await AttachRoiImageAsync(record, ct);
        if (_scraper is null) return record;
        var (merged, sourcePath) = await _scraper.TryMergeAsync(record, ct);
        return sourcePath is not null ? merged with { WebscanSourcePath = sourcePath } : merged;
    }

    // ── Replay mode (image already loaded) ───────────────────────────────────

    /// <summary>
    /// (Replay mode) Sends IMAGE.REPLAY on the currently-loaded device image buffer
    /// and returns the parsed result.  Does NOT load a new image — the device must
    /// already have an image in its buffer from a prior IMAGE.LOAD or live scan.
    ///
    /// Used by Repeatability Analysis to loop N re-grades on a fixed image without
    /// any physical trigger or image reload between iterations.
    ///
    /// Returns null if the device returns no result within <paramref name="timeoutMs"/>.
    /// </summary>
    public async Task<VerificationRecord?> ReplayAndGetResultAsync(
        VerificationRecord? sessionContext = null,
        int                 timeoutMs      = 15_000,
        CancellationToken   ct             = default)
    {
        ThrowIfDisposed();
        if (!_client.IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        string? xml = await _client.ReplayAndWaitForXmlAsync(timeoutMs, ct);
        if (string.IsNullOrWhiteSpace(xml)) return null;

        var record = DmstResultParser.Parse(xml, _map, sessionContext ?? ContextFromDeviceInfo());
        record = await AttachRoiImageAsync(record, ct);
        if (_scraper is null) return record;
        var (merged, sourcePath) = await _scraper.TryMergeAsync(record, ct);
        return sourcePath is not null ? merged with { WebscanSourcePath = sourcePath } : merged;
    }

    /// <summary>
    /// (Image-load mode) Loads a JPEG/BMP image from <paramref name="imagePath"/> into the
    /// device image buffer using IMAGE.LOAD + IMAGE.REPLAY and returns the parsed
    /// <see cref="VerificationRecord"/>.
    ///
    /// Called by <see cref="VtccpApp.ViewModels.StitchingViewModel"/> after the operator
    /// composites two captured frames into a single stitched symbol image.
    ///
    /// The result is tagged <c>OpticsSource = "StitchedImage"</c> so the report layer
    /// can display the appropriate disclaimer.
    ///
    /// Returns null if the device returns no result within <paramref name="timeoutMs"/>.
    /// Throws <see cref="InvalidOperationException"/> if the SDK is not connected.
    /// </summary>
    public async Task<VerificationRecord?> LoadImageAndVerifyAsync(
        string            imagePath,
        int               timeoutMs = 30_000,
        CancellationToken ct        = default)
    {
        ThrowIfDisposed();
        if (!_client.IsConnected)
            throw new InvalidOperationException(
                "SDK not connected — cannot use IMAGE.LOAD. Connect the device first.");

        string? xml = await _client.LoadAndReplayImageAsync(imagePath, timeoutMs, ct);
        if (string.IsNullOrWhiteSpace(xml)) return null;

        var ctx    = ContextFromDeviceInfo();
        var record = DmstResultParser.Parse(xml, _map, ctx);

        record = record with { OpticsSource = "StitchedImage" };
        record = await AttachRoiImageAsync(record, ct);
        if (_scraper is null) return record;
        var (merged, sourcePath) = await _scraper.TryMergeAsync(record, ct);
        return sourcePath is not null ? merged with { WebscanSourcePath = sourcePath } : merged;
    }

    // ── Push mode ─────────────────────────────────────────────────────────────

    /// <summary>
    /// (Push mode) Starts a TCP listener on <see cref="DeviceConfig.DmstListenPort"/>.
    /// When the device pushes a DMST XML result, it is parsed and <see cref="ResultReceived"/>
    /// is raised on the thread-pool.
    /// </summary>
    public async Task StartPushListenerAsync(
        VerificationRecord? sessionContext = null,
        CancellationToken   ct            = default)
    {
        ThrowIfDisposed();
        if (_listener is not null) return; // already running

        _listener = new DmstListener(_cfg.DmstListenPort, _map,
            sessionContext ?? ContextFromDeviceInfo(),
            rec =>
            {
                // In push mode, TryMergeAsync waits up to 4 s for the DMST HTML
                // file.  The DmstListener callback is synchronous, so we fire the
                // async merge on the thread-pool and raise ResultReceived only after
                // the merge (or timeout) completes.  Scans are never rapid-fire in
                // TruCheck mode, so ordering is preserved in practice.
                var scraper = _scraper;
                if (scraper is null)
                {
                    ResultReceived?.Invoke(this, rec);
                    return;
                }
                _ = Task.Run(async () =>
                {
                    var (enriched, sourcePath) = await scraper.TryMergeAsync(rec);
                    var result = sourcePath is not null
                        ? enriched with { WebscanSourcePath = sourcePath }
                        : enriched;
                    ResultReceived?.Invoke(this, result);
                });
            });

        await _listener.StartAsync(ct);
    }

    /// <summary>Stops the push listener if active.</summary>
    public async Task StopPushListenerAsync()
    {
        if (_listener is null) return;
        await _listener.StopAsync();
        _listener = null;
    }

    // ── HTTP event subscriber mode ────────────────────────────────────────────

    /// <summary>
    /// (HTTP subscriber mode) Opens a raw TCP connection to the device on the DMCC
    /// port and subscribes to the device's HTTP event push stream.
    ///
    /// This mode delivers complete scan results — including all HTML supplemental
    /// fields (ECLevel, DataMaskPattern, ECI, ImagePolarity, DataCodewords,
    /// ErrorCorrectionBudget, EncodedCharacters) — with zero DMST dependency.
    /// The device pushes results directly to VTCCP over the same port 44444 used
    /// for DMCC, multiplexed by connection intent.
    ///
    /// Result flow (all on one Keep-Alive TCP connection):
    ///   PUT /pcm_report.html  →  DmstHtmlScraper.ParseHtml()  →  _pendingHtml
    ///   PUT /codes.xml        →  DmstResultParser.Parse() + MergeAndValidate()
    ///                         →  ResultReceived raised on thread pool
    ///
    /// When this subscriber is running, <see cref="StartPushListenerAsync"/> and
    /// <see cref="TriggerAndGetResultAsync"/> are not needed for result delivery.
    /// The HTML scraper (<see cref="DmstHtmlScraper"/>) is harmless alongside it
    /// (it simply watches the filesystem; the subscriber enriches records inline).
    ///
    /// Safe to call if already running — returns immediately.
    /// </summary>
    public async Task StartHttpSubscriberAsync(
        VerificationRecord? sessionContext = null,
        CancellationToken   ct            = default)
    {
        ThrowIfDisposed();
        if (_httpSubscriber?.IsRunning == true) return;

        _httpSubscriber = new HttpEventSubscriber(
            _cfg.Host,
            _cfg.Port,
            _map,
            sessionContext ?? ContextFromDeviceInfo(),
            rec =>
            {
                // Mirror StartPushListenerAsync: when the file-system scraper is running
                // (Manual/AutoPoll mode with DMST open), call TryMergeAsync so the pending
                // HTML report is drained from the scraper's queue and its source path is
                // propagated onto the record.  In Replace mode this path is the write
                // target for the hybrid report; in Alongside mode the scraper leaves the
                // original file on disk (DeleteAfterParse = false) so both coexist.
                // HttpEventSubscriber already merges HTML data inline from PUT /pcm_report.html;
                // TryMergeAsync here simply adds the filesystem source path, with a short
                // effective wait because the HTML is already in _pending by the time
                // PUT /codes.xml fires the callback.
                var scraper = _scraper;
                if (scraper is null)
                {
                    ResultReceived?.Invoke(this, rec);
                    return;
                }
                _ = Task.Run(async () =>
                {
                    var (enriched, sourcePath) = await scraper.TryMergeAsync(rec);
                    var result = sourcePath is not null
                        ? enriched with { WebscanSourcePath = sourcePath }
                        : enriched;
                    ResultReceived?.Invoke(this, result);
                });
            });

        await _httpSubscriber.StartAsync(ct);
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-SESSION] HTTP subscriber started on {_cfg.Host}:{_cfg.Port}.");
    }

    /// <summary>Stops the HTTP event subscriber if active.</summary>
    public async Task StopHttpSubscriberAsync()
    {
        if (_httpSubscriber is null) return;
        await _httpSubscriber.StopAsync();
        _httpSubscriber = null;
    }

    // ── ROI image retrieval ───────────────────────────────────────────────────

    /// <summary>
    /// Retrieves the device's current image buffer as raw JPEG bytes via IMAGE.SEND.
    ///
    /// For live scans: returns the operator-configured ROI frame — wider than the
    /// barcode crop in push XML, includes surrounding label area (HRI, lot numbers, etc.).
    ///
    /// For loaded-image replays: returns the loaded image bytes (identical to the
    /// source image that was sent to the device — useful as the OCR source for
    /// externally-sourced symbols where the original file already contains label context).
    ///
    /// Returns null if IMAGE.SEND failed or the response is not a valid JPEG.
    ///
    /// Callers that need the base64 string: use Convert.ToBase64String on the result.
    /// </summary>
    public Task<byte[]?> GetRoiImageAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (!_client.IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");
        return _client.GetRoiImageAsync(ct: ct);
    }

    /// <summary>
    /// Fetches the ROI image via IMAGE.SEND and attaches it to <paramref name="record"/>
    /// as <c>RoiJpegImageBase64</c>.  If IMAGE.SEND fails, returns the original record unchanged.
    /// </summary>
    private async Task<VerificationRecord> AttachRoiImageAsync(
        VerificationRecord record,
        CancellationToken  ct)
    {
        byte[]? roiBytes = await _client.GetRoiImageAsync(ct: ct);
        if (roiBytes is null || roiBytes.Length == 0)
            return record;

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-ROI] Attached to record: {roiBytes.Length} bytes  " +
            $"Symbology={record.Symbology}  Grade={record.OverallGrade}");

        return record with { RoiJpegImageBase64 = Convert.ToBase64String(roiBytes) };
    }

    // ── D4 Image Load ─────────────────────────────────────────────────────────

    /// <summary>
    /// (D4 Image Load) Loads a locally-stored symbol image into the device's image
    /// buffer and fires <c>IMAGE.REPLAY</c> to trigger a full TruCheck verification
    /// pass on the loaded image.  Waits for and returns the parsed result record.
    ///
    /// The device performs the same ISO grading pipeline as a live scan, using the
    /// stored pixels rather than a live camera frame.  OpticsSource is always
    /// "LoadedImage" (CU=-1 / MRD=-1 discriminator, confirmed scan #13, 2026-05-24).
    ///
    /// Does NOT invoke TryMergeAsync — the HTML report written by DMST for a
    /// loaded-image replay is not available on the filesystem in DMST-less mode.
    /// If the HTTP subscriber is running, HTML supplemental fields arrive inline
    /// from pcm_report.html and are merged automatically before ResultReceived fires.
    ///
    /// Use cases:
    ///   • Re-verify a previously captured symbol image
    ///   • Batch grading from an external image library (see Batch Upload feature log)
    ///   • Repeatability analysis on a fixed stored image
    ///
    /// Returns null on load failure, IMAGE.REPLAY rejection, or result timeout.
    /// </summary>
    public async Task<VerificationRecord?> LoadAndVerifyImageAsync(
        string              imagePath,
        VerificationRecord? sessionContext = null,
        CancellationToken   ct            = default)
    {
        ThrowIfDisposed();
        if (!_client.IsConnected)
            throw new InvalidOperationException("Not connected. Call ConnectAsync() first.");

        string? xml = await _client.LoadAndReplayImageAsync(imagePath, ct: ct);
        if (string.IsNullOrWhiteSpace(xml)) return null;

        var record = DmstResultParser.Parse(xml, _map, sessionContext ?? ContextFromDeviceInfo());

        // Force OpticsSource to "LoadedImage" — the parser derives this from
        // CU=-1/MRD=-1, which is correct on fw 6.1.16_sr4, but make it explicit
        // so callers can rely on it without knowledge of the discrimination logic.
        if (record.OpticsSource != "LoadedImage")
            record = record with { OpticsSource = "LoadedImage" };

        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-D4] LoadAndVerifyImageAsync complete: " +
            $"OpticsSource={record.OpticsSource}  Grade={record.OverallGrade}  " +
            $"Symbology={record.Symbology}  Path='{imagePath}'");

        return record;
    }

    // ── Diagnostic probes ────────────────────────────────────────────────────

    /// <summary>
    /// Diagnostic probe: retrieves the raw DMCC SYMBOL.RESULT in FULL format via a
    /// raw TCP bypass connection (the SDK throws InvalidCommandException on GET SYMBOL.RESULT).
    ///
    /// Call this immediately after a scan completes (push XML or poll result received)
    /// to capture the firmware's last result in its native FULL format.
    ///
    /// Purpose: determine whether DMCC.RESULT-FORMAT FULL exposes ECLevel,
    /// DataMaskPattern, ECI, or ImagePolarity beyond the push script XML, which our
    /// 15-scan probe campaign (v1.29–v1.33) confirmed does NOT contain these fields.
    /// If the FULL-format XML is richer, these fields can be sourced via DMCC.
    /// If identical, HTML report scraping is the only remaining path.
    ///
    /// Output is written to Debug and returned as raw string (may be large XML).
    /// Returns null on connection failure or empty response.
    /// </summary>
    public async Task<string?> GetRawSymbolResultDiagnosticAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        return await _client.GetRawSymbolResultAsync(ct: ct);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Builds a thin context record from device info and connection config for
    /// pre-seeding result records.  All fields here are static for the lifetime
    /// of the session — they are captured once at ConnectAsync and stamped on
    /// every scan record at zero marginal cost.
    /// </summary>
    private VerificationRecord ContextFromDeviceInfo() => new()
    {
        Symbology          = "Unknown",  // filled in by parser
        DeviceSerial       = DeviceInfo.Serial,
        DeviceName         = DeviceInfo.Name,
        DeviceModel        = DeviceInfo.Type,
        FirmwareVersion    = DeviceInfo.FirmwareVersion,
        CalibrationDate    = DeviceInfo.CalibrationDate,
        ConnectionAddress  = $"{_cfg.Host}:{_cfg.Port}",
        ConnectionMedium   = _cfg.ResolvedConnectionMedium(),
        SensorWidthPx      = DeviceInfo.SensorWidthPx,
        SensorHeightPx     = DeviceInfo.SensorHeightPx,
        SensorPixelPitchUm = DeviceInfo.SensorPixelPitchUm,
        ImageSizeSetting   = DeviceInfo.ImageSizeSetting,
    };

    private static DateTime? ParseCalibrationDate(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        return DateTime.TryParse(raw, out var dt) ? dt : null;
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;
        await DisconnectAsync();
        await _client.DisposeAsync();
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DeviceSession));
    }
}

/// <summary>Metadata about the connected device, populated during ConnectAsync.</summary>
public sealed class DeviceInfo
{
    public string?   Type            { get; init; }
    public string?   Serial          { get; init; }
    public string?   Name            { get; init; }
    public string?   FirmwareVersion { get; init; }
    public DateTime? CalibrationDate { get; init; }

    // ── Sensor / imaging metadata ──────────────────────────────────────────
    // Populated at ConnectAsync from DeviceSensorSpecs lookup (static per model)
    // and from GET IMAGE.SIZE (device-stored setting).

    /// <summary>Native sensor width in pixels. Null if model not in lookup table.</summary>
    public int?    SensorWidthPx      { get; init; }
    /// <summary>Native sensor height in pixels. Null if model not in lookup table.</summary>
    public int?    SensorHeightPx     { get; init; }
    /// <summary>Pixel pitch in µm, e.g. 3.45. Null if model not in lookup table.</summary>
    public double? SensorPixelPitchUm { get; init; }
    /// <summary>
    /// Device's current IMAGE.SIZE setting: "Full", "1/4", "1/16", "1/64".
    /// Controls IMAGE.SEND output resolution only; does NOT affect push XML JPEG crop.
    /// </summary>
    public string? ImageSizeSetting   { get; init; }
}

/// <summary>Thrown when a device connection cannot be established or is lost.</summary>
public sealed class DeviceConnectionException : Exception
{
    public DeviceConnectionException(string message, Exception? inner = null)
        : base(message, inner) { }
}
