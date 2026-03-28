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
    private bool                        _disposed;

    /// <summary>
    /// Original trigger type read from the device in ConnectAsync.
    /// Restored in DisconnectAsync so VTCCP does not permanently alter device settings.
    /// Null if GET TRIGGER.TYPE was not supported by this firmware.
    /// </summary>
    private string? _originalTriggerType;

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

        // Result format is set via SDK's SetResultTypes() in DataManSdkClient.ConnectAsync.

        // Query device identity info.
        // FirmwareVersion is read from the SDK's native property (avoids InvalidCommandException
        // thrown by SendCommand("GET FIRMWARE.VER") on this firmware).
        DeviceInfo = new DeviceInfo
        {
            Type            = (await _client.SendAsync(DmccCommand.GetDeviceType, ct)).Body,
            FirmwareVersion = _client.FirmwareVersion
                           ?? (await _client.SendAsync(DmccCommand.GetFirmwareVer, ct)).Body,
            Name            = (await _client.SendAsync(DmccCommand.GetDeviceName, ct)).Body,
            Serial          = (await _client.SendAsync(DmccCommand.GetDeviceId,   ct)).Body,
        };

        // ── Trigger mode ─────────────────────────────────────────────────────
        // Many DMV devices default to External (hardware-only) trigger mode.
        // Software TRIGGER commands are received but the device arms and waits
        // for an electrical hardware pulse that never arrives, then times out
        // (~6 s) and returns No Read without ever flashing the illumination.
        //
        // Query the current trigger type; if it is External (or any non-Single
        // variant), switch to Single so VTCCP's software TRIGGER fires the scan
        // immediately.  The original type is restored in DisconnectAsync so we
        // do not permanently alter device settings.
        var trigResp = await _client.SendAsync(DmccCommand.GetTriggerType, ct);
        System.Diagnostics.Debug.WriteLine(
            $"[VTCCP-DMCC] Trigger type on connect: code={trigResp.StatusCode}  value='{trigResp.Body}'");

        if (trigResp.IsSuccess && !string.IsNullOrWhiteSpace(trigResp.Body))
        {
            _originalTriggerType = trigResp.Body.Trim();

            // Switch to Single if the device is NOT already in a software-trigger mode.
            bool needsSingle = !_originalTriggerType.Equals("Single",    StringComparison.OrdinalIgnoreCase)
                            && !_originalTriggerType.Equals("Software",  StringComparison.OrdinalIgnoreCase)
                            && !_originalTriggerType.Equals("Continuous",StringComparison.OrdinalIgnoreCase)
                            && !_originalTriggerType.Equals("Self",      StringComparison.OrdinalIgnoreCase);

            if (needsSingle)
            {
                var setResp = await _client.SendAsync(DmccCommand.SetTriggerTypeSingle, ct);
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-DMCC] SET TRIGGER.TYPE Single: code={setResp.StatusCode}");
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-DMCC] GET TRIGGER.TYPE not supported (code={trigResp.StatusCode}) — skipping trigger mode setup.");
        }
    }

    /// <summary>Closes the DMCC connection and stops any active push listener.</summary>
    public async Task DisconnectAsync()
    {
        // Restore the original trigger type so VTCCP does not permanently change
        // the device's configured trigger source.
        if (_originalTriggerType is not null && _client.IsConnected)
        {
            try
            {
                var restoreCmd = $"SET TRIGGER.TYPE {_originalTriggerType}";
                await _client.SendAsync(restoreCmd);
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-DMCC] Restored trigger type to '{_originalTriggerType}'.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[VTCCP-DMCC] Could not restore trigger type: {ex.Message}");
            }
        }

        if (_listener is not null) await _listener.StopAsync();
        await _client.DisconnectAsync();
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

        return DmstResultParser.Parse(xml, _map, sessionContext ?? ContextFromDeviceInfo());
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
            rec => ResultReceived?.Invoke(this, rec));

        await _listener.StartAsync(ct);
    }

    /// <summary>Stops the push listener if active.</summary>
    public async Task StopPushListenerAsync()
    {
        if (_listener is null) return;
        await _listener.StopAsync();
        _listener = null;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>Builds a thin context record from device info for pre-seeding result records.</summary>
    private VerificationRecord ContextFromDeviceInfo() => new()
    {
        Symbology       = "Unknown",  // filled in by parser
        DeviceSerial    = DeviceInfo.Serial,
        DeviceName      = DeviceInfo.Name,
        FirmwareVersion = DeviceInfo.FirmwareVersion,
        CalibrationDate = DeviceInfo.CalibrationDate,
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
}

/// <summary>Thrown when a device connection cannot be established or is lost.</summary>
public sealed class DeviceConnectionException : Exception
{
    public DeviceConnectionException(string message, Exception? inner = null)
        : base(message, inner) { }
}
