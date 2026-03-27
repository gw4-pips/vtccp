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
    private readonly DeviceConfig      _cfg;
    private readonly VerificationXmlMap _map;
    private readonly DmccClient        _client;
    private DmstListener?              _listener;
    private bool                       _disposed;

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
        _client = new DmccClient(_cfg);
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

        // Configure device for full result output.
        await _client.SendAsync(DmccCommand.SetResultFormatFull, ct);

        // Query device info.
        DeviceInfo = new DeviceInfo
        {
            Type            = (await _client.SendAsync(DmccCommand.GetDeviceType,    ct)).Body,
            FirmwareVersion = (await _client.SendAsync(DmccCommand.GetFirmwareVer,   ct)).Body,
            Name            = (await _client.SendAsync(DmccCommand.GetDeviceName,    ct)).Body,
            Serial          = (await _client.SendAsync(DmccCommand.GetDeviceId,      ct)).Body,
            CalibrationDate = ParseCalibrationDate(
                              (await _client.SendAsync(DmccCommand.GetCalibrationDate, ct)).Body),
        };
    }

    /// <summary>Closes the DMCC connection and stops any active push listener.</summary>
    public async Task DisconnectAsync()
    {
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

        DmccResponse trigResp = await _client.SendAsync(DmccCommand.Trigger, ct);
        if (trigResp.StatusCode == DmccStatus.NoRead) return null;

        DmccResponse resultResp = await _client.SendAsync(DmccCommand.GetSymbolResult, ct);
        if (!resultResp.IsSuccess || !resultResp.IsXml) return null;

        return DmstResultParser.Parse(resultResp.Body, _map, sessionContext ?? ContextFromDeviceInfo());
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
