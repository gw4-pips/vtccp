namespace InlineIo;

/// <summary>
/// Hardware abstraction for a multi-channel relay board.
/// Implementations:
///   <see cref="MockRelayBoard"/>  — in-memory mock for development/testing (no hardware required).
///   RealRelayBoard               — ENGINEERING TODO: plug in USB/serial relay board driver here.
/// </summary>
public interface IRelayBoard : IAsyncDisposable
{
    /// <summary>True once <see cref="ConnectAsync"/> has completed successfully.</summary>
    bool IsConnected { get; }

    /// <summary>Total number of relay channels available on this board (1-based addressing).</summary>
    int RelayCount { get; }

    /// <summary>
    /// Open the connection to the relay board.
    /// For physical boards this typically means opening a COM/USB port.
    /// Throws <see cref="InvalidOperationException"/> if already connected.
    /// </summary>
    Task ConnectAsync(CancellationToken ct = default);

    /// <summary>Gracefully closes the board connection. Safe to call when already disconnected.</summary>
    Task DisconnectAsync(CancellationToken ct = default);

    /// <summary>
    /// Energise (<paramref name="on"/> = true) or de-energise a single relay channel.
    /// Throws <see cref="ArgumentOutOfRangeException"/> if <paramref name="channel"/> is outside 1..<see cref="RelayCount"/>.
    /// Throws <see cref="InvalidOperationException"/> if not connected.
    /// </summary>
    Task SetRelayAsync(int channel, bool on, CancellationToken ct = default);

    /// <summary>
    /// Returns the current energised state of a relay channel.
    /// Throws <see cref="ArgumentOutOfRangeException"/> if channel is out of range.
    /// </summary>
    Task<bool> GetRelayStateAsync(int channel, CancellationToken ct = default);

    /// <summary>De-energise all relay channels simultaneously. Safe to call at any time when connected.</summary>
    Task AllOffAsync(CancellationToken ct = default);
}
