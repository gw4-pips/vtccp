namespace InlineIo;

/// <summary>
/// In-memory mock relay board for development and test-harness use.
/// No hardware required.  All relay state changes are logged to <see cref="Console"/>.
/// Thread-safe via a dedicated lock.
/// </summary>
public sealed class MockRelayBoard : IRelayBoard
{
    private readonly int _relayCount;
    private readonly bool[] _states;
    private readonly object _lock = new();
    private bool _connected;
    private bool _disposed;

    /// <param name="relayCount">Number of relay channels to simulate.  Default 8.</param>
    public MockRelayBoard(int relayCount = 8)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(relayCount, 1);
        _relayCount = relayCount;
        _states = new bool[relayCount]; // all off
    }

    // ── IRelayBoard ──────────────────────────────────────────────────────────

    /// <inheritdoc/>
    public bool IsConnected
    {
        get { lock (_lock) return _connected; }
    }

    /// <inheritdoc/>
    public int RelayCount => _relayCount;

    /// <inheritdoc/>
    public Task ConnectAsync(CancellationToken ct = default)
    {
        lock (_lock)
        {
            ThrowIfDisposed();
            if (_connected) throw new InvalidOperationException("MockRelayBoard: already connected.");
            _connected = true;
        }
        Console.WriteLine($"[MockRelayBoard] Connected ({_relayCount} channels).");
        return Task.CompletedTask;
    }

    /// <inheritdoc/>
    public Task DisconnectAsync(CancellationToken ct = default)
    {
        lock (_lock)
        {
            if (!_connected) return Task.CompletedTask;
            _connected = false;
            Array.Clear(_states);
        }
        Console.WriteLine("[MockRelayBoard] Disconnected.");
        return Task.CompletedTask;
    }

    /// <inheritdoc/>
    public Task SetRelayAsync(int channel, bool on, CancellationToken ct = default)
    {
        ValidateChannel(channel);
        lock (_lock)
        {
            ThrowIfNotConnected();
            _states[channel - 1] = on;
        }
        Console.WriteLine($"[MockRelayBoard] CH{channel:D2} → {(on ? "ON " : "OFF")}");
        return Task.CompletedTask;
    }

    /// <inheritdoc/>
    public Task<bool> GetRelayStateAsync(int channel, CancellationToken ct = default)
    {
        ValidateChannel(channel);
        bool state;
        lock (_lock)
        {
            ThrowIfNotConnected();
            state = _states[channel - 1];
        }
        return Task.FromResult(state);
    }

    /// <inheritdoc/>
    public Task AllOffAsync(CancellationToken ct = default)
    {
        lock (_lock)
        {
            ThrowIfNotConnected();
            Array.Clear(_states);
        }
        Console.WriteLine("[MockRelayBoard] All channels OFF.");
        return Task.CompletedTask;
    }

    // ── IAsyncDisposable ─────────────────────────────────────────────────────

    public async ValueTask DisposeAsync()
    {
        bool wasConnected;
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;
            wasConnected = _connected;
        }
        if (wasConnected) await DisconnectAsync();
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private void ValidateChannel(int channel)
    {
        if (channel < 1 || channel > _relayCount)
            throw new ArgumentOutOfRangeException(nameof(channel),
                $"Channel {channel} is outside valid range 1–{_relayCount}.");
    }

    private void ThrowIfNotConnected()
    {
        if (!_connected) throw new InvalidOperationException("MockRelayBoard: not connected.");
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(MockRelayBoard));
    }
}
