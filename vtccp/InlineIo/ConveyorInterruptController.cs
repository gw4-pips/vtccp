namespace InlineIo;

/// <summary>
/// Controls the conveyor stop/restart relay channels for CP Inline.
///
/// Two wiring modes are supported (configure via <see cref="Models.RelayChannelMap.ConveyorRestart"/>):
/// <list type="bullet">
///   <item>
///     Single-channel (ConveyorRestart = -1): energising ConveyorStop opens the drive circuit;
///     de-energising it resumes the belt automatically.  Confirm with engineering.
///   </item>
///   <item>
///     Dual-channel (ConveyorRestart ≠ -1): separate STOP and RESTART relay outputs,
///     typically for a PLC with dedicated inputs.  A brief pulse is sent on restart.
///   </item>
/// </list>
///
/// ⚠ ENGINEERING TODO: confirm channel assignments, pulse duration, and whether the
/// site PLC needs a momentary or latched input before populating <see cref="Models.RelayChannelMap"/>.
/// </summary>
public sealed class ConveyorInterruptController : IAsyncDisposable
{
    private readonly IRelayBoard _board;
    private readonly int _stopChannel;
    private readonly int _restartChannel; // -1 = not wired

    /// <summary>Duration of the restart pulse sent on the restart channel (dual-channel mode).</summary>
    public TimeSpan RestartPulseDuration { get; set; } = TimeSpan.FromMilliseconds(250);

    private bool _stopped;
    private bool _disposed;

    public ConveyorInterruptController(IRelayBoard board, Models.RelayChannelMap? map = null)
    {
        _board          = board ?? throw new ArgumentNullException(nameof(board));
        var m           = map ?? Models.RelayChannelMap.Default;
        _stopChannel    = m.ConveyorStop;
        _restartChannel = m.ConveyorRestart;
    }

    // ── Public API ───────────────────────────────────────────────────────────

    /// <summary>True after <see cref="StopAsync"/> until <see cref="ResumeAsync"/> completes.</summary>
    public bool IsConveyorStopped => _stopped;

    /// <summary>
    /// Halt the conveyor by energising the stop relay.
    /// Idempotent — safe to call when already stopped.
    /// </summary>
    public async Task StopAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();
        if (_stopChannel < 1)
        {
            Console.WriteLine("[ConveyorInterruptController] Stop channel not wired — skipping.");
            return;
        }
        await _board.SetRelayAsync(_stopChannel, true, ct);
        _stopped = true;
        Console.WriteLine("[ConveyorInterruptController] Conveyor STOPPED.");
    }

    /// <summary>
    /// Resume the conveyor.
    /// In single-channel mode: de-energises the stop relay.
    /// In dual-channel mode: de-energises stop, then sends a timed pulse on the restart channel.
    /// Idempotent — safe to call when already running.
    /// </summary>
    public async Task ResumeAsync(CancellationToken ct = default)
    {
        ThrowIfDisposed();

        if (_stopChannel > 0)
            await _board.SetRelayAsync(_stopChannel, false, ct);

        if (_restartChannel > 0)
        {
            // Momentary restart pulse
            await _board.SetRelayAsync(_restartChannel, true, ct);
            await Task.Delay(RestartPulseDuration, ct);
            await _board.SetRelayAsync(_restartChannel, false, ct);
        }

        _stopped = false;
        Console.WriteLine("[ConveyorInterruptController] Conveyor RESUMED.");
    }

    // ── IAsyncDisposable ─────────────────────────────────────────────────────

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;
        // Safety: ensure conveyor is not left stopped if the app exits cleanly
        if (_stopped && _board.IsConnected)
        {
            try { await ResumeAsync(); }
            catch { /* best-effort on dispose */ }
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(ConveyorInterruptController));
    }
}
