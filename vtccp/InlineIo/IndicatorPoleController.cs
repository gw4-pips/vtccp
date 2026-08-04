using InlineIo.Models;

namespace InlineIo;

/// <summary>
/// Controls a multi-colour indicator pole wired via a relay board.
/// Supports STEADY and FLASH modes; flash is driven by an internal background task.
///
/// Grade → colour mapping (per CP Inline specification):
/// <list type="bullet">
///   <item>No decode             → RED   FLASH  + conveyor stop</item>
///   <item>Grade &lt; 1.8        → RED   STEADY + conveyor stop</item>
///   <item>1.8 ≤ grade ≤ 2.3    → AMBER STEADY</item>
///   <item>2.4 ≤ grade ≤ 2.8    → AMBER FLASH</item>
///   <item>2.9 ≤ grade ≤ 3.4    → GREEN STEADY</item>
///   <item>3.5 ≤ grade ≤ 4.0    → GREEN FLASH</item>
/// </list>
///
/// This class is NOT thread-safe across concurrent callers; the expected usage is
/// a single WPF dispatcher / decode-result handler calling into it serially.
/// </summary>
public sealed class IndicatorPoleController : IAsyncDisposable
{
    private readonly IRelayBoard _board;
    private readonly RelayChannelMap _map;

    private CancellationTokenSource? _flashCts;
    private Task _flashTask = Task.CompletedTask;

    private IndicatorColour _activeColour = IndicatorColour.Off;
    private IndicatorMode   _activeMode   = IndicatorMode.Steady;

    /// <summary>
    /// Flash on/off half-period in milliseconds.  Default 500 ms → 1 Hz flash (0.5 s on, 0.5 s off).
    /// Adjust before first <see cref="SetAsync"/> call if a different frequency is required.
    /// </summary>
    public int FlashPeriodMs { get; set; } = 500;

    public IndicatorPoleController(IRelayBoard board, RelayChannelMap? map = null)
    {
        _board = board ?? throw new ArgumentNullException(nameof(board));
        _map   = map ?? RelayChannelMap.Default;
    }

    // ── Public API ───────────────────────────────────────────────────────────

    /// <summary>Current active colour (Off when cleared).</summary>
    public IndicatorColour ActiveColour => _activeColour;

    /// <summary>Current active mode.</summary>
    public IndicatorMode ActiveMode => _activeMode;

    /// <summary>
    /// Light the specified colour in the specified mode.
    /// Any previously active colour is extinguished first.
    /// Starts a background flash loop when <paramref name="mode"/> is <see cref="IndicatorMode.Flash"/>.
    /// </summary>
    public async Task SetAsync(IndicatorColour colour, IndicatorMode mode, CancellationToken ct = default)
    {
        // Stop any running flash loop
        await StopFlashLoopAsync();

        // Extinguish all lamps
        await TurnAllLampsOffAsync(ct);

        if (colour == IndicatorColour.Off)
        {
            _activeColour = IndicatorColour.Off;
            _activeMode   = IndicatorMode.Steady;
            return;
        }

        _activeColour = colour;
        _activeMode   = mode;

        int ch = ChannelFor(colour);
        if (ch == -1) return; // channel not wired; log but carry on

        if (mode == IndicatorMode.Steady)
        {
            await _board.SetRelayAsync(ch, true, ct);
        }
        else
        {
            // Start background flash loop
            _flashCts  = new CancellationTokenSource();
            _flashTask = RunFlashLoopAsync(ch, _flashCts.Token);
        }
    }

    /// <summary>
    /// Convenience method: derive the correct colour + mode from a numeric grade.
    /// Returns the chosen colour so the caller can decide whether to also stop the conveyor.
    /// </summary>
    /// <param name="numericGrade">ISO/IEC 15415 overall numeric grade (0.0 – 4.0).</param>
    /// <returns>The colour that was set.</returns>
    public async Task<IndicatorColour> SetForGradeAsync(decimal numericGrade, CancellationToken ct = default)
    {
        var (colour, mode) = ClassifyGrade(numericGrade);
        await SetAsync(colour, mode, ct);
        return colour;
    }

    /// <summary>
    /// Set indicator for a no-decode result: RED FLASH.
    /// The caller is responsible for triggering a conveyor stop when appropriate.
    /// </summary>
    public async Task SetForNoDecodeAsync(CancellationToken ct = default)
        => await SetAsync(IndicatorColour.Red, IndicatorMode.Flash, ct);

    /// <summary>Turn all indicator lamps off and stop any flash loop.</summary>
    public async Task ClearAsync(CancellationToken ct = default)
        => await SetAsync(IndicatorColour.Off, IndicatorMode.Steady, ct);

    // ── Grade classification ─────────────────────────────────────────────────

    /// <summary>
    /// Map a numeric grade to (colour, mode) per the CP Inline specification.
    /// Exposed as static so the WPF ViewModel can preview the mapping in unit tests
    /// without constructing a controller instance.
    /// </summary>
    public static (IndicatorColour Colour, IndicatorMode Mode) ClassifyGrade(decimal grade)
    {
        return grade switch
        {
            < 1.8m                    => (IndicatorColour.Red,   IndicatorMode.Steady),
            >= 1.8m and <= 2.3m       => (IndicatorColour.Amber, IndicatorMode.Steady),
            >= 2.4m and <= 2.8m       => (IndicatorColour.Amber, IndicatorMode.Flash),
            >= 2.9m and <= 3.4m       => (IndicatorColour.Green, IndicatorMode.Steady),
            >= 3.5m and <= 4.0m       => (IndicatorColour.Green, IndicatorMode.Flash),
            // Out-of-spec (>4.0 should not occur; treat as best-pass)
            _                         => (IndicatorColour.Green, IndicatorMode.Flash),
        };
    }

    /// <summary>
    /// True if the grade (or no-decode) result requires the line to be stopped.
    /// Per spec: grade &lt; 1.8 (any decode) or no-decode.
    /// </summary>
    public static bool RequiresConveyorStop(decimal? grade)
        => grade is null || grade < 1.8m;

    // ── IAsyncDisposable ─────────────────────────────────────────────────────

    public async ValueTask DisposeAsync()
    {
        await StopFlashLoopAsync();
        if (_board.IsConnected)
            await TurnAllLampsOffAsync();
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    private int ChannelFor(IndicatorColour colour) => colour switch
    {
        IndicatorColour.Red   => _map.Red,
        IndicatorColour.Amber => _map.Amber,
        IndicatorColour.Green => _map.Green,
        IndicatorColour.Blue  => _map.Blue,
        _                     => -1,
    };

    private async Task TurnAllLampsOffAsync(CancellationToken ct = default)
    {
        // Only touch channels that are actually wired (-1 = skip)
        foreach (int ch in new[] { _map.Red, _map.Amber, _map.Green, _map.Blue })
        {
            if (ch > 0) await _board.SetRelayAsync(ch, false, ct);
        }
    }

    private async Task RunFlashLoopAsync(int channel, CancellationToken ct)
    {
        try
        {
            bool on = true;
            while (!ct.IsCancellationRequested)
            {
                await _board.SetRelayAsync(channel, on, ct);
                on = !on;
                await Task.Delay(FlashPeriodMs, ct);
            }
        }
        catch (OperationCanceledException) { /* normal shutdown */ }
        finally
        {
            // Ensure lamp is left OFF when the flash loop exits
            try { await _board.SetRelayAsync(channel, false, CancellationToken.None); }
            catch { /* best-effort */ }
        }
    }

    private async Task StopFlashLoopAsync()
    {
        if (_flashCts is not null)
        {
            await _flashCts.CancelAsync();
            try { await _flashTask; } catch { /* absorbed */ }
            _flashCts.Dispose();
            _flashCts = null;
        }
    }
}
