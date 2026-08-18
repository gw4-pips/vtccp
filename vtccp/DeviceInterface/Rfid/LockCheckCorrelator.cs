// Copyright © 2026 VCCS. All rights reserved.

namespace DeviceInterface.Rfid;

/// <summary>
/// Correlates ASR-P35U <c>CheckTagStatus</c> QC callbacks (cbSuccess codes
/// 40=PermaLock, 41=Lock, 42=Unlock; cbError) with the request that issued them.
///
/// Why this exists (FW 1.8.0 / SDK 1.3.0 defect interaction):
/// a timed-out <c>ReadMemory</c> (TID read) emits a *delayed* stray
/// <c>cbSuccess 41</c> after the hardware finishes the RF operation — see
/// references/asr-p35u/docs/ASREADER_TID_DEFECT.md.  If a lock check is armed
/// naively right after the TID read, that stale 41 would be mis-read as
/// "Locked".  This class prevents that with three mechanisms:
///
///   1. <see cref="NoteReadMemoryIssued"/> — records that a stale QC ack is
///      expected from a preceding ReadMemory.
///   2. <see cref="DrainAsync"/> — before arming, waits until the expected
///      stale ack has been observed (or a maximum drain window elapses) AND a
///      quiet period has passed with no unclaimed QC callbacks.
///   3. Busy signalling — cbError code 4 ("device conflict while busy") while
///      armed resolves to <see cref="Busy"/> so the caller can retry instead of
///      reporting a bogus status.
///
/// QC success codes that arrive while NOT armed are recorded as "unclaimed"
/// (they reset the quiet timer) and are never surfaced as a lock result.
///
/// Thread safety: all state is guarded by <see cref="_gate"/>.  Callbacks fire
/// on DLL threads; DrainAsync polls, so no callback ever blocks.
/// </summary>
public sealed class LockCheckCorrelator
{
    /// <summary>Sentinel result: device busy (cbError 4) — caller should retry.</summary>
    public const string Busy = "__busy__";

    private readonly object _gate = new();
    private readonly Func<long> _clock;

    private TaskCompletionSource<string>? _armed;
    private bool _staleAckExpected;
    private long _lastUnclaimedQcMs;

    /// <param name="clock">
    /// Monotonic millisecond clock; defaults to <see cref="Environment.TickCount64"/>.
    /// Injectable for unit tests.
    /// </param>
    public LockCheckCorrelator(Func<long>? clock = null)
        => _clock = clock ?? (static () => Environment.TickCount64);

    /// <summary>
    /// Record that a ReadMemory command was issued and its delayed stray QC ack
    /// (cbSuccess 41 on FW 1.8.0) may still be in flight.
    /// </summary>
    public void NoteReadMemoryIssued()
    {
        lock (_gate) _staleAckExpected = true;
    }

    /// <summary>
    /// Feed a cbSuccess callback.  Returns true when the code was consumed as
    /// the armed lock-check result; false when it was ignored or recorded as an
    /// unclaimed (stale) QC ack.
    /// </summary>
    public bool OnSuccess(uint code)
    {
        if (code is not (40 or 41 or 42))
            return false;   // unrelated success callback — not a QC status

        lock (_gate)
        {
            if (_armed is { } tcs)
            {
                _armed = null;
                tcs.TrySetResult(CodeToStatus(code));
                return true;
            }
            // Unclaimed QC ack — most likely the stale ReadMemory 41.
            _staleAckExpected  = false;
            _lastUnclaimedQcMs = _clock();
            return false;
        }
    }

    /// <summary>
    /// Feed a cbError callback.  Returns true when consumed by an armed lock
    /// check (code 4 → <see cref="Busy"/>, anything else → "Unknown");
    /// false when no lock check is armed (caller handles the error normally).
    /// </summary>
    public bool OnError(uint code)
    {
        lock (_gate)
        {
            if (_armed is not { } tcs)
                return false;
            _armed = null;
            tcs.TrySetResult(code == 4 ? Busy : "Unknown");
            return true;
        }
    }

    /// <summary>
    /// Wait until it is safe to issue CheckTagStatus: the expected stale
    /// ReadMemory ack (if any) has arrived or <paramref name="maxWait"/>
    /// elapsed, AND at least <paramref name="quietPeriod"/> has passed since
    /// drain start / the last unclaimed QC callback.
    /// </summary>
    public async Task DrainAsync(
        TimeSpan quietPeriod,
        TimeSpan maxWait,
        CancellationToken ct = default,
        TimeSpan? pollInterval = null)
    {
        long start = _clock();
        long quietMs = (long)quietPeriod.TotalMilliseconds;
        long maxMs   = (long)maxWait.TotalMilliseconds;
        var poll     = pollInterval ?? TimeSpan.FromMilliseconds(50);

        while (true)
        {
            bool stale; long lastQc;
            lock (_gate) { stale = _staleAckExpected; lastQc = _lastUnclaimedQcMs; }

            long now = _clock();
            if (now - start >= maxMs)
                return;   // drained as long as we are willing to
            if (!stale && now - Math.Max(start, lastQc) >= quietMs)
                return;   // no ack pending and the line has been quiet

            await Task.Delay(poll, ct).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Arm a one-shot lock check.  The returned task completes with
    /// "PermaLocked"/"Locked"/"Unlocked" (cbSuccess), "Unknown" (cbError ≠ 4),
    /// or <see cref="Busy"/> (cbError 4).  Any previously armed task is cancelled.
    /// </summary>
    public Task<string> Arm()
    {
        lock (_gate)
        {
            _armed?.TrySetCanceled();
            _armed = new TaskCompletionSource<string>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            return _armed.Task;
        }
    }

    /// <summary>Disarm without a result (timeout / command rejected).</summary>
    public void Disarm()
    {
        lock (_gate)
        {
            _armed?.TrySetCanceled();
            _armed = null;
        }
    }

    /// <summary>Clear all state (disconnect / inventory abort).</summary>
    public void Reset()
    {
        lock (_gate)
        {
            _armed?.TrySetCanceled();
            _armed             = null;
            _staleAckExpected  = false;
            _lastUnclaimedQcMs = 0;
        }
    }

    /// <summary>Map an ASR-P35U QC success code to the report status string.</summary>
    public static string CodeToStatus(uint code) => code switch
    {
        40 => "PermaLocked",
        41 => "Locked",
        42 => "Unlocked",
        _  => "Unknown",
    };
}
