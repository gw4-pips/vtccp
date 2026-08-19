// Copyright © 2026 VCCS. All rights reserved.

namespace DeviceInterface.Rfid;

/// <summary>
/// User-facing status plus an optional physical SDK call that is still running.
/// The reader must retain exclusive hardware ownership until
/// <see cref="PendingSdkCall"/> completes.
/// </summary>
internal readonly record struct LockStatusCheckOutcome(
    string Status,
    Task? PendingSdkCall = null);

/// <summary>
/// Resolves the ASR-P35U <c>CheckTagStatus</c> operation from either the
/// synchronous <c>Types.TagStatus</c> return value or an SDK QC callback.
/// </summary>
/// <remarks>
/// SDK 1.3.0 returns <c>Types.TagStatus</c> directly:
/// 0=UnLock, 1=Lock, 2=PermaLock, 3=Unknown, 4=Error.  Some observed SDK paths
/// also emit cbSuccess 40/41/42 or cbError, so the callback task remains armed
/// as a fallback and for complete diagnostics.
/// </remarks>
internal static class LockStatusCheckExecutor
{
    /// <summary>Map a synchronous Types.TagStatus value to the report contract.</summary>
    internal static string MapImmediateResult(uint raw) => raw switch
    {
        0 => "Unlocked",
        1 => "Locked",
        2 => "PermaLocked",
        3 => "Unknown",
        4 => "Unknown",
        _ => "Unknown",
    };

    /// <summary>
    /// Run CheckTagStatus as a physically single-flight vendor operation.
    /// Error/raw 4 and cbError 4 are retried within the overall budget.
    /// </summary>
    internal static async Task<LockStatusCheckOutcome> ExecuteAsync(
        Func<uint> checkTagStatus,
        LockCheckCorrelator correlator,
        TimeSpan timeout,
        TimeSpan busyRetryDelay,
        Action<string>? diagnostic = null,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(checkTagStatus);
        ArgumentNullException.ThrowIfNull(correlator);

        var sw = System.Diagnostics.Stopwatch.StartNew();
        TimeSpan Remaining() => timeout - sw.Elapsed;

        while (Remaining() > TimeSpan.Zero && !ct.IsCancellationRequested)
        {
            var callbackTask = correlator.Arm();
            diagnostic?.Invoke("armed");

            // The vendor call has no cancellation primitive and may block. Run it
            // off-thread so the UI-facing deadline remains enforceable. If the
            // deadline wins, PendingSdkCall tells the reader to retain exclusive
            // hardware ownership until this physical invocation actually finishes.
            Task<uint> returnTask = Task.Run(() =>
            {
                try
                {
                    uint value = checkTagStatus();
                    diagnostic?.Invoke(
                        $"immediate-return raw={value} mapped={MapImmediateResult(value)}");
                    return value;
                }
                catch (Exception ex)
                {
                    diagnostic?.Invoke(
                        $"immediate-exception type={ex.GetType().Name} message={OneLine(ex.Message)}");
                    throw;
                }
            });

            TimeSpan remaining = Remaining();
            if (remaining <= TimeSpan.Zero)
            {
                correlator.Disarm();
                diagnostic?.Invoke("timeout before-wait");
                return new("Unknown", returnTask);
            }

            Task deadlineTask = Task.Delay(remaining, ct);
            await Task.WhenAny(returnTask, deadlineTask).ConfigureAwait(false);

            if (!returnTask.IsCompleted)
            {
                string deadlineStatus = TryGetSuccessfulCallback(callbackTask, out string callbackStatus)
                    ? callbackStatus
                    : "Unknown";
                correlator.Disarm();
                diagnostic?.Invoke(ct.IsCancellationRequested
                    ? $"cancelled status={deadlineStatus} physical-call=pending"
                    : $"timeout status={deadlineStatus} physical-call=pending");
                return new(deadlineStatus, returnTask);
            }

            uint raw;
            try
            {
                raw = await returnTask.ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                if (TryGetSuccessfulCallback(callbackTask, out string callbackFallback))
                {
                    diagnostic?.Invoke(
                        $"result source=callback-after-exception status={callbackFallback}");
                    return new(callbackFallback);
                }

                correlator.Disarm();
                diagnostic?.Invoke(
                    $"result unavailable: immediate-exception {ex.GetType().Name}");
                return new("Unknown");
            }

            // The synchronous Types.TagStatus result is authoritative for all
            // defined non-error values, even if an SDK callback also arrived.
            if (raw <= 3)
            {
                correlator.Disarm();
                string mapped = MapImmediateResult(raw);
                diagnostic?.Invoke($"result source=immediate status={mapped}");
                return new(mapped);
            }

            // A successful callback may be used only when the direct call itself
            // was unavailable (Error/unknown raw value or exception).
            if (TryGetSuccessfulCallback(callbackTask, out string fallback))
            {
                diagnostic?.Invoke($"result source=callback-fallback status={fallback}");
                return new(fallback);
            }

            correlator.Disarm();
            if (raw != 4)
            {
                diagnostic?.Invoke($"result unavailable: undefined-immediate raw={raw}");
                return new("Unknown");
            }

            diagnostic?.Invoke("busy-or-error source=immediate raw=4");
            if (!await DelayForRetryAsync(Remaining(), busyRetryDelay, diagnostic, ct)
                      .ConfigureAwait(false))
                return new("Unknown");
        }

        diagnostic?.Invoke(ct.IsCancellationRequested
            ? "cancelled"
            : "timeout budget-exhausted");
        return new("Unknown");
    }

    private static bool TryGetSuccessfulCallback(
        Task<string> callbackTask,
        out string status)
    {
        if (callbackTask.Status == TaskStatus.RanToCompletion &&
            callbackTask.Result is not (LockCheckCorrelator.Busy or "Unknown"))
        {
            status = callbackTask.Result;
            return true;
        }

        status = string.Empty;
        return false;
    }

    private static async Task<bool> DelayForRetryAsync(
        TimeSpan remaining,
        TimeSpan retryDelay,
        Action<string>? diagnostic,
        CancellationToken ct)
    {
        if (remaining <= retryDelay)
        {
            diagnostic?.Invoke("result unavailable: retry-budget-exhausted");
            return false;
        }

        diagnostic?.Invoke($"retry delayMs={(int)retryDelay.TotalMilliseconds}");
        try
        {
            await Task.Delay(retryDelay, ct).ConfigureAwait(false);
            return true;
        }
        catch (OperationCanceledException)
        {
            diagnostic?.Invoke("cancelled during-retry-delay");
            return false;
        }
    }

    private static string OneLine(string value) =>
        value.Replace('\r', ' ').Replace('\n', ' ');
}