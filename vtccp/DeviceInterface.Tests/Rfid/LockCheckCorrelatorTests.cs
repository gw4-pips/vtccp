// Copyright © 2026 VCCS. All rights reserved.

using DeviceInterface.Rfid;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

/// <summary>
/// Regression tests for <see cref="LockCheckCorrelator"/> — the state machine
/// that prevents the FW 1.8.0 delayed stray cbSuccess 41 (emitted by a
/// timed-out TID ReadMemory) from being mis-read as a CheckTagStatus "Locked"
/// result.  See references/asr-p35u/docs/ASREADER_TID_DEFECT.md.
/// </summary>
public sealed class LockCheckCorrelatorTests
{
    private static readonly TimeSpan Quiet = TimeSpan.FromMilliseconds(80);
    private static readonly TimeSpan Max   = TimeSpan.FromMilliseconds(1500);
    private static readonly TimeSpan Poll  = TimeSpan.FromMilliseconds(5);

    // ── The reviewer scenario: delayed stale 41 must not become "Locked" ──────

    [Fact]
    public async Task StaleReadMemory41_DuringDrain_IsNotConsumed_RealResponseWins()
    {
        var c = new LockCheckCorrelator();
        c.NoteReadMemoryIssued();                    // TID ReadMemory issued

        var drain = c.DrainAsync(Quiet, Max, pollInterval: Poll);

        await Task.Delay(30);
        Assert.False(drain.IsCompleted);             // still waiting for the stale ack

        // Delayed stray cbSuccess 41 arrives — must be recorded, not consumed.
        Assert.False(c.OnSuccess(41));

        await drain;                                 // completes after quiet period

        // Now arm the real lock check; only the genuine response resolves it.
        var armed = c.Arm();
        Assert.False(armed.IsCompleted);
        Assert.True(c.OnSuccess(42));                // actual CheckTagStatus result
        Assert.Equal("Unlocked", await armed);
    }

    [Fact]
    public async Task SuccessfulTidPath_NoStaleExpectation_DrainIsFast()
    {
        // Normal TID read (result via cbTag) sets NO stale-ack expectation:
        // drain must complete after just the quiet period, nowhere near maxWait,
        // so CheckTagStatus is issued while the tag is still in RF range.
        var c = new LockCheckCorrelator();

        var sw = System.Diagnostics.Stopwatch.StartNew();
        await c.DrainAsync(Quiet, TimeSpan.FromMilliseconds(2500), pollInterval: Poll);
        sw.Stop();

        Assert.True(sw.ElapsedMilliseconds < 500,
            $"drain took {sw.ElapsedMilliseconds} ms — expected quiet-period only");

        var armed = c.Arm();
        Assert.True(c.OnSuccess(40));
        Assert.Equal("PermaLocked", await armed);
    }

    [Fact]
    public async Task Drain_WaitsForStaleAck_UpToMaxWait()
    {
        var c = new LockCheckCorrelator();
        c.NoteReadMemoryIssued();

        // Stale ack never arrives — drain must still return by maxWait.
        var sw = System.Diagnostics.Stopwatch.StartNew();
        await c.DrainAsync(Quiet, TimeSpan.FromMilliseconds(200), pollInterval: Poll);
        Assert.True(sw.ElapsedMilliseconds >= 190);
    }

    [Fact]
    public async Task Drain_QuietTimer_ResetsOnUnclaimedQcAck()
    {
        var c = new LockCheckCorrelator();
        var drain = c.DrainAsync(Quiet, Max, pollInterval: Poll);

        await Task.Delay(40);
        c.OnSuccess(41);                             // unclaimed ack mid-drain
        await Task.Delay(40);
        Assert.False(drain.IsCompleted);             // quiet timer was reset

        await drain;                                 // eventually quiets down
    }

    // ── Armed behaviour ────────────────────────────────────────────────────────

    [Theory]
    [InlineData(40u, "PermaLocked")]
    [InlineData(41u, "Locked")]
    [InlineData(42u, "Unlocked")]
    public async Task ArmedSuccessCode_MapsToStatus(uint code, string expected)
    {
        var c = new LockCheckCorrelator();
        var armed = c.Arm();
        Assert.True(c.OnSuccess(code));
        Assert.Equal(expected, await armed);
    }

    [Fact]
    public void NonQcSuccessCodes_AreIgnored_EvenWhileArmed()
    {
        var c = new LockCheckCorrelator();
        var armed = c.Arm();
        Assert.False(c.OnSuccess(0));
        Assert.False(c.OnSuccess(39));
        Assert.False(c.OnSuccess(43));
        Assert.False(armed.IsCompleted);
    }

    [Fact]
    public async Task ArmedError4_ResolvesBusySentinel()
    {
        var c = new LockCheckCorrelator();
        var armed = c.Arm();
        Assert.True(c.OnError(4));
        Assert.Equal(LockCheckCorrelator.Busy, await armed);
    }

    [Fact]
    public async Task ArmedOtherError_ResolvesUnknown()
    {
        var c = new LockCheckCorrelator();
        var armed = c.Arm();
        Assert.True(c.OnError(1));
        Assert.Equal("Unknown", await armed);
    }

    [Fact]
    public void UnarmedError_IsNotConsumed()
    {
        var c = new LockCheckCorrelator();
        Assert.False(c.OnError(4));                  // caller handles it normally
    }

    // ── Disarm / reset ─────────────────────────────────────────────────────────

    [Fact]
    public async Task Disarm_CancelsArmedTask_AndLaterAckIsUnclaimed()
    {
        var c = new LockCheckCorrelator();
        var armed = c.Arm();
        c.Disarm();
        await Assert.ThrowsAsync<TaskCanceledException>(() => armed);

        Assert.False(c.OnSuccess(41));               // recorded as unclaimed, not consumed
    }

    [Fact]
    public async Task Rearm_CancelsPreviousArmedTask()
    {
        var c = new LockCheckCorrelator();
        var first  = c.Arm();
        var second = c.Arm();
        await Assert.ThrowsAsync<TaskCanceledException>(() => first);
        Assert.True(c.OnSuccess(40));
        Assert.Equal("PermaLocked", await second);
    }
}
