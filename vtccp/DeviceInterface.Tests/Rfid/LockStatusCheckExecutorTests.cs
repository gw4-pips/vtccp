// Copyright © 2026 VCCS. All rights reserved.

using DeviceInterface.Rfid;
using Xunit;

namespace DeviceInterface.Tests.Rfid;

public sealed class LockStatusCheckExecutorTests
{
    [Theory]
    [InlineData(0u, "Unlocked")]
    [InlineData(1u, "Locked")]
    [InlineData(2u, "PermaLocked")]
    [InlineData(3u, "Unknown")]
    [InlineData(4u, "Unknown")]
    [InlineData(99u, "Unknown")]
    public void ImmediateTagStatus_MapsToReportContract(uint raw, string expected)
    {
        Assert.Equal(expected, LockStatusCheckExecutor.MapImmediateResult(raw));
    }

    [Fact]
    public async Task ImmediatePermaLock_IsReportedAsPermaLocked()
    {
        var correlator = new LockCheckCorrelator();

        LockStatusCheckOutcome outcome = await LockStatusCheckExecutor.ExecuteAsync(
            () => 2,
            correlator,
            TimeSpan.FromSeconds(1),
            TimeSpan.FromMilliseconds(5));

        Assert.Equal("PermaLocked", outcome.Status);
        Assert.Null(outcome.PendingSdkCall);
    }

    [Fact]
    public async Task Callback40_IsFallback_WhenImmediateCallIsUnavailable()
    {
        var correlator = new LockCheckCorrelator();

        LockStatusCheckOutcome outcome = await LockStatusCheckExecutor.ExecuteAsync(
            () =>
            {
                correlator.OnSuccess(40);
                throw new IOException("simulated SDK return failure");
            },
            correlator,
            TimeSpan.FromSeconds(1),
            TimeSpan.FromMilliseconds(5));

        Assert.Equal("PermaLocked", outcome.Status);
        Assert.Null(outcome.PendingSdkCall);
    }

    [Fact]
    public async Task ImmediateResult_WinsOverConflictingCallback()
    {
        var correlator = new LockCheckCorrelator();

        LockStatusCheckOutcome outcome = await LockStatusCheckExecutor.ExecuteAsync(
            () =>
            {
                correlator.OnSuccess(42);
                return 2;
            },
            correlator,
            TimeSpan.FromSeconds(1),
            TimeSpan.FromMilliseconds(5));

        Assert.Equal("PermaLocked", outcome.Status);
    }

    [Fact]
    public async Task ImmediateError4_Retries_ThenUsesVerifiedResult()
    {
        var correlator = new LockCheckCorrelator();
        int attempts = 0;

        LockStatusCheckOutcome outcome = await LockStatusCheckExecutor.ExecuteAsync(
            () => Interlocked.Increment(ref attempts) == 1 ? 4u : 2u,
            correlator,
            TimeSpan.FromSeconds(1),
            TimeSpan.FromMilliseconds(5));

        Assert.Equal("PermaLocked", outcome.Status);
        Assert.Equal(2, attempts);
    }

    [Fact]
    public async Task ImmediateError4_UntilBudgetExpires_ReportsUnknown()
    {
        var correlator = new LockCheckCorrelator();

        LockStatusCheckOutcome outcome = await LockStatusCheckExecutor.ExecuteAsync(
            () => 4,
            correlator,
            TimeSpan.FromMilliseconds(40),
            TimeSpan.FromMilliseconds(5));

        Assert.Equal("Unknown", outcome.Status);
    }

    [Fact]
    public async Task NonBusyErrorCallback_ReportsUnknown()
    {
        var correlator = new LockCheckCorrelator();

        LockStatusCheckOutcome outcome = await LockStatusCheckExecutor.ExecuteAsync(
            () =>
            {
                correlator.OnError(9);
                return 4;
            },
            correlator,
            TimeSpan.FromMilliseconds(40),
            TimeSpan.FromMilliseconds(5));

        Assert.Equal("Unknown", outcome.Status);
    }

    [Fact]
    public async Task BlockingSdkCall_ReturnsAtDeadline_WithPhysicalCallPending()
    {
        var correlator = new LockCheckCorrelator();

        LockStatusCheckOutcome outcome = await LockStatusCheckExecutor.ExecuteAsync(
            () =>
            {
                Thread.Sleep(30);
                return 2;
            },
            correlator,
            TimeSpan.FromMilliseconds(5),
            TimeSpan.FromMilliseconds(1));

        Assert.Equal("Unknown", outcome.Status);
        Assert.NotNull(outcome.PendingSdkCall);
        await outcome.PendingSdkCall!;
    }

    [Fact]
    public async Task DelayedPriorCallback_DrainedBeforeNextError_CannotBecomeFallback()
    {
        var correlator = new LockCheckCorrelator();

        LockStatusCheckOutcome first = await LockStatusCheckExecutor.ExecuteAsync(
            () => 2,
            correlator,
            TimeSpan.FromSeconds(1),
            TimeSpan.FromMilliseconds(5));
        Assert.Equal("PermaLocked", first.Status);

        Task postCheckDrain = correlator.DrainAsync(
            TimeSpan.FromMilliseconds(40),
            TimeSpan.FromMilliseconds(500),
            pollInterval: TimeSpan.FromMilliseconds(5));

        await Task.Delay(15);
        Assert.False(correlator.OnSuccess(40));
        await postCheckDrain;

        LockStatusCheckOutcome second = await LockStatusCheckExecutor.ExecuteAsync(
            () => 4,
            correlator,
            TimeSpan.FromMilliseconds(35),
            TimeSpan.FromMilliseconds(5));

        Assert.Equal("Unknown", second.Status);
    }
}