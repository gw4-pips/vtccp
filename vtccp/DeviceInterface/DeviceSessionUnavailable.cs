namespace DeviceInterface;

using DeviceInterface.Dmst;
using ExcelEngine.Models;

/// <summary>
/// Compile-time fallback for <see cref="DeviceSession"/> when the Cognex DataMan
/// SDK is not installed. It keeps desktop callers statically verifiable in CI
/// and on non-Windows development machines, while rejecting device operations
/// clearly at runtime.
///
/// This file is excluded whenever the real SDK-backed DeviceSession is compiled.
/// </summary>
public sealed class DeviceSession : IAsyncDisposable
{
    private const string UnavailableMessage =
        "Cognex DataMan SDK is unavailable in this build. Install " +
        "Cognex.DataMan.SDK.PC.dll to use device-session features.";

    /// <summary>Always null because the fallback never connects to a device.</summary>
    public string? OriginalTriggerType => null;

    /// <summary>Always null because the fallback does not watch DMST reports.</summary>
    public string? LastMatchedSourcePath => null;

    /// <summary>Default metadata because no device can be queried in this build.</summary>
    public DeviceInfo DeviceInfo { get; } = new();

    /// <summary>
    /// Retained for compile-time compatibility. The fallback never produces
    /// device results.
    /// </summary>
    public event EventHandler<VerificationRecord>? ResultReceived
    {
        add { }
        remove { }
    }

    public DeviceSession(DeviceConfig config, VerificationXmlMap? map = null)
    {
        ArgumentNullException.ThrowIfNull(config);
        _ = map;
    }

    /// <summary>Compatibility no-op; the fallback does not own report paths.</summary>
    public void RegisterOwnedHybridPath(string path) => ArgumentNullException.ThrowIfNull(path);

    public Task ConnectAsync(CancellationToken ct = default) =>
        Task.FromException(new DeviceConnectionException(UnavailableMessage));

    /// <summary>Safe cleanup no-op when a connection was never established.</summary>
    public Task DisconnectAsync() => Task.CompletedTask;

    public Task RebootAndDisconnectAsync() => UnavailableTask();

    public Task<int> GetUpcEanSupplementalAsync(CancellationToken ct = default) =>
        UnavailableTask<int>();

    public Task<bool> SetUpcEanSupplementalAsync(int mode, CancellationToken ct = default) =>
        UnavailableTask<bool>();

    public Task<VerificationRecord?> TriggerAndGetResultAsync(
        VerificationRecord? sessionContext = null,
        CancellationToken ct = default) =>
        UnavailableTask<VerificationRecord?>();

    public Task<VerificationRecord?> ReplayAndGetResultAsync(
        VerificationRecord? sessionContext = null,
        int timeoutMs = 15_000,
        CancellationToken ct = default) =>
        UnavailableTask<VerificationRecord?>();

    public Task<VerificationRecord?> LoadImageAndVerifyAsync(
        string imagePath,
        int timeoutMs = 30_000,
        CancellationToken ct = default) =>
        UnavailableTask<VerificationRecord?>();

    public Task StartPushListenerAsync(
        VerificationRecord? sessionContext = null,
        CancellationToken ct = default) =>
        UnavailableTask();

    /// <summary>Safe cleanup no-op when a listener was never started.</summary>
    public Task StopPushListenerAsync() => Task.CompletedTask;

    public Task StartHttpSubscriberAsync(
        VerificationRecord? sessionContext = null,
        CancellationToken ct = default) =>
        UnavailableTask();

    /// <summary>Safe cleanup no-op when a subscriber was never started.</summary>
    public Task StopHttpSubscriberAsync() => Task.CompletedTask;

    public Task<byte[]?> GetRoiImageAsync(CancellationToken ct = default) =>
        UnavailableTask<byte[]?>();

    public Task<VerificationRecord?> LoadAndVerifyImageAsync(
        string imagePath,
        VerificationRecord? sessionContext = null,
        CancellationToken ct = default) =>
        UnavailableTask<VerificationRecord?>();

    public Task<string?> GetRawSymbolResultDiagnosticAsync(CancellationToken ct = default) =>
        UnavailableTask<string?>();

    /// <summary>Safe cleanup no-op for the stateless fallback.</summary>
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;

    private static Task UnavailableTask() =>
        Task.FromException(new PlatformNotSupportedException(UnavailableMessage));

    private static Task<T> UnavailableTask<T>() =>
        Task.FromException<T>(new PlatformNotSupportedException(UnavailableMessage));
}

/// <summary>Metadata about a connected device, populated by the SDK-backed session.</summary>
public sealed class DeviceInfo
{
    public string? Type { get; init; }
    public string? Serial { get; init; }
    public string? Name { get; init; }
    public string? FirmwareVersion { get; init; }
    public string? SoftwareVersion { get; init; }
    public DateTime? CalibrationDate { get; init; }
    public int? SensorWidthPx { get; init; }
    public int? SensorHeightPx { get; init; }
    public double? SensorPixelPitchUm { get; init; }
    public string? ImageSizeSetting { get; init; }
}

/// <summary>Thrown when a device connection cannot be established or is lost.</summary>
public sealed class DeviceConnectionException : Exception
{
    public DeviceConnectionException(string message, Exception? inner = null)
        : base(message, inner) { }
}