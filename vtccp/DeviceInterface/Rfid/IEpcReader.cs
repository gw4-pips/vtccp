using DeviceInterface.Rfid.Models;

namespace DeviceInterface.Rfid;

/// <summary>
/// Hardware abstraction for an RFID reader capable of EPC inventory.
///
/// Implementations:
///   <see cref="AsReaderP35UEpcReader"/> — AsReader ASR-P35U (SDK DLL, VCP; canonical hardware).
///   <see cref="GoToTagsE310Reader"/>    — GoToTags Desktop E310 (FTDI VCP; rejected 2026-08).
///   <see cref="MtiLlcsEpcReader"/>      — MTI RU-824-100 (LLCS; discontinued hardware).
///
/// Use <see cref="EpcReaderFactory.CreateAsReaderP35U"/> for all new work.
/// </summary>
public interface IEpcReader : IAsyncDisposable
{
    /// <summary>True when the reader port is open and communication is established.</summary>
    bool IsConnected { get; }

    /// <summary>
    /// Open the reader on the specified serial port (e.g. "COM3" on Windows, "/dev/ttyUSB0" on Linux).
    /// Throws <see cref="InvalidOperationException"/> if already connected.
    /// Throws <see cref="IOException"/> if the port cannot be opened.
    /// </summary>
    Task ConnectAsync(string portName, CancellationToken ct = default);

    /// <summary>Gracefully closes the reader connection.</summary>
    Task DisconnectAsync(CancellationToken ct = default);

    /// <summary>
    /// Trigger a single inventory cycle: command the reader to scan for all tags in the field
    /// and return all unique EPCs detected within <paramref name="timeout"/>.
    ///
    /// Returns an empty list if no tags are detected before the timeout.
    /// Throws <see cref="InvalidOperationException"/> if not connected.
    /// </summary>
    Task<IReadOnlyList<EpcReadResult>> TriggerInventoryAsync(
        TimeSpan timeout,
        CancellationToken ct = default);

    /// <summary>
    /// Send a cancel command to stop any in-progress inventory cycle.
    /// Safe to call even when idle.
    /// </summary>
    Task CancelAsync(CancellationToken ct = default);
}
