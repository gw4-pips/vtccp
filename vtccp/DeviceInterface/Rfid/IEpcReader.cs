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

    /// <summary>
    /// Read the TID memory bank of a specific tag after inventory.
    ///
    /// Call this after <see cref="TriggerInventoryAsync"/> returns an EPC, before
    /// starting the next inventory.
    ///
    /// Default implementation returns null — concrete readers that support TID
    /// reading (e.g. <see cref="AsReaderP35UEpcReader"/>) override this.
    /// </summary>
    /// <param name="epcBytes">EPC bytes identifying the target tag.</param>
    /// <param name="timeout">Maximum wait for the TID result callback.</param>
    /// <returns>TID hex string (uppercase), or null on timeout / unsupported / error.</returns>
    Task<string?> ReadTidAsync(byte[] epcBytes, TimeSpan timeout, CancellationToken ct = default)
        => Task.FromResult<string?>(null);

    /// <summary>
    /// Query the EPC memory bank lock status of a specific tag after inventory
    /// (e.g. via the ASR-P35U SDK's CheckTagStatus command).
    ///
    /// Default implementation returns null — concrete readers that support a
    /// lock check (e.g. <see cref="AsReaderP35UEpcReader"/>) override this.
    /// </summary>
    /// <param name="epcBytes">EPC bytes identifying the target tag.</param>
    /// <param name="timeout">Maximum wait for the status result callback.</param>
    /// <returns>
    /// "PermaLocked" / "Locked" / "Unlocked" / "Unknown", or null when
    /// unsupported or the command was rejected.
    /// </returns>
    Task<string?> ReadLockStatusAsync(byte[] epcBytes, TimeSpan timeout, CancellationToken ct = default)
        => Task.FromResult<string?>(null);
}
