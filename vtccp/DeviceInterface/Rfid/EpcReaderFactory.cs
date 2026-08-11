namespace DeviceInterface.Rfid;

/// <summary>
/// Creates <see cref="IEpcReader"/> instances.
///
/// Canonical hardware: AsReader ASR-P35U — use <see cref="CreateAsReaderP35U"/>.
///
/// Superseded readers (kept for reference; do not use for new work):
///   <see cref="CreateGoToTagsE310"/> — GoToTags Desktop E310 (rejected 2026-08).
///   <see cref="CreateMtiLlcs"/>      — MTI RU-824-100 (discontinued hardware).
/// </summary>
public static class EpcReaderFactory
{
    /// <summary>
    /// Create an <see cref="IEpcReader"/> for the <b>AsReader ASR-P35U</b> UHF RFID Reader
    /// (canonical VCCS hardware, confirmed 2026-08-06).
    ///
    /// Prerequisites:
    ///   1. <c>AsReaderP3xU.dll</c> (SDK v1.3.0) placed in
    ///      <c>vtccp\lib\asreader-p3xu-sdk-1.3.0\</c> — see <c>PLACE-DLL-HERE.md</c>.
    ///   2. Reader connected via USB; COM port assigned by Windows Device Manager
    ///      (VID=0x339C / PID=0x271B).
    ///   3. Pass the COM port name to <see cref="IEpcReader.ConnectAsync"/>
    ///      (e.g. "COM4").  Use <see cref="AsReaderP35UEpcReader.GetAvailablePorts"/>
    ///      to enumerate candidates.
    /// </summary>
    /// <param name="txPowerDbm">TX power in dBm (13–27). Default 20 dBm.</param>
    public static IEpcReader CreateAsReaderP35U(int txPowerDbm = 20)
        => new AsReaderP35UEpcReader(txPowerDbm);

    /// <summary>
    /// [Superseded — do not use for new work]
    /// Create an <see cref="IEpcReader"/> for the GoToTags Desktop E310.
    /// The E310 was evaluated and rejected in favour of the ASR-P35U (2026-08).
    /// </summary>
    [Obsolete("GoToTags E310 rejected. Use CreateAsReaderP35U() instead.")]
    public static IEpcReader CreateGoToTagsE310() => new GoToTagsE310Reader();

    /// <summary>
    /// [Superseded — do not use for new work]
    /// Create an <see cref="IEpcReader"/> for the MTI RU-824-100 (LLCS protocol).
    /// The MTI RU-824-100 is discontinued hardware superseded by the ASR-P35U.
    /// </summary>
    [Obsolete("MTI RU-824-100 discontinued. Use CreateAsReaderP35U() instead.")]
    public static IEpcReader CreateMtiLlcs() => new MtiLlcsEpcReader();

    /// <summary>
    /// Enumerate the serial port names available on this machine.
    /// On Windows returns "COM1", "COM2", etc.
    /// For ASR-P35U-aware enumeration, use
    /// <see cref="AsReaderP35UEpcReader.GetAvailablePorts"/> instead.
    /// </summary>
    public static IReadOnlyList<string> GetAvailablePorts() =>
        System.IO.Ports.SerialPort.GetPortNames();
}
