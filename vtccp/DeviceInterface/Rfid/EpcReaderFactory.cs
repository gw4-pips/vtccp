namespace DeviceInterface.Rfid;

/// <summary>
/// Creates <see cref="IEpcReader"/> instances.
/// Call <see cref="CreateGoToTagsE310"/> for the GoToTags Desktop E310 USB UHF reader.
/// Call <see cref="CreateMtiLlcs"/> for the (superseded) MTI RU-824-100 reader.
/// </summary>
public static class EpcReaderFactory
{
    /// <summary>
    /// Create an <see cref="IEpcReader"/> for the GoToTags Desktop E310 UHF RFID Reader
    /// (Impinj E310 chipset, SKU TDLP3LCFPP).
    ///
    /// Prerequisites:
    ///   1. FTDI driver installed (CDM212364_Setup.zip from the GoToTags GitLab repo).
    ///   2. Reader plugged in and appearing as a COM port in Device Manager.
    ///   3. Pass the assigned port name to <see cref="IEpcReader.ConnectAsync"/>.
    ///
    /// Protocol: GoToTags UHF RFID Reader Communication Protocol rev 5-30-23.
    /// Baud rate: 115 200 8N1 (factory default).
    /// </summary>
    public static IEpcReader CreateGoToTagsE310() => new GoToTagsE310Reader();

    /// <summary>
    /// Create an <see cref="IEpcReader"/> for the MTI RU-824-100 using the
    /// LLCS binary packet protocol over a serial/USB-VCP port.
    /// NOTE: MTI RU-824-100 is discontinued. Prefer <see cref="CreateGoToTagsE310"/>.
    /// </summary>
    public static IEpcReader CreateMtiLlcs() => new MtiLlcsEpcReader();

    /// <summary>
    /// Enumerate the serial port names available on this machine.
    /// On Windows returns "COM1", "COM2", etc.
    /// </summary>
    public static IReadOnlyList<string> GetAvailablePorts() =>
        System.IO.Ports.SerialPort.GetPortNames();
}
