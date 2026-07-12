namespace DeviceInterface.Rfid;

/// <summary>
/// Creates <see cref="IEpcReader"/> instances.
/// Call <see cref="CreateMtiLlcs"/> for the MTI RU-824-100 USB UHF reader.
/// </summary>
public static class EpcReaderFactory
{
    /// <summary>
    /// Create an <see cref="IEpcReader"/> for the MTI RU-824-100 using the
    /// LLCS binary packet protocol over a serial/USB-VCP port.
    /// </summary>
    public static IEpcReader CreateMtiLlcs() => new MtiLlcsEpcReader();

    /// <summary>
    /// Enumerate the serial port names available on this machine.
    /// On Windows returns "COM1", "COM2", etc.
    /// </summary>
    public static IReadOnlyList<string> GetAvailablePorts() =>
        System.IO.Ports.SerialPort.GetPortNames();
}
