namespace DeviceInterface.Rfid;

/// <summary>
/// Maps the direct return value from the ASR-P35U SDK's
/// <c>CheckTagStatus(byte[] epc)</c> call.
/// </summary>
internal static class AsReaderLockStatus
{
    public static string FromCheckTagStatus(uint statusCode) => statusCode switch
    {
        0 => "Unlocked",
        1 => "Locked",
        2 => "PermaLocked",
        _ => "Unknown",
    };
}