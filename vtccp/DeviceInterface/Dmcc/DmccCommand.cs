namespace DeviceInterface.Dmcc;

/// <summary>
/// Well-known DMCC command strings for the Cognex DataMan DMV device,
/// and utilities for sanitizing values before embedding them in commands.
///
/// DMCC command syntax: plain ASCII text, CRLF-terminated.
/// Example:  GET DEVICE.TYPE\r\n
///           SET DEVICE.NAME MyReader\r\n
/// </summary>
public static class DmccCommand
{
    // ── Query commands ────────────────────────────────────────────────────────

    /// <summary>Returns the device model string, e.g. "DM260Q".</summary>
    public const string GetDeviceType     = "GET DEVICE.TYPE";

    /// <summary>Returns the firmware version string, e.g. "5.7.4.0015".</summary>
    public const string GetFirmwareVer    = "GET FIRMWARE.VER";

    /// <summary>Returns the user-configurable device name.</summary>
    public const string GetDeviceName     = "GET DEVICE.NAME";

    /// <summary>Returns the device serial number / ID.</summary>
    public const string GetDeviceId       = "GET DEVICE.ID";

    /// <summary>Returns the last calibration date (format varies by firmware).</summary>
    public const string GetCalibrationDate = "GET CALIBRATION.DATE";

    // ── Trigger / result commands ─────────────────────────────────────────────

    /// <summary>
    /// Issues a single software trigger (capture + verify one symbol).
    /// Response is status only — no body.
    /// </summary>
    public const string Trigger           = "TRIGGER";

    /// <summary>
    /// Returns the verification result for the most recent scan as DMST XML.
    /// The full XML payload is the response body.
    /// </summary>
    public const string GetSymbolResult   = "GET SYMBOL.RESULT";

    /// <summary>
    /// Configures the result output format to the full XML report.
    /// Should be sent once after connection before polling GET SYMBOL.RESULT.
    /// </summary>
    public const string SetResultFormatFull = "SET DMCC.RESULT-FORMAT FULL";

    // ── Device control ────────────────────────────────────────────────────────

    /// <summary>Reboot the device. Use with caution in production.</summary>
    public const string Reboot            = "REBOOT";

    // ── Sanitization ──────────────────────────────────────────────────────────

    /// <summary>
    /// Characters that are illegal inside DMCC command argument strings.
    /// The DMCC wire protocol uses these as structural delimiters.
    /// </summary>
    private static readonly char[] _dmccIllegal =
        ['&', '<', '>', '"', '\r', '\n', '\0'];

    /// <summary>
    /// Removes or replaces characters that are illegal in a DMCC command argument.
    ///
    /// Note: '&amp;' is legal in Windows filenames but NOT in DMCC command strings —
    /// this is the "Phase 2 DMCC restriction" referenced in ExcelFileManager.  Call
    /// this method before embedding any user-supplied string in a SET command.
    ///
    /// Replacement strategy: illegal characters → '_'.
    /// </summary>
    public static string SanitizeForDmcc(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;

        var sb = new System.Text.StringBuilder(value.Length);
        foreach (char c in value)
            sb.Append(_dmccIllegal.Contains(c) ? '_' : c);

        return sb.ToString();
    }

    /// <summary>
    /// Builds a "SET {key} {value}" command with the value sanitized for DMCC.
    /// </summary>
    public static string SetValue(string key, string value) =>
        $"SET {key} {SanitizeForDmcc(value)}";
}
