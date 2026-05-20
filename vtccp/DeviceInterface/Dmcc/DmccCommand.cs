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

    // ── Trigger mode ──────────────────────────────────────────────────────────

    /// <summary>
    /// Returns the current trigger type string, e.g. "External", "Single", "Continuous".
    /// Used at session start to detect hardware-only trigger mode.
    /// </summary>
    public const string GetTriggerType = "GET TRIGGER.TYPE";

    /// <summary>
    /// Sets the trigger type to Single (software trigger enabled).
    /// Uses integer code 1 — some DMV firmware rejects the string form "Single".
    /// Allows VTCCP to fire scans via the TRIGGER command without a hardware signal.
    /// Saved in session only — DeviceSession restores the original value on disconnect.
    /// </summary>
    public const string SetTriggerTypeSingle = "SET TRIGGER.TYPE 1";

    // ── UPC/EAN Code Properties ───────────────────────────────────────────────

    /// <summary>
    /// Reads the current UPC/EAN supplemental (add-on) mode from firmware.
    /// Returns an integer string: 0=Ignore, 1=Parse, 2=Required,
    /// 3=Required 2-digit, 4=Required 5-digit, 5=Not Required.
    ///
    /// DMCC key: CODE.UPCEAN-SUPPLEMENT-DIGIT
    /// Matches DMST Code Details → UPC/EAN Properties → Supplementals dropdown.
    /// NOTE: Verify exact key name via a probe scan if firmware rejects it —
    ///       alternative candidates: CODE.UPCEAN.SUPPLEMENTAL, CODE.UPCEAN-ADDON.
    /// </summary>
    public const string GetUpcEanSupplemental = "GET CODE.UPCEAN-SUPPLEMENT-DIGIT";

    /// <summary>
    /// Writes the UPC/EAN supplemental mode to firmware (persistent — no explicit SAVE needed).
    /// <paramref name="mode"/> must be 0–5:
    ///   0 = Ignore, 1 = Parse, 2 = Required (any),
    ///   3 = Required 2-digit, 4 = Required 5-digit, 5 = Not Required.
    /// </summary>
    public static string SetUpcEanSupplemental(int mode) =>
        $"SET CODE.UPCEAN-SUPPLEMENT-DIGIT {mode}";

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
