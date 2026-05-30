namespace DeviceInterface.Dmcc;

/// <summary>
/// Well-known DMCC command strings for the Cognex DataMan DMV device,
/// confirmed against the DataMan DMCC Reference 6.1.16_sr4 (2026-04-21).
///
/// DMCC wire protocol (raw TCP):
///   Command format:  ||>{command}\r\n
///   e.g.  ||>GET DEVICE.TYPE\r\n
///         ||>TRIGGER ON\r\n
///   The "||>" is the bare header (full form: ||checksum:command-id>; all fields optional).
///
///   Response format (Extended mode, COM.DMCC-RESPONSE=2):
///     ||[0]\r\n                    ← status 0 = OK
///     ||[101]\r\n                  ← invalid command
///     ||[102]\r\n                  ← invalid parameter
///     ||[104]\r\n                  ← parameter rejected (reader state)
///   Silent mode (default, COM.DMCC-RESPONSE=0): no ACK at all — set Extended immediately after banner.
///
/// When commands are sent via the Cognex SDK (DataManSdkClient.SendAsync), the SDK
/// adds the header/framing automatically.  Raw TCP callers must prepend WireHeader.
/// </summary>
public static class DmccCommand
{
    // ── Port constants ─────────────────────────────────────────────────────────

    /// <summary>
    /// Raw DMCC text interface (classic Telnet/DMCC port).
    /// <c>||&gt;COMMAND\r\n</c> works here.  Confirmed on DM475V fw 6.1.16_sr4.
    /// Use this port for any transient raw-TCP DMCC session (e.g. COM.DMCC-RESET,
    /// TRIGGER ON) — the device accepts and ACKs commands on this port.
    /// </summary>
    public const int RawDmccPort = 23;

    /// <summary>
    /// DataMan SDK / HTTP event-subscription port.
    /// Multiplexes the Cognex SDK binary protocol and HTTP pub-sub
    /// (GET /events?enable → PUT /codes.xml).
    /// Bare TCP connections sending <c>||&gt;COMMAND\r\n</c> on this port are
    /// silently ignored — the device does not recognise them as DMCC sessions
    /// and returns zero bytes.  Do NOT send raw DMCC text here.
    /// </summary>
    public const int SdkHttpPort = 44_444;

    // ── System / device identity ───────────────────────────────────────────────

    /// <summary>Returns the device model string, e.g. "DM475V". ALL platforms.</summary>
    public const string GetDeviceType     = "GET DEVICE.TYPE";

    /// <summary>Returns the firmware version string, e.g. "6.1.16.0015". ALL platforms.</summary>
    public const string GetFirmwareVer    = "GET DEVICE.FIRMWARE-VER";

    /// <summary>Returns the user-configurable device name. ALL platforms.</summary>
    public const string GetDeviceName     = "GET DEVICE.NAME";

    /// <summary>Returns the device serial number. ALL platforms.</summary>
    public const string GetDeviceSerialNumber = "GET DEVICE.SERIAL-NUMBER";

    /// <summary>
    /// Returns a comma-delimited string of installed feature keys (e.g. OCR key).
    /// GET only. ALL platforms.
    /// </summary>
    public const string GetFeatureKeys    = "GET DEVICE.FEATURE-KEYS";

    // ── Camera / image ────────────────────────────────────────────────────────

    /// <summary>
    /// Returns (or sets) the output downscale factor used by IMAGE.SEND.
    /// Values: 0 = Full, 1 = 1/4, 2 = 1/16, 3 = 1/64. ALL platforms. Version 4.4.0+.
    /// Note: this controls IMAGE.SEND output only — it does NOT affect the
    /// firmware ROI crop carried in push XML JpegImageBase64.
    /// </summary>
    public const string GetImageSize = "GET IMAGE.SIZE";

    /// <summary>
    /// Requests the device to transmit the current image buffer as a binary JPEG.
    ///
    /// Three-level image stack (DataMan DM475V / DM395V):
    ///   Level 1 — Barcode crop  : r.trucheck.jpegImage in push XML (200–600 px).
    ///             Tight crop around the decoded symbol.  Same image shown in
    ///             the DMST verification panel and embedded in VTCCP Excel cells.
    ///             VTCCP already captures this in VerificationRecord.JpegImageBase64.
    ///
    ///   Level 2 — ROI frame     : IMAGE.SEND after a live scan (this command).
    ///             The operator-configured Region Of Interest rectangle in DMST.
    ///             Wider than the barcode crop; includes surrounding label area
    ///             including adjacent human-readable text / lot numbers.
    ///             Resolution depends on IMAGE.SIZE setting (0=Full through 3=1/64).
    ///             OPEN QUESTION: confirmed via IMAGE.SEND test on DM475V? → probe in D4.
    ///
    ///   Level 3 — Full frame    : DataManSystem.GetLastReadImage() (SDK method).
    ///             Full 2448×2048 (DM475V / DM395V) or 2048×1536 (DM390 / DM394).
    ///             Always available after any scan regardless of IMAGE.SIZE setting.
    ///             Not yet exploited in VTCCP.
    ///
    /// Response is raw binary (JPEG bytes) via SendCommandWithExpectedBinaryResult().
    /// The IMAGE.SIZE setting controls downscale applied before transmission.
    ///
    /// Platforms: ALL. Version: confirmed in DMCC 6.1.16_sr4 digest.
    ///
    /// OCR targeting guidance:
    ///   Barcode crop (Level 1) — adequate for HRI text directly adjacent to the symbol.
    ///   ROI frame (Level 2)    — required for lot / expiry / IUID text outside the
    ///                            barcode's immediate area; preferred for label OCR.
    ///   Full frame (Level 3)   — maximum context; use for layout analysis or when
    ///                            ROI coverage is unknown.
    /// </summary>
    public const string ImageSend = "IMAGE.SEND";

    /// <summary>
    /// Mirrors the acquired image over the vertical axis (= Flip Horizontal in DMST UI).
    /// SET/GET. Values: ON / OFF.
    /// DMST path: Advanced → Imager Settings → Image Mirroring → Flip Horizontal.
    /// Platforms: DM80, DM280, DM290, DM300, DM360, DM370, DM390, DM470, DM503.
    /// DM475V confirmed via DMST UI — probe before relying on firmware support.
    /// Version: 5.2.0.
    /// </summary>
    public const string GetMirrorHorizontal  = "GET CAMERA.MIRROR-HORIZONTAL";
    public const string SetMirrorHorizontalOn  = "SET CAMERA.MIRROR-HORIZONTAL ON";
    public const string SetMirrorHorizontalOff = "SET CAMERA.MIRROR-HORIZONTAL OFF";

    /// <summary>
    /// Mirrors the acquired image over the horizontal axis (= Flip Vertical in DMST UI).
    /// SET/GET. Values: ON / OFF.
    /// DMST path: Advanced → Imager Settings → Image Mirroring → Flip Vertical.
    /// Platforms: DM70, DM80, DM260, DM280, DM290, DM300, DM360, DM370, DM390, DM470, DM503.
    /// DM475V confirmed via DMST UI — probe before relying on firmware support.
    /// Version: 5.2.0.
    /// </summary>
    public const string GetMirrorVertical  = "GET CAMERA.MIRROR-VERTICAL";
    public const string SetMirrorVerticalOn  = "SET CAMERA.MIRROR-VERTICAL ON";
    public const string SetMirrorVerticalOff = "SET CAMERA.MIRROR-VERTICAL OFF";

    // ── Raw TCP wire helpers ───────────────────────────────────────────────────

    /// <summary>
    /// Bare DMCC command header for raw TCP connections.
    /// All commands sent on raw TCP must be prefixed: WireHeader + command + "\r\n".
    /// The Cognex SDK adds its own framing automatically — do NOT use this prefix
    /// when calling DataManSdkClient.SendAsync / _system.SendCommand.
    /// </summary>
    public const string WireHeader = "||>";

    /// <summary>
    /// Switches the DMCC response mode to Extended (mode 2).
    /// Send this immediately after draining the welcome banner on a raw TCP connection.
    /// In Extended mode every command returns ||[status]\r\n — enables reliable ACK detection.
    /// Default at connect time is Silent (mode 0): no ACK for any command.
    /// Wire form: ||>SET COM.DMCC-RESPONSE 2\r\n
    /// </summary>
    public const string SetDmccResponseExtended = "SET COM.DMCC-RESPONSE 2";

    // ── Trigger / result ──────────────────────────────────────────────────────

    /// <summary>
    /// Fires a single software trigger (capture + verify one symbol).
    /// DMCC wire form (confirmed from DMCC Reference): ||>TRIGGER ON\r\n
    /// Extended-mode response: ||[0]\r\n  (then result via HTTP push or XmlResultArrived).
    /// The SDK's SendCommand may also accept "TRIGGER ON" — try that path first.
    /// </summary>
    public const string TriggerOn  = "TRIGGER ON";

    /// <summary>Cancels a held Manual trigger. Wire form: ||>TRIGGER OFF\r\n</summary>
    public const string TriggerOff = "TRIGGER OFF";

    /// <summary>
    /// Bare "TRIGGER" — kept only for SDK compatibility probes.
    /// The correct DMCC form is TriggerOn ("TRIGGER ON").
    /// </summary>
    public const string Trigger = "TRIGGER";

    /// <summary>
    /// Returns the verification result for the most recent scan as DMST XML.
    /// The full XML payload is the response body.
    /// </summary>
    public const string GetSymbolResult   = "GET SYMBOL.RESULT";

    /// <summary>
    /// Replays the most recently loaded image through TruCheck verification.
    /// Sent after the SDK's LoadImage() call to fire a full grading pass on
    /// the stored image buffer.  Result arrives via XmlResultArrived event.
    ///
    /// D4 Image Load sequence (confirmed from scans #11/#13, 2026-05-24):
    ///   1. DataManSdkClient.LoadAndReplayImageAsync(filePath)
    ///      a. SDK LoadImage(Bitmap) via reflection — loads pixels into device buffer
    ///      b. DMCC IMAGE.REPLAY — fires TruCheck verification on loaded image
    ///      c. Await XmlResultArrived — parse result (OpticsSource = "LoadedImage")
    ///
    /// Platforms: DM475V + DMV-8072V. Version: confirmed fw 6.1.16_sr4.
    /// </summary>
    public const string ImageReplay = "IMAGE.REPLAY";

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

    // ── UPC/EAN supplemental ──────────────────────────────────────────────────

    /// <summary>
    /// Reads the current UPC/EAN supplemental (add-on) mode from firmware.
    /// Returns an integer string: 0=Ignore, 1=Required, 2=Required 2-digit,
    /// 3=Required 5-digit, 4=Not Required.
    ///
    /// DMCC key confirmed: UPC-EAN.SUPPLEMENT (DMCC Reference 6.1.16_sr4).
    /// ALL platforms. Version: 3.0.0.
    /// </summary>
    public const string GetUpcEanSupplemental = "GET UPC-EAN.SUPPLEMENT";

    /// <summary>
    /// Writes the UPC/EAN supplemental mode to firmware (persistent).
    /// <paramref name="mode"/> must be 0–4:
    ///   0 = Ignore
    ///   1 = Required (any)
    ///   2 = Required 2-digit
    ///   3 = Required 5-digit
    ///   4 = Not Required (optional — decode if present)
    /// </summary>
    public static string SetUpcEanSupplemental(int mode) =>
        $"SET UPC-EAN.SUPPLEMENT {mode}";

    // ── TruCheck — application standard & grading ────────────────────────────
    // All TRUCHECK.* commands below: DM475V + DM8072V only. Version: 6.1.10.

    /// <summary>
    /// Gets/sets the active Application Standard.
    /// 0=GS1, 1=HIBCC, 2=UDI (HIBCC+GS1), 3=UID, 4=Auto, 5=Custom, 6=Cryptocode.
    /// </summary>
    public const string GetApplicationStandard  = "GET TRUCHECK.APPLICATION-STANDARD";
    public static string SetApplicationStandard(int std) => $"SET TRUCHECK.APPLICATION-STANDARD {std}";

    /// <summary>
    /// Gets/sets the grading standard.
    /// 0=ISO/IEC 15415/6, 1=ISO/IEC 29158:2020.
    /// </summary>
    public const string GetGradingStandard  = "GET TRUCHECK.GRADING-STANDARD";
    public static string SetGradingStandard(int std) => $"SET TRUCHECK.GRADING-STANDARD {std}";

    /// <summary>
    /// Gets/sets the verification aperture type when Application Standard = Custom.
    /// 0=User Set, 1=Auto 80%/50%, 2=Auto aperture.
    /// </summary>
    public const string GetAperture  = "GET TRUCHECK.APERTURE";
    public static string SetAperture(int mode) => $"SET TRUCHECK.APERTURE {mode}";

    /// <summary>
    /// Gets/sets the aperture size (User Set mode only).
    /// Integer [1–300] in units of ten-thousandths of an inch.
    /// </summary>
    public const string GetApertureSize  = "GET TRUCHECK.APERTURE-SIZE";
    public static string SetApertureSize(int tenThousandthsInch) => $"SET TRUCHECK.APERTURE-SIZE {tenThousandthsInch}";

    /// <summary>
    /// Gets/sets the data parsing standard when Application Standard = Custom.
    /// 0=None, 1=GS1, 2=HIBCC, 3=UID.
    /// </summary>
    public const string GetCustomDataParsingStandard = "GET TRUCHECK.APPLICATION-CUSTOM-DATA-PARSING-STANDARD";
    public static string SetCustomDataParsingStandard(int std) =>
        $"SET TRUCHECK.APPLICATION-CUSTOM-DATA-PARSING-STANDARD {std}";

    /// <summary>Pass grade threshold when Application Standard = Custom. [0–40], without decimal.</summary>
    public const string GetCustomPassGrade  = "GET TRUCHECK.APPLICATION-CUSTOM-PASS-GRADE";
    public static string SetCustomPassGrade(int grade) => $"SET TRUCHECK.APPLICATION-CUSTOM-PASS-GRADE {grade}";

    /// <summary>Minimum X-Dimension when Application Standard = Custom. [1–1000] thousandths of an inch.</summary>
    public const string GetCustomMinXDim  = "GET TRUCHECK.APPLICATION-CUSTOM-MINIMUM-X-DIM";
    public static string SetCustomMinXDim(int mils) => $"SET TRUCHECK.APPLICATION-CUSTOM-MINIMUM-X-DIM {mils}";

    /// <summary>Maximum X-Dimension when Application Standard = Custom. [1–1000] thousandths of an inch.</summary>
    public const string GetCustomMaxXDim  = "GET TRUCHECK.APPLICATION-CUSTOM-MAXIMUM-X-DIM";
    public static string SetCustomMaxXDim(int mils) => $"SET TRUCHECK.APPLICATION-CUSTOM-MAXIMUM-X-DIM {mils}";

    /// <summary>GS1 table index when Application Standard = GS1. [0–11]; 0=Auto.</summary>
    public const string GetGS1Table  = "GET TRUCHECK.APPLICATION-GS1-TABLE";
    public static string SetGS1Table(int table) => $"SET TRUCHECK.APPLICATION-GS1-TABLE {table}";

    // ── TruCheck — report header fields ──────────────────────────────────────

    /// <summary>Operator name shown in report headers. String. DM475V + DM8072V.</summary>
    public const string GetOperatorName  = "GET TRUCHECK.OPERATOR-NAME";
    public static string SetOperatorName(string name) =>
        $"SET TRUCHECK.OPERATOR-NAME {SanitizeForDmcc(name)}";

    /// <summary>
    /// Custom note shown in report headers. String. DM475V + DM8072V.
    /// Maps to VTCCP SessionManager.CustomNote.
    /// </summary>
    public const string GetCustomNote  = "GET TRUCHECK.CUSTOM-NOTE";
    public static string SetCustomNote(string note) =>
        $"SET TRUCHECK.CUSTOM-NOTE {SanitizeForDmcc(note)}";

    /// <summary>
    /// Batch number shown in report headers when Auto-Batch is OFF. String. DM475V + DM8072V.
    /// Maps to VTCCP SessionManager.BatchNumber.
    /// </summary>
    public const string GetBatchNumber  = "GET TRUCHECK.BATCH-NUMBER";
    public static string SetBatchNumber(string batch) =>
        $"SET TRUCHECK.BATCH-NUMBER {SanitizeForDmcc(batch)}";

    /// <summary>
    /// When ON, firmware auto-increments the batch number in report headers.
    /// When OFF, TRUCHECK.BATCH-NUMBER is used as-is. DM475V + DM8072V.
    /// </summary>
    public const string GetAutoBatch  = "GET TRUCHECK.AUTO-BATCH";
    public const string SetAutoBatchOn  = "SET TRUCHECK.AUTO-BATCH ON";
    public const string SetAutoBatchOff = "SET TRUCHECK.AUTO-BATCH OFF";

    // ── TruCheck — calibration ────────────────────────────────────────────────

    /// <summary>
    /// Begins calibration using a conformance standard test card.
    /// RMax and RMin must match the test card values. Read the calibration symbol
    /// once calibration is started, then send TRUCHECK.CALIBRATE-OFF.
    /// DM475V + DM8072V.
    /// </summary>
    public static string CalibrateOn(double rMax, double rMin) =>
        $"TRUCHECK.CALIBRATE-ON {rMax:F1} {rMin:F1}";

    /// <summary>
    /// Begins calibration using any symbol as the calibration target.
    /// RMax, RMin, and XDimension must match the symbol. DM475V + DM8072V.
    /// </summary>
    public static string CalibrateCustomOn(double rMax, double rMin, double xDim) =>
        $"TRUCHECK.CALIBRATE-CUSTOM-ON {rMax:F1} {rMin:F1} {xDim:F1}";

    /// <summary>Ends the calibration process. Must be called after reading the calibration target.</summary>
    public const string CalibrateOff = "TRUCHECK.CALIBRATE-OFF";

    /// <summary>
    /// Company name shown in report headers. String. DM475V + DM8072V.
    /// Maps to VTCCP SessionManager.CompanyName.
    /// </summary>
    public const string GetCompanyName  = "GET TRUCHECK.COMPANY-NAME";
    public static string SetCompanyName(string name) =>
        $"SET TRUCHECK.COMPANY-NAME {SanitizeForDmcc(name)}";

    // ── TruCheck — display / misc ─────────────────────────────────────────────

    /// <summary>Enables dot-peen stick algorithm when grading with ISO/IEC 29158. DM475V + DM8072V.</summary>
    public const string GetDotPeen  = "GET TRUCHECK.DOT-PEEN";
    public const string SetDotPeenOn  = "SET TRUCHECK.DOT-PEEN ON";
    public const string SetDotPeenOff = "SET TRUCHECK.DOT-PEEN OFF";

    /// <summary>Distance units. OFF=Standard (mils/inches), ON=Metric (µm/mm). DM475V + DM8072V.</summary>
    public const string GetMetricUnits  = "GET TRUCHECK.METRIC-UNITS";
    public const string SetMetricUnitsOn  = "SET TRUCHECK.METRIC-UNITS ON";
    public const string SetMetricUnitsOff = "SET TRUCHECK.METRIC-UNITS OFF";

    /// <summary>
    /// Enables or disables a named section in the firmware Verification Report.
    /// section values: "CODE_IMAGE", "GENERAL-CHARACTERISTICS-TABLE",
    ///   "QUALITY-DETAIL-TABLE", "MODULATION-TABLE", "ENCODATION-DETAIL-TABLE",
    ///   "ASCII_TABLE", "APPLICATION-DATA-TABLE", "CODEWORD-TABLE".
    /// DM280, DM370, DM470, DM475V, DM8072V. Version: 5.7.10 SR1.
    /// </summary>
    public static string SetReportSection(string section, bool enabled) =>
        $"TRUCHECK.REPORT-SECTION \"{section}\" {(enabled ? "ON" : "OFF")}";

    // ── QR quality metrics mode ───────────────────────────────────────────────

    /// <summary>
    /// Gets/sets QR Code quality metrics mode. ALL platforms. Version: 4.5.0.
    /// 0=None, 1=ISO/IEC 15415, 2=AIM-DPM / ISO/IEC TR 29158.
    /// </summary>
    public const string GetQrQualityMetrics = "GET QR.QUALITY-METRICS";
    public static string SetQrQualityMetrics(int mode) => $"SET QR.QUALITY-METRICS {mode}";

    // ── Device time / NTP ────────────────────────────────────────────────────

    /// <summary>
    /// Gets/sets the timezone for local time on the device.
    /// Argument is a Location string (e.g. "America/New_York") or POSIX timezone.
    /// DMST path: System Settings → Device Time Settings → Time Zone / Value.
    /// Platforms: All. Version: 5.7.0.
    /// After CONFIG.DEFAULT: must be restored to "America/New_York" (PIPS-Verif-Lab default).
    /// </summary>
    public const string GetTimezone = "GET DEVICE.TIMEZONE";
    public static string SetTimezone(string tz) => $"SET DEVICE.TIMEZONE {tz}";

    /// <summary>
    /// Enables or disables NTP time synchronisation.
    /// SET/GET. Values: ON / OFF.
    /// DMST path: System Settings → Device Time Settings → Enable NTP.
    /// Platforms: All. Version: 5.7.0.
    /// After CONFIG.DEFAULT: restore to ON.
    /// </summary>
    public const string GetNtpEnable  = "GET NTP.ENABLE";
    public const string SetNtpEnableOn  = "SET NTP.ENABLE ON";
    public const string SetNtpEnableOff = "SET NTP.ENABLE OFF";

    /// <summary>
    /// Primary NTP server address (IP or DNS hostname).
    /// DMST path: System Settings → Device Time Settings → NTP Server 1.
    /// Platforms: All. Version: 5.7.0.
    /// After CONFIG.DEFAULT: restore to "time.nist.gov" (PIPS-Verif-Lab default).
    /// </summary>
    public const string GetNtpServer1 = "GET NTP.SERVER1";
    public static string SetNtpServer1(string server) => $"SET NTP.SERVER1 {server}";

    /// <summary>
    /// Secondary NTP server address (IP or DNS hostname). Empty string to clear.
    /// DMST path: System Settings → Device Time Settings → NTP Server 2.
    /// Platforms: All. Version: 5.7.0.
    /// After CONFIG.DEFAULT: leave empty (PIPS-Verif-Lab default).
    /// </summary>
    public const string GetNtpServer2 = "GET NTP.SERVER2";
    public static string SetNtpServer2(string server) => $"SET NTP.SERVER2 {server}";

    // ── Device control ────────────────────────────────────────────────────────

    /// <summary>Reboot the device. Use with caution in production.</summary>
    public const string Reboot = "REBOOT";

    // ── DMCC session settings ─────────────────────────────────────────────────

    /// <summary>
    /// Resets DMCC connection settings back to firmware defaults for all future
    /// connections.  Affected settings: COM.DMCC-RESPONSE, COM.DMCC-CHECKSUM,
    /// COM.DMCC-HEADER, DATA.RESULT-TYPE, DATA.RESULT-ENCODING,
    /// DATA.RESULT-ALWAYSSEND, DATA.IMAGE-TYPE.
    ///
    /// VTCCP sends this command during DisconnectAsync() to ensure no DMCC
    /// configuration change from a prior session persists on the device and
    /// blanks DMST's image panel or otherwise affects other clients.
    ///
    /// Root-cause note: the Cognex SDK's SetResultTypes() internally calls
    /// COM.DMCC-SAVE, which persists DATA.IMAGE-TYPE to a value that strips
    /// images from the result-delivery channel.  This condition survives
    /// disconnect and is only cleared by a physical reboot — or by this command.
    ///
    /// Platforms: ALL. Version: 4.4.0.
    /// </summary>
    public const string DmccReset = "COM.DMCC-RESET";

    /// <summary>
    /// Saves the current DMCC connection settings as the defaults for all future
    /// connections (same set of settings as DmccReset restores).
    ///
    /// NOT called by VTCCP — documented here for reference. The SDK may call this
    /// internally when SetResultTypes() is used, which is why SetResultTypes() is
    /// intentionally avoided in DataManSdkClient.
    ///
    /// Platforms: ALL. Version: 4.4.0.
    /// </summary>
    public const string DmccSave  = "COM.DMCC-SAVE";

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
    /// this is the "Phase 2 DMCC restriction" referenced in ExcelFileManager. Call
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
