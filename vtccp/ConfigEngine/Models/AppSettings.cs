namespace ConfigEngine.Models;

/// <summary>
/// Application-wide preferences persisted alongside device profiles and job templates.
/// </summary>
public sealed class AppSettings
{
    // ── Defaults ──────────────────────────────────────────────────────────────

    /// <summary>Id of the <see cref="DeviceProfile"/> selected at last exit, or null.</summary>
    public string? LastDeviceProfileId { get; set; }

    /// <summary>Id of the <see cref="JobTemplate"/> selected at last exit, or null.</summary>
    public string? LastJobTemplateId { get; set; }

    // ── Output ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Root output directory for all sessions.
    /// Defaults to <c>%USERPROFILE%\Documents\VTCCP</c> on Windows,
    /// <c>~/VTCCP</c> on other platforms.
    /// </summary>
    public string DefaultOutputDirectory { get; set; } = GetDefaultOutputDirectory();

    // ── UI ────────────────────────────────────────────────────────────────────

    /// <summary>"Light" or "Dark". Applies to the WPF resource dictionary swap.</summary>
    public string Theme { get; set; } = "Light";

    /// <summary>MainWindow width remembered across sessions.</summary>
    public double WindowWidth { get; set; } = 1024;

    /// <summary>MainWindow height remembered across sessions.</summary>
    public double WindowHeight { get; set; } = 680;

    // ── Operator quick-access ─────────────────────────────────────────────────

    /// <summary>Operator ID typed at last session start (pre-fills the session launcher).</summary>
    public string? LastOperatorId { get; set; }

    // ── RFID cross-validation (Phase 0 POC) ──────────────────────────────────

    /// <summary>
    /// Serial port name for the AsReader ASR-P35U UHF RFID reader (e.g. "COM3").
    /// Null or empty = RFID feature disabled.
    /// </summary>
    public string? RfidComPort { get; set; }

    /// <summary>
    /// How long (ms) to hold the RFID scan window open after each barcode trigger.
    /// Valid range: 500–10000. Default: 3000 ms.
    /// </summary>
    public int RfidScanWindowMs { get; set; } = 3000;

    /// <summary>
    /// When true, an RFID cross-validation failure (GTIN/serial mismatch or no tag)
    /// is recorded as a soft flag in the RFID Scans sheet.
    /// The barcode grade is never altered.
    /// </summary>
    public bool RfidFlagMismatch { get; set; } = true;

    /// <summary>
    /// Path to the local GS1 GCP Prefix Format List XML file.
    /// Defaults to &lt;DefaultOutputDirectory&gt;\data\gcp-prefix-format-list.xml.
    /// Set by GcpUpdateService on first download.
    /// </summary>
    public string? GcpDataPath { get; set; }

    /// <summary>
    /// Last-known date string from the GCPPrefixFormatList root element ("date" attribute).
    /// Stored so the update check can compare without re-parsing the full 8 MB file.
    /// </summary>
    public string? GcpLastModified { get; set; }

    // ── Schema version awareness ──────────────────────────────────────────────

    /// <summary>App version that wrote this settings file.</summary>
    public string AppVersion { get; set; } = "1.0.0";

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string GetDefaultOutputDirectory()
    {
        string home = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        if (string.IsNullOrEmpty(home))
            home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return Path.Combine(home, "VTCCP");
    }
}
