namespace ConfigEngine.Models;

/// <summary>
/// Controls how the hybrid HTML report interacts with the original Webscan HTML output.
/// </summary>
public enum HybridReportMode
{
    /// <summary>
    /// Keep the Webscan HTML in the CodeQuality folder and write the hybrid HTML
    /// to the session output directory (or <see cref="AppSettings.HybridReportOutputDirectory"/>).
    /// This is the default — both files coexist.
    /// </summary>
    Alongside,

    /// <summary>
    /// After the DmstHtmlScraper parses and removes the Webscan HTML, write the
    /// hybrid HTML back to the same path (same folder, same filename) so downstream
    /// tools watching the CodeQuality folder see only the enriched report.
    /// </summary>
    Replace,
}

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

    // ── Hybrid report ─────────────────────────────────────────────────────────

    /// <summary>
    /// When true, VTCCP generates a self-contained hybrid HTML report for every
    /// accepted scan.  The report merges the Webscan TruCheck™ barcode grading
    /// data with the VCCS FlexWedge™ RFID validation result.  It can be
    /// opened in any browser and printed directly to PDF.
    /// </summary>
    public bool GenerateHybridReport { get; set; } = true;

    /// <summary>
    /// Optional directory for hybrid HTML reports.
    /// Null or empty = session output directory (same folder as the Excel workbook).
    /// Ignored when <see cref="HybridReportMode"/> is <see cref="HybridReportMode.Replace"/>,
    /// because the hybrid is written to the Webscan CodeQuality folder in that mode.
    /// </summary>
    public string? HybridReportOutputDirectory { get; set; }

    /// <summary>
    /// Controls whether the hybrid report is written alongside the Webscan original
    /// (Alongside, default) or replaces it in the CodeQuality folder (Replace).
    /// Only relevant when <see cref="GenerateHybridReport"/> is true.
    /// </summary>
    public HybridReportMode HybridReportMode { get; set; } = HybridReportMode.Alongside;

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
