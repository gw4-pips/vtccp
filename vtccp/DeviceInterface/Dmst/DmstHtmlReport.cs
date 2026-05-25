namespace DeviceInterface.Dmst;

/// <summary>
/// Fields extracted from a DMST TruCheck HTML report for one scan.
///
/// HTML format confirmed 2026-05-25 from live scan (QR GUID, fw 6.1.16_sr4,
/// DM475-63530E-PIPS-Verif-Lab). Report file saved to:
///   {Documents}\{DeviceName}\CodeQuality\yyyy-MM-dd_HH-mm-ss-mmm_*.html
///
/// Structure: two distinct tables.
///   1. Simple characteristics table (label → next-cell value pairs):
///        "QR Size"/"29x29", "Error Correction Level"/"M",
///        "Data Mask Pattern"/"2", "Image"/"Black on white", "ECI"/"000003", etc.
///   2. Grade parameters table (6-cell rows per parameter):
///        [label][secondary][pct%][numeric_grade][letter][PASS/FAIL]
///
/// DateTime in the HTML header is CORRUPT (shows Unix epoch "31-Dec-1970").
/// Use filename timestamp (yyyy-MM-dd_HH-mm-ss prefix) for correlation.
///
/// Source routing:
///   Push XML fields  → DmstResultParser   → VerificationRecord (primary)
///   HTML-only fields → DmstHtmlScraper    → VerificationRecord (supplemental,
///                                            flagged in DataSourceExceptions)
///   Overlapping fields are cross-validated by DmstReportValidator.
/// </summary>
public sealed class DmstHtmlReport
{
    // ── Correlation ───────────────────────────────────────────────────────────

    /// <summary>
    /// Scan timestamp parsed from the HTML filename prefix (yyyy-MM-dd_HH-mm-ss).
    /// The in-page DateTime is corrupt (epoch). Used to correlate with the
    /// corresponding push XML VerificationRecord (±2s tolerance).
    /// </summary>
    public DateTime? ScanDateTime { get; init; }

    // ── Fields permanently unresolvable from push XML (primary targets) ───────

    /// <summary>
    /// QR Error Correction Level: L / M / Q / H.
    /// Label in HTML: "Error Correction Level". Confirmed value: "M" (2026-05-25).
    /// Maps to VerificationRecord.QR_ECLevel.
    /// </summary>
    public string? ECLevel { get; init; }

    /// <summary>
    /// QR Data Mask Pattern number (0–7).
    /// Label in HTML: "Data Mask Pattern". Confirmed value: "2" (2026-05-25).
    /// Maps to VerificationRecord.QR_MaskPattern.
    /// </summary>
    public string? DataMaskPattern { get; init; }

    /// <summary>
    /// ECI (Extended Channel Interpretation) assignment value, e.g. "000003".
    /// Label in HTML: "ECI". Confirmed value: "000003" (2026-05-25).
    /// Maps to VerificationRecord.QR_ECI.
    /// </summary>
    public string? ECI { get; init; }

    /// <summary>
    /// Image polarity as shown in DMST report.
    /// Label in HTML: "Image". Confirmed value: "Black on white" (2026-05-25).
    /// Other expected value: "White on black".
    /// Maps to VerificationRecord.ImagePolarity (enum).
    /// </summary>
    public string? ImagePolarity { get; init; }

    // ── Bonus: present in HTML, empty in push XML on fw 6.1.16_sr4 ───────────

    /// <summary>
    /// Data Codewords count.
    /// Label: "Data Codewords". Confirmed value: 44 (QR 29×29, ECL=M, 2026-05-25).
    /// Push XML tag &lt;DataCodewords&gt; is EMPTY on fw 6.1.16_sr4.
    /// HTML is authoritative — no C# table lookup needed.
    /// Maps to VerificationRecord.DataCodewords.
    /// </summary>
    public int? DataCodewords { get; init; }

    /// <summary>
    /// Error Correction Budget (total ECC codewords available).
    /// Label: "Error Correction Budget". Confirmed value: 26 (2026-05-25).
    /// Push XML tag &lt;ErrorCorrectionBudget&gt; is EMPTY on fw 6.1.16_sr4.
    /// Maps to VerificationRecord.ErrorCorrectionBudget.
    /// </summary>
    public int? ErrorCorrectionBudget { get; init; }

    /// <summary>
    /// Encoded Characters count from HTML.
    /// Label: "Encoded characters". Confirmed value: 36 (2026-05-25).
    /// Push XML eaLen fallback gives WRONG value (39 for this scan).
    /// HTML matches DMST display and is authoritative.
    /// Maps to VerificationRecord.EncodedCharacters (will also flag discrepancy).
    /// </summary>
    public int? EncodedCharacters { get; init; }

    /// <summary>
    /// Total Codewords (data + ECC). Label: "Total Codewords". Confirmed: 70.
    /// For cross-validation against push XML (push XML has this correctly).
    /// </summary>
    public int? TotalCodewords { get; init; }

    /// <summary>Errors Corrected. Label: "Errors Corrected". Cross-validation target.</summary>
    public int? ErrorsCorrected { get; init; }

    /// <summary>Error Capacity Used. Label: "Error Capacity Used". Cross-validation target.</summary>
    public int? ErrorCapacityUsed { get; init; }

    // ── Cross-validation targets (shared with push XML) ───────────────────────

    /// <summary>Overall grade letter (A/B/C/D/F), parsed from "4.0 (A)" display string.</summary>
    public string? OverallGrade { get; init; }

    /// <summary>Matrix dimensions, e.g. "29x29". Label: "QR Size".</summary>
    public string? MatrixSize { get; init; }

    /// <summary>Nominal X Dimension string, e.g. "12.6 mil". Label: "Nominal X Dim".</summary>
    public string? NominalXDim { get; init; }

    /// <summary>Horizontal BWG percentage (numeric). Label: "Horizontal BWG". Strips "%".</summary>
    public decimal? HorizontalBWG { get; init; }

    /// <summary>Vertical BWG percentage (numeric). Label: "Vertical BWG". Strips "%".</summary>
    public decimal? VerticalBWG { get; init; }

    /// <summary>UEC percentage. From grade parameters table (label + 2 cells = pct).</summary>
    public decimal? UECPercent { get; init; }

    /// <summary>ANU percentage. From grade parameters table.</summary>
    public decimal? ANUPercent { get; init; }

    /// <summary>GNU percentage. From grade parameters table.</summary>
    public decimal? GNUPercent { get; init; }

    /// <summary>SC percentage. "nan%" for IMAGE.LOAD scans → null (correct).</summary>
    public decimal? SCPercent { get; init; }

    // ── Parse provenance ──────────────────────────────────────────────────────

    /// <summary>
    /// Path to the HTML file this report was parsed from.
    /// Retained for diagnostics; file is deleted immediately after parsing.
    /// </summary>
    public string? SourceFilePath { get; init; }

    /// <summary>
    /// True if parsing succeeded and at least ScanDateTime was found.
    /// False if the structure was unrecognised (e.g. new DMST version changed layout).
    /// </summary>
    public bool ParseSucceeded { get; init; }

    /// <summary>Parser exception message, or null if parsing succeeded.</summary>
    public string? ParseError { get; init; }
}
