namespace DeviceInterface.Dmst;

using ExcelEngine.Models;

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

    // ── Multi-mode (linear symbol) ────────────────────────────────────────────
    // Populated when the HTML report covers a multi-mode scan (EAN/UPC + 2D).
    // All fields are null for single-mode scans.
    //
    // Detection: one of the KnownLinearSymbologies cell values ("EAN-13", "EAN-8",
    // "UPC-A", "UPC-E") found in the cell list signals a multi-mode report.
    //
    // In the Webscan TruCheck multi-mode HTML layout the linear (1D) section appears
    // before the 2D section.  Parsing relies on this ordering to assign the first
    // "D.D (L)" grade pattern to the linear symbol and the second to the 2D symbol.

    /// <summary>
    /// True when the HTML report covers a multi-mode (EAN/UPC + 2D) scan.
    /// When false, all Linear* fields below are null.
    /// </summary>
    public bool IsMultiMode { get; init; }

    /// <summary>
    /// Symbology name of the linear symbol: "EAN-13", "EAN-8", "UPC-A", or "UPC-E".
    /// Taken directly from the matching symbology cell in the HTML.
    /// </summary>
    public string? LinearSymbology { get; init; }

    /// <summary>
    /// Decoded payload of the linear symbol — EAN/UPC digit string (no check AI).
    /// Extracted from the all-digit cell immediately following the symbology cell.
    /// </summary>
    public string? LinearDecodedData { get; init; }

    /// <summary>
    /// Overall grade letter (A/B/C/D/F) for the linear symbol.
    /// Taken from the first "D.D (L)" pattern found in the cells (linear section
    /// precedes 2D section in the Webscan multi-mode HTML layout).
    /// When IsMultiMode is true the standard OverallGrade field is updated to
    /// reflect the second "D.D (L)" pattern (the 2D symbol).
    /// </summary>
    public string? LinearOverallGrade { get; init; }

    /// <summary>
    /// Decimal numeric grade for the linear symbol, e.g. 4.0 or 3.5.
    /// Parsed directly from the "D.D (L)" display string; preserves ISO 15416
    /// fractional precision (not rounded to integer letter equivalents).
    /// Null when the pattern was not found or could not be parsed.
    /// </summary>
    public decimal? LinearOverallGradeNumeric { get; init; }

    /// <summary>
    /// Formal grade string for the linear symbol, e.g. "A/06/660/Diffuse".
    /// Matches the Webscan TruCheck ISO 15416 format:
    ///   Letter/ApertureRef/Wavelength[/Lighting]
    /// </summary>
    public string? LinearFormalGrade { get; init; }

    /// <summary>Aperture reference number parsed from LinearFormalGrade, e.g. 6.</summary>
    public int? LinearAperture { get; init; }

    /// <summary>Wavelength (nm) parsed from LinearFormalGrade, e.g. 660.</summary>
    public int? LinearWavelength { get; init; }

    /// <summary>Lighting mode parsed from LinearFormalGrade, e.g. "Diffuse".</summary>
    public string? LinearLighting { get; init; }

    /// <summary>
    /// Grading standard for the linear symbol.
    /// Always "ISO/IEC 15416" when IsMultiMode is true.
    /// </summary>
    public string? LinearStandard { get; init; }

    // ── GS1 Data Format Check (scraped directly from DM TC HTML) ─────────────

    /// <summary>
    /// GS1 Data Format Check result scraped from the "Data Format Check" table
    /// in the DM TC HTML report.  The device has already validated the GS1
    /// Application Identifiers — this is the authoritative source.
    ///
    /// Null when the DM TC HTML does not contain a DFC table (non-GS1 symbol,
    /// linear-only scan, or older firmware that omits the section).
    ///
    /// DmstReportValidator.MergeAndValidate() prefers this over the computed
    /// BuildDataFormatCheck() result, which re-parses the push XML decoded-data
    /// string and is unreliable when BarcodeDataFormatter has transformed FNC1.
    /// </summary>
    public DataFormatCheckResult? ScrapedDataFormatCheck { get; init; }

    // ── HTML Verification Grades row (verbatim display strings) ──────────────
    //
    // Scraped directly from the Verification Grades table header row in the
    // TruCheck HTML.  No parsing, no reformatting — used verbatim in the PDF
    // BVG table so it stays in perfect sync with the TruCheck report.
    //
    // Source cells (column order confirmed from live DM475V scan 2026-08-18):
    //   Standard | Grade | Aperture | Wavelength | Lighting | Formal Grade
    //   ISO 15415:2024 | 4.0 (A) | 16 | 660 | 45Q | 4.0/16/660/45Q

    /// <summary>Verbatim standard string, e.g. "ISO 15415:2024".</summary>
    public string? HtmlStandard { get; init; }

    /// <summary>Verbatim overall grade display, e.g. "4.0 (A)".</summary>
    public string? HtmlOverallGradeDisplay { get; init; }

    /// <summary>Verbatim aperture string, e.g. "16".</summary>
    public string? HtmlAperture { get; init; }

    /// <summary>Verbatim wavelength string, e.g. "660".</summary>
    public string? HtmlWavelength { get; init; }

    /// <summary>Verbatim lighting string, e.g. "45Q".</summary>
    public string? HtmlLighting { get; init; }

    /// <summary>Verbatim formal grade string, e.g. "4.0/16/660/45Q".</summary>
    public string? HtmlFormalGrade { get; init; }

    // ── HTML header fields ────────────────────────────────────────────────────

    /// <summary>
    /// Raw "Verified: …" string scraped from the HTML report header,
    /// e.g. "Tue 18-Aug-2026 05:10:32(520ms) PM".
    /// Already in local Eastern time (device clock is DEVICE.TIMEZONE=America/New_York).
    /// Used verbatim as {{REPORT_DATETIME}} in the PDF — more faithful than the push
    /// XML VerificationDateTime, which may carry a UTC offset.
    /// </summary>
    public string? HtmlVerifiedString { get; init; }

    /// <summary>
    /// Filename-only (no directory) of the HTML source file, e.g.
    /// "2026-08-18_17-10-34-142_1787087819821.html".
    /// Captured before the file is deleted in DeleteAfterParse mode so the
    /// name is always available even after the transient file is gone.
    /// Maps to VerificationRecord.WebscanSourcePath (used for TruCheck Report Name in PDF).
    /// </summary>
    public string? HtmlSourceFileName { get; init; }

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
