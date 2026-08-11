namespace ExcelEngine.Models;

/// <summary>
/// Represents a single barcode verification event — one row in the VTCCP Excel log.
/// Holds fields for all symbology types (most will be null for any given record).
/// Maps directly to the 163-column TruCheckCompatible schema.
///
/// Field grouping follows the column order:
///   Block 1: Universal/Session
///   Block 2: 1D ISO 15416 parameters
///   Block 3: 2D Common parameters
///   Block 4: 2D Data Matrix standard parameters
///   Block 5: 2D Data Matrix quadrant-expanded parameters (32×32+)
///   Block 6: Military/Standards-specific
///   Block 7: Vendor/Part tracking
/// </summary>
public sealed record class VerificationRecord
{
    // ─── Block 1: Universal / Session ──────────────────────────────────────────

    public DateTime VerificationDateTime { get; init; } = DateTime.Now;
    public required string Symbology { get; init; }
    public SymbologyFamily SymbologyFamily { get; init; } = SymbologyFamily.Unknown;
    public string? DecodedData { get; init; }

    // Operator-supplied (from SessionState — NOT from device in DMV context)
    public string? OperatorId { get; init; }
    public string? JobName { get; init; }
    public int? RollNumber { get; init; }
    public string? BatchNumber { get; init; }
    public string? CompanyName { get; init; }
    public string? ProductName { get; init; }
    public string? CustomNote { get; init; }
    public string? User1 { get; init; }
    public string? User2 { get; init; }

    // Device-supplied identity + connection metadata
    public string? DeviceSerial { get; init; }
    public string? DeviceName { get; init; }
    /// <summary>
    /// DEVICE.TYPE string returned by the reader, e.g. "DM475V", "DM390", "DM395V".
    /// Populated at ConnectAsync; never null on a real device scan.
    /// </summary>
    public string? DeviceModel { get; init; }
    public string? FirmwareVersion { get; init; }
    /// <summary>
    /// "host:port" as configured in DeviceConfig, e.g. "10.10.10.7:44444".
    /// Captured once at session start; zero cost to store per record.
    /// </summary>
    public string? ConnectionAddress { get; init; }
    /// <summary>
    /// Resolved physical / logical medium: "GigE", "USB-Ethernet", or "USB-COM".
    /// Inferred from the IP address unless overridden in DeviceConfig.ConnectionMedium.
    /// DM475V = always "GigE".  DM390/DM395V Coglink = "USB-Ethernet".
    /// </summary>
    public string? ConnectionMedium { get; init; }

    // Sensor / imaging metadata — static per device model, captured at ConnectAsync
    /// <summary>Native sensor width in pixels, e.g. 2448. From per-model lookup table.</summary>
    public int?    SensorWidthPx      { get; init; }
    /// <summary>Native sensor height in pixels, e.g. 2048. From per-model lookup table.</summary>
    public int?    SensorHeightPx     { get; init; }
    /// <summary>Square pixel pitch in µm, e.g. 3.45. From per-model lookup table.</summary>
    public double? SensorPixelPitchUm { get; init; }
    /// <summary>
    /// Device IMAGE.SIZE setting at connect: "Full", "1/4", "1/16", "1/64".
    /// Controls IMAGE.SEND output; does NOT affect push XML JPEG crop dimensions.
    /// </summary>
    public string? ImageSizeSetting   { get; init; }

    public DateTime? CalibrationDate { get; init; }

    // Device scan properties (r.symbology.* — confirmed v1.24)
    /// <summary>AIM symbology identifier e.g. ]d1 (Data Matrix), ]Q2 (QR Code Model 2)</summary>
    public string? SymbologyId { get; init; }
    /// <summary>Decoder confidence score 0–100</summary>
    public int? SymbolQuality { get; init; }
    public decimal? SymbolAngle { get; init; }    // rotation in degrees
    public decimal? ModuleSizePx { get; init; }   // pixels per module (r.symbology.moduleSize)

    // Calibration status (rp.status3D.* — confirmed v1.24)
    public bool? FieldCalibrated { get; init; }
    public bool? FactoryCalibrated { get; init; }

    // Acceptance threshold (r.metrics.minPassGrade — confirmed v1.24; "NA" when none set)
    public string? MinPassGrade { get; init; }
    public decimal? MinPassRaw { get; init; }

    // Application-syntax check (q.overall.applicationStandard* — confirmed v1.24)
    public string? ApplicationStandard { get; init; }   // e.g. "GS1", "ISO 15434"
    /// <summary>Full device string e.g. "Pass" / "Fail (Quality)" / "Fail (X Dimension out of Range)"</summary>
    public string? ApplicationPass { get; init; }
    /// <summary>v1.25: reason suffix parsed from ApplicationPass, empty on pass</summary>
    public string? ApplicationPassReason { get; init; }

    // Optics discriminator (v1.25: ContrastUniformity == -1 AND MRD == -1 → LoadedImage)
    public string? OpticsSource { get; init; }    // "LiveScan" | "LoadedImage"

    // JPEG image payload (v1.25: r.trucheck.jpegImage base64 string)
    // Level 1 barcode crop — tight firmware crop around the decoded symbol.
    // Same image shown in the DMST verification panel (~200–600 px).
    public string? JpegImageBase64 { get; init; }

    // Level 2 ROI frame — operator-configured Region of Interest rectangle.
    // Retrieved via IMAGE.SEND immediately after each scan.  Wider than the
    // barcode crop; includes surrounding label area, HRI, lot numbers, etc.
    // Null when IMAGE.SEND failed or was not attempted (loaded-image replays
    // return the loaded image bytes rather than a live camera ROI).
    // Base64-encoded JPEG string, parallel structure to JpegImageBase64.
    public string? RoiJpegImageBase64 { get; init; }

    // OCR result — populated after DualEngineOcrRunner processes the ROI image
    // (or barcode crop for UPC/EAN where HRI is canonically part of the symbol).
    // Null when OCR has not been run or the image was not available.
    // OcrResult type lives in OcrEngine project; stored as object to avoid
    // a hard compile-time dependency from ExcelEngine → OcrEngine.
    // Command Pilot wires the concrete type; serialisation uses OcrResultDto.
    public OcrResultDto? OcrResult { get; init; }

    // ── Data provenance ────────────────────────────────────────────────────────

    /// <summary>
    /// Semicolon-separated list of fields whose values were NOT sourced from the
    /// push XML (the canonical per-scan output of the Format Data push script).
    ///
    /// Format: "FieldName:Source" pairs.
    /// Example: "ECLevel:HtmlReport;DataMaskPattern:HtmlReport;ECI:HtmlReport"
    /// Example: "ECLevel:SymbolResultFull;ECI:SymbolResultFull"
    ///
    /// Null (omitted) if every populated field came from the push XML.
    ///
    /// Populated fields on fw 6.1.16_sr4 — four fields confirmed permanently
    /// unresolvable from push XML (v1.33 probe campaign, 2026-05-25):
    ///   ECLevel, DataMaskPattern, ECI, ImagePolarity
    /// These will set DataSourceExceptions if sourced from a secondary channel
    /// (DMCC RESULT-FORMAT FULL or DMST HTML report scrape).
    /// </summary>
    public string? DataSourceExceptions { get; set; }

    /// <summary>
    /// Semicolon-separated list of fields where the push XML value differs from
    /// the DMST HTML report value for the same scan.
    ///
    /// Format: "FieldName:Push={pushVal},Html={htmlVal}" pairs.
    /// Example: "OverallGrade:Push=D,Html=D;ContrastUniformity:Push=75,Html=75"
    ///          — in this case all match, so field would be null.
    /// Example: "NominalXDim:Push=20.3 mil,Html=20.3 mil" would be null (match).
    /// A non-null value means at least one field disagreed.
    ///
    /// Purpose: cross-validation sanity check. Every field producible from both
    /// sources is compared; mismatches surface parser errors or firmware anomalies.
    /// Discrepancy patterns that recur across scans are candidates for Cognex bug reports.
    ///
    /// Null if no HTML report was scraped for this scan, or if all compared fields matched.
    /// </summary>
    public string? ValidationDiscrepancies { get; set; }

    // Overall grade outcome
    /// <summary>e.g. "4.0/16/660/45Q" or "4.0/06/660"</summary>
    public string? FormalGrade { get; init; }
    public GradingResult? OverallGrade { get; init; }
    public OverallPassFail CustomPassFail { get; init; } = OverallPassFail.NotApplicable;

    // Verification settings
    public int? Aperture { get; init; }
    public int? Wavelength { get; init; }
    public string? Lighting { get; init; }
    public string? Standard { get; init; }

    // ─── Block 2: 1D ISO 15416 Parameters ─────────────────────────────────────

    /// <summary>Symbol ANSI Grade (overall for 1D)</summary>
    public GradingResult? SymbolAnsiGrade { get; init; }

    /// <summary>Start/Stop Grade (Code 39, Code 128, etc.)</summary>
    public GradingResult? StartStopGrade { get; init; }

    /// <summary>Start/Stop SRP Grade</summary>
    public GradingResult? StartStopSrpGrade { get; init; }

    /// <summary>Up to 10 individual scan results per verification</summary>
    public IReadOnlyList<ScanResult1D> ScanResults { get; init; } = [];

    // 1D Summary averages
    public decimal? Avg_Edge { get; init; }
    public string? Avg_RlRd { get; init; }
    public decimal? Avg_SC { get; init; }
    public decimal? Avg_MinEC { get; init; }
    public decimal? Avg_MOD { get; init; }
    public decimal? Avg_Defect { get; init; }
    public string? Avg_DCOD { get; init; }
    public decimal? Avg_DEC { get; init; }
    public decimal? Avg_LQZ { get; init; }   // Average Left Quiet Zone measurement
    public decimal? Avg_RQZ { get; init; }   // Average Right Quiet Zone measurement
    public decimal? Avg_HQZ { get; init; }   // Average High (top/header) Quiet Zone (some symbologies)
    public decimal? Avg_MinQZ { get; init; } // Min(LQZ, RQZ [, HQZ]) — derived summary

    // 1D General Characteristics
    public decimal? BWG_Percent { get; init; }
    public decimal? BWG_Mil { get; init; }
    public decimal? Magnification { get; init; }
    public decimal? NominalXDim_1D { get; init; }
    public decimal? InspectionZoneHeight { get; init; }
    public decimal? DecodableSymbolHeight { get; init; }
    public decimal? Ratio { get; init; }  // Code 39 only

    /// <summary>Element widths data — written to a separate "Element Widths" worksheet</summary>
    public ElementWidthData? ElementWidths { get; init; }

    // ─── Block 3: 2D Common Parameters ────────────────────────────────────────

    /// <summary>UEC% e.g. 100%, 42%</summary>
    public decimal? UEC_Percent { get; init; }
    public GradingResult? UEC_Grade { get; init; }

    /// <summary>Symbol Contrast % e.g. 84%</summary>
    public decimal? SC_Percent { get; init; }
    public string? SC_RlRd { get; init; }  // e.g. "89/4"
    public GradingResult? SC_Grade { get; init; }

    public GradingResult? MOD_Grade { get; init; }

    /// <summary>Reflectance Margin — parameter 3b, added in newer firmware/standard</summary>
    public GradingResult? RM_Grade { get; init; }

    /// <summary>Axial Non-uniformity % e.g. 0.2%</summary>
    public decimal? ANU_Percent { get; init; }
    public GradingResult? ANU_Grade { get; init; }

    /// <summary>Grid Non-uniformity % e.g. 2.3%</summary>
    public decimal? GNU_Percent { get; init; }
    public GradingResult? GNU_Grade { get; init; }

    public GradingResult? FPD_Grade { get; init; }
    public decimal? FPD_Value { get; init; }        // q.fixedPatternDamage.raw
    public GradingResult? DECODE_Grade { get; init; }

    /// <summary>Average Grade (AG) — parameter 17, ISO 15415</summary>
    public decimal? AG_Value { get; init; }
    public GradingResult? AG_Grade { get; init; }

    /// <summary>Distributed Damage grade (q.distributedDamageGrade)</summary>
    public GradingResult? DD_Grade { get; init; }

    /// <summary>ISO 15415 average grade across all parameters (q.averageGrade) — distinct from AG/Print Growth</summary>
    public GradingResult? AverageGrade { get; init; }
    public decimal? AverageGradeNumeric { get; init; }

    // 2D General Characteristics (shared across DM, GS1-DM, QR)
    public string? MatrixSize { get; init; }       // e.g. "22x22 (Data: 20x20)"
    public decimal? HorizontalBWG { get; init; }   // % e.g. -11
    public decimal? VerticalBWG { get; init; }
    public int? EncodedCharacters { get; init; }
    public int? TotalCodewords { get; init; }
    public int? DataCodewords { get; init; }
    public int? ErrorCorrectionBudget { get; init; }
    public int? ErrorsCorrected { get; init; }
    public int? ErrorCapacityUsed { get; init; }
    public string? ErrorCorrectionType { get; init; }   // "ECC 200"
    public ImagePolarity ImagePolarity { get; init; } = ImagePolarity.Unknown;
    public decimal? NominalXDim_2D { get; init; }
    public decimal? PixelsPerModule { get; init; }
    public string? ContrastUniformity { get; init; }    // e.g. "72 at module(10,5)"
    public string? MRD { get; init; }                   // e.g. "71% (77% - 6%)"
    public string? ContrastUniformityRow { get; init; } // row index of worst module
    public string? ContrastUniformityCol { get; init; } // col index of worst module
    public decimal? MinReflectance { get; init; }       // q.minimumReflectance.raw (suppressed when F+0)

    // ─── Block 4: 2D Data Matrix Standard Parameters (≤26×26) ─────────────────

    /// <summary>Left 'L' Side</summary>
    public GradingResult? LLS_Grade { get; init; }

    /// <summary>Bottom 'L' Side</summary>
    public GradingResult? BLS_Grade { get; init; }

    /// <summary>Left Quiet Zone</summary>
    public GradingResult? LQZ_Grade { get; init; }

    /// <summary>Bottom Quiet Zone</summary>
    public GradingResult? BQZ_Grade { get; init; }

    /// <summary>Top Quiet Zone (standard, single-region)</summary>
    public GradingResult? TQZ_Grade { get; init; }

    /// <summary>Right Quiet Zone (standard, single-region)</summary>
    public GradingResult? RQZ_Grade { get; init; }

    /// <summary>Top Transition Ratio % (standard)</summary>
    public decimal? TTR_Percent { get; init; }
    public GradingResult? TTR_Grade { get; init; }

    /// <summary>Right Transition Ratio % (standard)</summary>
    public decimal? RTR_Percent { get; init; }
    public GradingResult? RTR_Grade { get; init; }

    /// <summary>Top Clock Track (standard)</summary>
    public GradingResult? TCT_Grade { get; init; }

    /// <summary>Right Clock Track (standard)</summary>
    public GradingResult? RCT_Grade { get; init; }

    // ─── Block 5: 2D Data Matrix Quadrant-Expanded Parameters (≥32×32) ─────────
    // Parameters 11–16 each split into 2 or 4 quadrant sub-parameters.
    // These are populated only when MatrixRows >= 32.

    // Quiet Zones (4 quadrant subdivisions)
    public GradingResult? ULQZ_Grade { get; init; }  // Upper Left Quiet Zone
    public GradingResult? URQZ_Grade { get; init; }  // Upper Right Quiet Zone
    public GradingResult? RUQZ_Grade { get; init; }  // Right Upper Quiet Zone
    public GradingResult? RLQZ_Grade { get; init; }  // Right Lower Quiet Zone

    // Top Transition Ratios by quadrant
    public decimal? ULQTTR_Percent { get; init; }
    public GradingResult? ULQTTR_Grade { get; init; }  // Upper Left Quadrant TTR
    public decimal? URQTTR_Percent { get; init; }
    public GradingResult? URQTTR_Grade { get; init; }  // Upper Right Quadrant TTR
    public decimal? LLQTTR_Percent { get; init; }
    public GradingResult? LLQTTR_Grade { get; init; }  // Lower Left Quadrant TTR
    public decimal? LRQTTR_Percent { get; init; }
    public GradingResult? LRQTTR_Grade { get; init; }  // Lower Right Quadrant TTR

    // Right Transition Ratios by quadrant
    public decimal? ULQRTR_Percent { get; init; }
    public GradingResult? ULQRTR_Grade { get; init; }
    public decimal? URQRTR_Percent { get; init; }
    public GradingResult? URQRTR_Grade { get; init; }
    public decimal? LLQRTR_Percent { get; init; }
    public GradingResult? LLQRTR_Grade { get; init; }
    public decimal? LRQRTR_Percent { get; init; }
    public GradingResult? LRQRTR_Grade { get; init; }

    // Top Clock Tracks by quadrant
    public GradingResult? ULQTCT_Grade { get; init; }
    public GradingResult? URQTCT_Grade { get; init; }
    public GradingResult? LLQTCT_Grade { get; init; }
    public GradingResult? LRQTCT_Grade { get; init; }

    // Right Clock Tracks by quadrant
    public GradingResult? ULQRCT_Grade { get; init; }
    public GradingResult? URQRCT_Grade { get; init; }
    public GradingResult? LLQRCT_Grade { get; init; }
    public GradingResult? LRQRCT_Grade { get; init; }

    // ─── Block 6: QR Code Parameters ──────────────────────────────────────────
    // Populated when SymbologyFamily is QRCode or GS1QRCode.
    // Symbol characteristics (from r.trucheck.symbols[0] — paths confirmed by v1.26 probe):
    public string? QR_Version { get; init; }        // e.g. "V3 (29×29)"
    public string? QR_ECLevel { get; init; }        // L / M / Q / H
    public string? QR_MaskPattern { get; init; }    // 0–7

    /// <summary>
    /// ECI (Extended Channel Interpretation) assignment value, e.g. "000003" (ISO 8859-1 / Latin-1).
    /// Present when the AIM modifier indicates ECI (]Q1 — modifier bit-0 = ECI assigning mode).
    /// PERMANENTLY UNRESOLVABLE from push XML on fw 6.1.16_sr4 (v1.33 probe campaign).
    /// Sourced from DMST HTML report scrape (or DMCC RESULT-FORMAT FULL if probe succeeds).
    /// Flagged in DataSourceExceptions when populated.
    /// </summary>
    public string? QR_ECI { get; init; }

    // ISO 15415 QR-specific grade parameters (8 parameters; paths from v1.26 DebugSymbols0 probe):
    public GradingResult? QR_ULP_Grade { get; init; }  // Upper-Left Finder Pattern
    public GradingResult? QR_URP_Grade { get; init; }  // Upper-Right Finder Pattern
    public GradingResult? QR_LLP_Grade { get; init; }  // Lower-Left Finder Pattern
    public GradingResult? QR_HCT_Grade { get; init; }  // Horizontal Clock Track
    public GradingResult? QR_VCT_Grade { get; init; }  // Vertical Clock Track
    public GradingResult? QR_ALP_Grade { get; init; }  // Alignment Pattern
    public GradingResult? QR_VIB_Grade { get; init; }  // Version Information Blocks
    public GradingResult? QR_FIB_Grade { get; init; }  // Format Information Blocks

    // ─── Block 7: Military / Standards-Specific ────────────────────────────────

    public string? UIDFormat { get; init; }           // MIL-STD-130 UID format
    public string? MilStd130VersionLetter { get; init; }
    public GradingResult? AS9132_Grade { get; init; }
    public string? AG_DDG { get; init; }              // AG/DDG composite field
    public string? SC_CC { get; init; }               // SC/CC composite field
    public string? MOD_CMOD { get; init; }            // MOD/CMOD composite field

    // Rmax/Rmin calibration values
    public decimal? Rmax { get; init; }
    public decimal? TargetRmax { get; init; }
    public decimal? RmaxDeviation { get; init; }
    public decimal? Rmin { get; init; }
    public decimal? TargetRmin { get; init; }
    public decimal? RminDeviation { get; init; }

    // ─── Block 8: Vendor / Part Tracking ──────────────────────────────────────

    public string? VendorName { get; init; }
    public string? PartNumber { get; init; }
    public string? SerialNumber { get; init; }

    // ─── GS1 / Data Format Check ───────────────────────────────────────────────

    public DataFormatCheckResult? DataFormatCheck { get; init; }

    // ─── Block B7: Modulation Values + Codeword Values ─────────────────────────

    /// <summary>
    /// Raw modulation grid for the "Modulation Values" worksheet.
    /// Populated from q.modulationArray by the push parser.
    /// Null for 1D records and for scans where the array was absent in the push payload.
    /// Grid dimensions: (SymbolRows+2) × (SymbolCols+2) — includes the 1-module QZ border.
    /// </summary>
    public ModulationValuesData? ModulationValues { get; init; }

    /// <summary>
    /// Codeword and encodation data for the "Codeword Values" worksheet.
    /// Populated from q.codewordArray + q.encodationAnalysisArray by the push parser.
    /// codewordArray length = total codewords (data + ECC combined) — confirmed v1.28.
    /// Null for 1D records and for scans where the arrays were absent.
    /// </summary>
    public CodewordValuesData? CodewordValues { get; init; }

    // ─── RFID Cross-Validation ────────────────────────────────────────────────
    // Populated after the RFID scan window closes. Null on all fields when RFID
    // is not configured (RfidComPort empty) or when the scan is skipped.

    /// <summary>Raw EPC hex string from the selected tag, e.g. "30342A7CC844C7D0F36A0676".</summary>
    public string? RfidEpcHex { get; init; }

    /// <summary>GTIN-14 decoded from the EPC (14 digits). Null if decode failed or no tag.</summary>
    public string? RfidGtin14 { get; init; }

    /// <summary>Serial number decoded from the EPC. Null if not present or decode failed.</summary>
    public string? RfidSerial { get; init; }

    /// <summary>
    /// Cross-validation outcome: "Pass", "Fail", "NoTag", "ParseError",
    /// "MultipleTagsDetected", or "Skipped".
    /// </summary>
    public string? RfidStatus { get; init; }

    /// <summary>
    /// Semicolon-separated mismatch description, e.g. "GTIN14:RFID=…,BC=…;Serial:RFID=…,BC=…".
    /// Null on Pass or when no comparison was possible.
    /// </summary>
    public string? RfidMismatchDetail { get; init; }

    /// <summary>Duration of the RFID scan window in milliseconds.</summary>
    public int? RfidScanWindowMs { get; init; }

    /// <summary>
    /// True = GCP registered in GS1 table; false = not found; null = GCP check not run.
    /// </summary>
    public bool? RfidGcpValid { get; init; }

    // ─── Helper Properties ────────────────────────────────────────────────────

    /// <summary>
    /// True if this record requires the quadrant-expanded parameter set (32×32 or larger matrix).
    /// </summary>
    public bool IsLargeMatrix
    {
        get
        {
            if (string.IsNullOrEmpty(MatrixSize)) return false;
            var parts = MatrixSize.Split('x', StringSplitOptions.TrimEntries);
            if (parts.Length >= 1 && int.TryParse(parts[0], out var rows))
                return rows >= 32;
            return false;
        }
    }

    public bool Is1D => SymbologyFamily == SymbologyFamily.Linear1D;
    public bool Is2D => !Is1D && SymbologyFamily != SymbologyFamily.Unknown;
}
