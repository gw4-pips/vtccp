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

    // Device-supplied
    public string? DeviceSerial { get; init; }
    public string? DeviceName { get; init; }
    public string? FirmwareVersion { get; init; }
    public DateTime? CalibrationDate { get; init; }

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
    public GradingResult? DECODE_Grade { get; init; }

    /// <summary>Average Grade (AG) — parameter 17, ISO 15415</summary>
    public decimal? AG_Value { get; init; }
    public GradingResult? AG_Grade { get; init; }

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
    public string? ContrastUniformity { get; init; }   // e.g. "72 at module(10,5)"
    public string? MRD { get; init; }                  // e.g. "71% (77% - 6%)"

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

    // ─── Block 6: QR Code Parameters (stubs — data writing in later task) ──────
    // QR Code uses ISO 15415 + QR-specific parameters; column positions reserved.
    // Populated when SymbologyFamily is QRCode or GS1QRCode.

    public string? QR_Version { get; init; }        // e.g. "V3 (29×29)"
    public string? QR_ECLevel { get; init; }        // L / M / Q / H
    public string? QR_MaskPattern { get; init; }

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
