namespace ExcelEngine.Writer;

using ExcelEngine.Models;
using ExcelEngine.Schema;

/// <summary>
/// Maps a 2D Data Matrix (or GS1-DM) VerificationRecord to a flat dictionary of
/// { ColumnDefinition.FieldId → cell value (string | double | DateTime | null) }.
/// Null means "leave the cell blank" (non-applicable for this symbology or record).
///
/// Called by ExcelWriter.AppendRecord — format-agnostic.
/// </summary>
public static class DataMatrix2DMapper
{
    public static Dictionary<string, object?> Map(VerificationRecord r, ColumnSchema schema)
    {
        var d = new Dictionary<string, object?>(schema.Columns.Count, StringComparer.Ordinal);

        // ── Block A: Universal / Session ──────────────────────────────────────
        d["Date"]           = r.VerificationDateTime.Date;
        d["Time"]           = r.VerificationDateTime;
        d["OperatorId"]     = r.OperatorId;
        d["RollNumber"]     = r.RollNumber.HasValue ? (object)r.RollNumber.Value : null;
        d["JobName"]        = r.JobName;
        d["BatchNumber"]    = r.BatchNumber;
        d["CompanyName"]    = r.CompanyName;
        d["ProductName"]    = r.ProductName;
        d["Symbology"]      = r.Symbology;
        d["DecodedData"]    = r.DecodedData;
        d["FormalGrade"]    = r.FormalGrade;
        d["OverallLetter"]  = r.OverallGrade?.LetterGradeString;
        d["OverallNumeric"] = r.OverallGrade?.NumericGrade.HasValue == true
                               ? (object)(double)r.OverallGrade.NumericGrade!.Value : null;
        d["CustomPassFail"] = r.CustomPassFail switch
        {
            OverallPassFail.Pass => "Pass",
            OverallPassFail.Fail => "Fail",
            _ => null,
        };
        d["User1"]          = r.User1;
        d["User2"]          = r.User2;
        d["DeviceSerial"]   = r.DeviceSerial;
        d["DeviceName"]     = r.DeviceName;
        d["FirmwareVersion"]= r.FirmwareVersion;
        d["CalibrationDate"]= r.CalibrationDate.HasValue ? (object)r.CalibrationDate.Value : null;

        // ── Block B: 1D fields — all null for 2D records ──────────────────────
        d["SymbolAnsiGrade_Numeric"] = null;
        d["StartStopGrade_Numeric"]  = null;
        d["StartStopSrpGrade_Numeric"]=null;
        d["Avg_Edge"]  = null;
        d["Avg_RlRd"]  = null;
        d["Avg_SC"]    = null;
        d["Avg_MinEC"] = null;
        d["Avg_MOD"]   = null;
        d["Avg_Defect"]= null;
        d["Avg_DCOD"]  = null;
        d["Avg_DEC"]   = null;
        d["Avg_MinQZ"] = null;
        d["BWG_Percent"]= null;
        d["BWG_Mil"]   = null;
        d["Magnification"]= null;
        d["NominalXDim_1D"]= null;
        d["InspectionZoneHeight"]= null;
        d["DecodableSymbolHeight"]= null;
        d["Ratio"]     = null;

        // ── Block C: 2D Common Parameters ────────────────────────────────────
        d["UEC_Percent"]       = ToDouble(r.UEC_Percent);
        d["UEC_Grade_Numeric"] = ToDouble(r.UEC_Grade?.NumericGrade);
        d["SC_Percent"]        = ToDouble(r.SC_Percent);
        d["SC_RlRd"]           = r.SC_RlRd;
        d["SC_Grade_Numeric"]  = ToDouble(r.SC_Grade?.NumericGrade);
        d["MOD_Grade_Numeric"] = ToDouble(r.MOD_Grade?.NumericGrade);
        d["RM_Grade_Numeric"]  = ToDouble(r.RM_Grade?.NumericGrade);
        d["ANU_Percent"]       = ToDouble(r.ANU_Percent);
        d["ANU_Grade_Numeric"] = ToDouble(r.ANU_Grade?.NumericGrade);
        d["GNU_Percent"]       = ToDouble(r.GNU_Percent);
        d["GNU_Grade_Numeric"] = ToDouble(r.GNU_Grade?.NumericGrade);
        d["FPD_Grade_Numeric"] = ToDouble(r.FPD_Grade?.NumericGrade);
        d["DECODE_Grade_Numeric"]= ToDouble(r.DECODE_Grade?.NumericGrade);
        d["AG_Value"]          = ToDouble(r.AG_Value);
        d["AG_Grade_Numeric"]  = ToDouble(r.AG_Grade?.NumericGrade);
        d["AG_Grade_Letter"]   = r.AG_Grade?.LetterGradeString;
        d["OverallPassFail2D"] = r.OverallGrade?.PassFailString;
        d["MatrixSize"]        = r.MatrixSize;
        d["HorizontalBWG"]     = ToDouble(r.HorizontalBWG);
        d["VerticalBWG"]       = ToDouble(r.VerticalBWG);
        d["EncodedCharacters"] = r.EncodedCharacters.HasValue ? (object)(double)r.EncodedCharacters.Value : null;
        d["TotalCodewords"]    = r.TotalCodewords.HasValue ? (object)(double)r.TotalCodewords.Value : null;
        d["DataCodewords"]     = r.DataCodewords.HasValue ? (object)(double)r.DataCodewords.Value : null;
        d["ErrorCorrectionBudget"] = r.ErrorCorrectionBudget.HasValue ? (object)(double)r.ErrorCorrectionBudget.Value : null;
        d["ErrorsCorrected"]   = r.ErrorsCorrected.HasValue ? (object)(double)r.ErrorsCorrected.Value : null;
        d["ErrorCapacityUsed"] = r.ErrorCapacityUsed.HasValue ? (object)(double)r.ErrorCapacityUsed.Value : null;
        d["ErrorCorrectionType"] = r.ErrorCorrectionType;
        d["ImagePolarity"]     = r.ImagePolarity switch
        {
            ImagePolarity.BlackOnWhite => "Black on white",
            ImagePolarity.WhiteOnBlack => "White on black",
            _ => null,
        };
        d["NominalXDim_2D"]    = ToDouble(r.NominalXDim_2D);
        d["PixelsPerModule"]   = ToDouble(r.PixelsPerModule);
        d["ContrastUniformity"]= r.ContrastUniformity;
        d["MRD"]               = r.MRD;

        // ── Block D: 2D Data Matrix Standard Parameters ───────────────────────
        d["LLS_Grade_Numeric"] = ToDouble(r.LLS_Grade?.NumericGrade);
        d["LLS_Grade_Letter"]  = r.LLS_Grade?.LetterGradeString;
        d["BLS_Grade_Numeric"] = ToDouble(r.BLS_Grade?.NumericGrade);
        d["BLS_Grade_Letter"]  = r.BLS_Grade?.LetterGradeString;
        d["LQZ_Grade_Numeric"] = ToDouble(r.LQZ_Grade?.NumericGrade);
        d["LQZ_Grade_Letter"]  = r.LQZ_Grade?.LetterGradeString;
        d["BQZ_Grade_Numeric"] = ToDouble(r.BQZ_Grade?.NumericGrade);
        d["BQZ_Grade_Letter"]  = r.BQZ_Grade?.LetterGradeString;
        d["TQZ_Grade_Numeric"] = ToDouble(r.TQZ_Grade?.NumericGrade);
        d["TQZ_Grade_Letter"]  = r.TQZ_Grade?.LetterGradeString;
        d["RQZ_Grade_Numeric"] = ToDouble(r.RQZ_Grade?.NumericGrade);
        d["RQZ_Grade_Letter"]  = r.RQZ_Grade?.LetterGradeString;
        d["TTR_Percent"]       = ToDouble(r.TTR_Percent);
        d["TTR_Grade_Numeric"] = ToDouble(r.TTR_Grade?.NumericGrade);
        d["RTR_Percent"]       = ToDouble(r.RTR_Percent);
        d["RTR_Grade_Numeric"] = ToDouble(r.RTR_Grade?.NumericGrade);
        d["TCT_Grade_Numeric"] = ToDouble(r.TCT_Grade?.NumericGrade);
        d["TCT_Grade_Letter"]  = r.TCT_Grade?.LetterGradeString;
        d["RCT_Grade_Numeric"] = ToDouble(r.RCT_Grade?.NumericGrade);
        d["RCT_Grade_Letter"]  = r.RCT_Grade?.LetterGradeString;

        // ── Block E: Quadrant-Expanded (only when IsLargeMatrix) ──────────────
        bool quad = r.IsLargeMatrix;
        d["ULQZ_Grade_Numeric"] = quad ? ToDouble(r.ULQZ_Grade?.NumericGrade) : null;
        d["URQZ_Grade_Numeric"] = quad ? ToDouble(r.URQZ_Grade?.NumericGrade) : null;
        d["RUQZ_Grade_Numeric"] = quad ? ToDouble(r.RUQZ_Grade?.NumericGrade) : null;
        d["RLQZ_Grade_Numeric"] = quad ? ToDouble(r.RLQZ_Grade?.NumericGrade) : null;

        d["ULQTTR_Percent"]      = quad ? ToDouble(r.ULQTTR_Percent) : null;
        d["ULQTTR_Grade_Numeric"]= quad ? ToDouble(r.ULQTTR_Grade?.NumericGrade) : null;
        d["URQTTR_Percent"]      = quad ? ToDouble(r.URQTTR_Percent) : null;
        d["URQTTR_Grade_Numeric"]= quad ? ToDouble(r.URQTTR_Grade?.NumericGrade) : null;
        d["LLQTTR_Percent"]      = quad ? ToDouble(r.LLQTTR_Percent) : null;
        d["LLQTTR_Grade_Numeric"]= quad ? ToDouble(r.LLQTTR_Grade?.NumericGrade) : null;
        d["LRQTTR_Percent"]      = quad ? ToDouble(r.LRQTTR_Percent) : null;
        d["LRQTTR_Grade_Numeric"]= quad ? ToDouble(r.LRQTTR_Grade?.NumericGrade) : null;

        d["ULQRTR_Percent"]      = quad ? ToDouble(r.ULQRTR_Percent) : null;
        d["ULQRTR_Grade_Numeric"]= quad ? ToDouble(r.ULQRTR_Grade?.NumericGrade) : null;
        d["URQRTR_Percent"]      = quad ? ToDouble(r.URQRTR_Percent) : null;
        d["URQRTR_Grade_Numeric"]= quad ? ToDouble(r.URQRTR_Grade?.NumericGrade) : null;
        d["LLQRTR_Percent"]      = quad ? ToDouble(r.LLQRTR_Percent) : null;
        d["LLQRTR_Grade_Numeric"]= quad ? ToDouble(r.LLQRTR_Grade?.NumericGrade) : null;
        d["LRQRTR_Percent"]      = quad ? ToDouble(r.LRQRTR_Percent) : null;
        d["LRQRTR_Grade_Numeric"]= quad ? ToDouble(r.LRQRTR_Grade?.NumericGrade) : null;

        d["ULQTCT_Grade_Numeric"]= quad ? ToDouble(r.ULQTCT_Grade?.NumericGrade) : null;
        d["URQTCT_Grade_Numeric"]= quad ? ToDouble(r.URQTCT_Grade?.NumericGrade) : null;
        d["LLQTCT_Grade_Numeric"]= quad ? ToDouble(r.LLQTCT_Grade?.NumericGrade) : null;
        d["LRQTCT_Grade_Numeric"]= quad ? ToDouble(r.LRQTCT_Grade?.NumericGrade) : null;

        d["ULQRCT_Grade_Numeric"]= quad ? ToDouble(r.ULQRCT_Grade?.NumericGrade) : null;
        d["URQRCT_Grade_Numeric"]= quad ? ToDouble(r.URQRCT_Grade?.NumericGrade) : null;
        d["LLQRCT_Grade_Numeric"]= quad ? ToDouble(r.LLQRCT_Grade?.NumericGrade) : null;
        d["LRQRCT_Grade_Numeric"]= quad ? ToDouble(r.LRQRCT_Grade?.NumericGrade) : null;

        // ── Block F: Military / Standards (blank for standard DM records) ─────
        d["UIDFormat"]                = r.UIDFormat;
        d["MilStd130VersionLetter"]   = r.MilStd130VersionLetter;
        d["AS9132_Grade_Numeric"]     = ToDouble(r.AS9132_Grade?.NumericGrade);
        d["Rmax"]                     = ToDouble(r.Rmax);
        d["TargetRmax"]               = ToDouble(r.TargetRmax);
        d["RmaxDeviation"]            = ToDouble(r.RmaxDeviation);
        d["Rmin"]                     = ToDouble(r.Rmin);
        d["TargetRmin"]               = ToDouble(r.TargetRmin);
        d["RminDeviation"]            = ToDouble(r.RminDeviation);

        // ── Block G: Vendor / Part Tracking ──────────────────────────────────
        d["VendorName"]   = r.VendorName;
        d["PartNumber"]   = r.PartNumber;
        d["SerialNumber"] = r.SerialNumber;

        // ── Block H: QR reserved — null for DM ───────────────────────────────
        d["QR_Version"]    = null;
        d["QR_ECLevel"]    = null;
        d["QR_MaskPattern"]= null;

        // ── Custom Note ───────────────────────────────────────────────────────
        d["CustomNote"] = r.CustomNote;

        // ── Block I: DFC columns — null values for non-DFC records ────────────
        // ExcelWriter.WriteDfcColumns() fills these for records with DataFormatCheck set.
        // Explicit nulls here ensure WriteDataRow doesn't error on unknown field IDs
        // for records without DFC data.
        d["DFC_Standard"] = null;
        for (int slot = 1; slot <= 8; slot++)
        {
            d[$"DFC_R{slot}_Name"]  = null;
            d[$"DFC_R{slot}_Data"]  = null;
            d[$"DFC_R{slot}_Check"] = null;
        }

        return d;
    }

    private static object? ToDouble(decimal? value)
        => value.HasValue ? (object)(double)value.Value : null;
}
