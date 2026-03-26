namespace ExcelEngine.Writer;

using ExcelEngine.Models;
using ExcelEngine.Schema;

/// <summary>
/// Maps a 1D ISO 15416 VerificationRecord to a flat dictionary of
/// { ColumnDefinition.FieldId → cell value (string | double | DateTime | null) }.
/// Null means "leave the cell blank" (non-applicable for this symbology or record).
///
/// Called by ExcelWriter.AppendRecord for Linear1D records.
/// The per-scan sub-table and Element Widths sheet are written by separate helpers.
/// </summary>
public static class ISO15416Mapper
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
        d["User1"]           = r.User1;
        d["User2"]           = r.User2;
        d["DeviceSerial"]    = r.DeviceSerial;
        d["DeviceName"]      = r.DeviceName;
        d["FirmwareVersion"] = r.FirmwareVersion;
        d["CalibrationDate"] = r.CalibrationDate.HasValue ? (object)r.CalibrationDate.Value : null;

        // ── Block B: 1D ISO 15416 Summary Parameters ──────────────────────────
        d["SymbolAnsiGrade_Numeric"]  = ToDouble(r.SymbolAnsiGrade?.NumericGrade);
        d["StartStopGrade_Numeric"]   = ToDouble(r.StartStopGrade?.NumericGrade);
        d["StartStopSrpGrade_Numeric"]= ToDouble(r.StartStopSrpGrade?.NumericGrade);

        // Average parameter values
        d["Avg_Edge"]   = ToDouble(r.Avg_Edge);
        d["Avg_RlRd"]   = r.Avg_RlRd;
        d["Avg_SC"]     = ToDouble(r.Avg_SC);
        d["Avg_MinEC"]  = ToDouble(r.Avg_MinEC);
        d["Avg_MOD"]    = ToDouble(r.Avg_MOD);
        d["Avg_Defect"] = ToDouble(r.Avg_Defect);
        d["Avg_DCOD"]   = r.Avg_DCOD;
        d["Avg_DEC"]    = ToDouble(r.Avg_DEC);
        d["Avg_MinQZ"]  = ToDouble(r.Avg_MinQZ);

        // General Characteristics
        d["BWG_Percent"]           = ToDouble(r.BWG_Percent);
        d["BWG_Mil"]               = ToDouble(r.BWG_Mil);
        d["Magnification"]         = ToDouble(r.Magnification);
        d["NominalXDim_1D"]        = ToDouble(r.NominalXDim_1D);
        d["InspectionZoneHeight"]  = ToDouble(r.InspectionZoneHeight);
        d["DecodableSymbolHeight"] = ToDouble(r.DecodableSymbolHeight);
        d["Ratio"]                 = ToDouble(r.Ratio);   // Code 39 only; null for UPC/EAN

        // ── Block C: 2D Common — all null for 1D records ──────────────────────
        d["UEC_Percent"]        = null;
        d["UEC_Grade_Numeric"]  = null;
        d["SC_Percent"]         = null;
        d["SC_RlRd"]            = null;
        d["SC_Grade_Numeric"]   = null;
        d["MOD_Grade_Numeric"]  = null;
        d["RM_Grade_Numeric"]   = null;
        d["ANU_Percent"]        = null;
        d["ANU_Grade_Numeric"]  = null;
        d["GNU_Percent"]        = null;
        d["GNU_Grade_Numeric"]  = null;
        d["FPD_Grade_Numeric"]  = null;
        d["DECODE_Grade_Numeric"]= null;
        d["AG_Value"]           = null;
        d["AG_Grade_Numeric"]   = null;
        d["AG_Grade_Letter"]    = null;
        d["OverallPassFail2D"]  = null;
        d["MatrixSize"]         = null;
        d["HorizontalBWG"]      = null;
        d["VerticalBWG"]        = null;
        d["EncodedCharacters"]  = null;
        d["TotalCodewords"]     = null;
        d["DataCodewords"]      = null;
        d["ErrorCorrectionBudget"] = null;
        d["ErrorsCorrected"]    = null;
        d["ErrorCapacityUsed"]  = null;
        d["ErrorCorrectionType"]= null;
        d["ImagePolarity"]      = null;
        d["NominalXDim_2D"]     = null;
        d["PixelsPerModule"]    = null;
        d["ContrastUniformity"] = null;
        d["MRD"]                = null;

        // ── Block D: 2D Data Matrix Standard — all null for 1D ───────────────
        d["LLS_Grade_Numeric"] = null;
        d["LLS_Grade_Letter"]  = null;
        d["BLS_Grade_Numeric"] = null;
        d["BLS_Grade_Letter"]  = null;
        d["LQZ_Grade_Numeric"] = null;
        d["LQZ_Grade_Letter"]  = null;
        d["BQZ_Grade_Numeric"] = null;
        d["BQZ_Grade_Letter"]  = null;
        d["TQZ_Grade_Numeric"] = null;
        d["TQZ_Grade_Letter"]  = null;
        d["RQZ_Grade_Numeric"] = null;
        d["RQZ_Grade_Letter"]  = null;
        d["TTR_Percent"]       = null;
        d["TTR_Grade_Numeric"] = null;
        d["RTR_Percent"]       = null;
        d["RTR_Grade_Numeric"] = null;
        d["TCT_Grade_Numeric"] = null;
        d["TCT_Grade_Letter"]  = null;
        d["RCT_Grade_Numeric"] = null;
        d["RCT_Grade_Letter"]  = null;

        // ── Block E: 2D Quadrant — all null for 1D ────────────────────────────
        d["ULQZ_Grade_Numeric"]  = null;
        d["URQZ_Grade_Numeric"]  = null;
        d["RUQZ_Grade_Numeric"]  = null;
        d["RLQZ_Grade_Numeric"]  = null;
        d["ULQTTR_Percent"]      = null;
        d["ULQTTR_Grade_Numeric"]= null;
        d["URQTTR_Percent"]      = null;
        d["URQTTR_Grade_Numeric"]= null;
        d["LLQTTR_Percent"]      = null;
        d["LLQTTR_Grade_Numeric"]= null;
        d["LRQTTR_Percent"]      = null;
        d["LRQTTR_Grade_Numeric"]= null;
        d["ULQRTR_Percent"]      = null;
        d["ULQRTR_Grade_Numeric"]= null;
        d["URQRTR_Percent"]      = null;
        d["URQRTR_Grade_Numeric"]= null;
        d["LLQRTR_Percent"]      = null;
        d["LLQRTR_Grade_Numeric"]= null;
        d["LRQRTR_Percent"]      = null;
        d["LRQRTR_Grade_Numeric"]= null;
        d["ULQTCT_Grade_Numeric"]= null;
        d["URQTCT_Grade_Numeric"]= null;
        d["LLQTCT_Grade_Numeric"]= null;
        d["LRQTCT_Grade_Numeric"]= null;
        d["ULQRCT_Grade_Numeric"]= null;
        d["URQRCT_Grade_Numeric"]= null;
        d["LLQRCT_Grade_Numeric"]= null;
        d["LRQRCT_Grade_Numeric"]= null;

        // ── Block F: Military / Standards — null for standard 1D records ──────
        d["UIDFormat"]              = null;
        d["MilStd130VersionLetter"] = null;
        d["AS9132_Grade_Numeric"]   = null;
        d["Rmax"]                   = null;
        d["TargetRmax"]             = null;
        d["RmaxDeviation"]          = null;
        d["Rmin"]                   = null;
        d["TargetRmin"]             = null;
        d["RminDeviation"]          = null;

        // ── Block G: Vendor / Part Tracking ──────────────────────────────────
        d["VendorName"]   = r.VendorName;
        d["PartNumber"]   = r.PartNumber;
        d["SerialNumber"] = r.SerialNumber;

        // ── Block H: QR — null for 1D ─────────────────────────────────────────
        d["QR_Version"]    = null;
        d["QR_ECLevel"]    = null;
        d["QR_MaskPattern"]= null;

        // ── Custom Note ───────────────────────────────────────────────────────
        d["CustomNote"] = r.CustomNote;

        // ── Block I: DFC — pre-null; ExcelWriter.WriteDfcColumns fills if present
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
