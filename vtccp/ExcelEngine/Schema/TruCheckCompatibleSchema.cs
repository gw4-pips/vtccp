namespace ExcelEngine.Schema;

using ExcelEngine.Models;

/// <summary>
/// The TruCheck-compatible column schema — defines the 167-column layout used by VTCCP
/// to produce verification log files that open correctly in tools expecting the DMV
/// TruCheck column structure.
///
/// Column order follows the "grown to the right" evolutionary pattern:
///   Block A (cols 1–14):   Universal / Session header fields
///   Block B (cols 15–34):  1D ISO 15416 summary parameters
///   Block C (cols 35–55):  2D Common parameters (ISO 15415 shared)
///   Block D (cols 56–70):  2D Data Matrix standard parameters (≤26×26)
///   Block E (cols 71–100): 2D Data Matrix quadrant-expanded (≥32×32)
///   Block F (cols 101–110): Military / Standards-specific
///   Block G (cols 111–120): Vendor / Part tracking
///   Block H (cols 121+):   VTCCP extensions (QR-specific, DMRE — reserved, not written yet)
///
/// VTCCP additions beyond the base TruCheck column set:
///   - SchemaVersion metadata (written to a fixed non-data cell, not a schema column)
///   - Reflectance Margin (RM) — new in ISO 15415:2024 firmware
///   - Error Correction Budget / Errors Corrected / Error Capacity Used (newer General Chars)
///   - Pixels per Module, MRD (newer General Chars)
/// </summary>
public static class TruCheckCompatibleSchema
{
    public const string SchemaName = "TruCheckCompatible";

    public static ColumnSchema Build() => new()
    {
        Name        = SchemaName,
        Description = "167-column TruCheck-compatible layout for drop-in file compatibility. " +
                      "VTCCP-only fields appended after the standard columns.",
        Columns = BuildColumns(),
    };

    private static IReadOnlyList<ColumnDefinition> BuildColumns()
    {
        var cols = new List<ColumnDefinition>();

        // ── Block A: Universal / Session ──────────────────────────────────────
        cols.Add(Col("Date",           "Date",            10, SymbologyGroup.Universal, numberFormat: "yyyy-mm-dd"));
        cols.Add(Col("Time",           "Time",             9, SymbologyGroup.Universal, numberFormat: "hh:mm:ss"));
        cols.Add(Col("OperatorId",     "Operator Number",  9, SymbologyGroup.Universal));
        cols.Add(Col("RollNumber",     "Roll Number",      6, SymbologyGroup.Universal));
        cols.Add(Col("JobName",        "Job Name",        14, SymbologyGroup.Universal));
        cols.Add(Col("BatchNumber",    "Batch",           10, SymbologyGroup.Universal));
        cols.Add(Col("CompanyName",    "Company",         12, SymbologyGroup.Universal));
        cols.Add(Col("ProductName",    "Product",         12, SymbologyGroup.Universal));
        cols.Add(Col("Symbology",      "Symbology",       12, SymbologyGroup.Universal));
        cols.Add(Col("DecodedData",    "Data",            30, SymbologyGroup.Universal));
        cols.Add(Col("FormalGrade",    "Formal Grade",    14, SymbologyGroup.Universal));
        cols.Add(Col("OverallLetter",  "ANSI Letter Grade", 6, SymbologyGroup.Universal));
        cols.Add(Col("OverallNumeric", "ANSI Numeric Grade", 6, SymbologyGroup.Universal, numberFormat: "0.0"));
        cols.Add(Col("CustomPassFail", "Custom",           6, SymbologyGroup.Universal));
        cols.Add(Col("User1",          "User 1",           8, SymbologyGroup.Universal));
        cols.Add(Col("User2",          "User 2",           8, SymbologyGroup.Universal));
        cols.Add(Col("DeviceSerial",   "Unit Serial",     10, SymbologyGroup.Universal));
        cols.Add(Col("DeviceName",     "Device Name",     10, SymbologyGroup.Universal));
        cols.Add(Col("FirmwareVersion","Firmware",        10, SymbologyGroup.Universal));
        cols.Add(Col("CalibrationDate","Last Calibrated", 14, SymbologyGroup.Universal, numberFormat: "yyyy-mm-dd hh:mm"));
        cols.Add(Col("Aperture",       "Aperture (mil)",   6, SymbologyGroup.Universal));
        cols.Add(Col("Wavelength",     "Wavelength (nm)",  6, SymbologyGroup.Universal));
        cols.Add(Col("Lighting",       "Lighting",         7, SymbologyGroup.Universal));
        cols.Add(Col("Standard",       "Standard",        12, SymbologyGroup.Universal));

        // ── Block B: 1D ISO 15416 Parameters ──────────────────────────────────
        // These are blank for 2D records.
        var fam1D = new[] { SymbologyFamily.Linear1D };

        cols.Add(Col("SymbolAnsiGrade_Numeric", "Symbol ANSI Grade",  6, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("StartStopGrade_Numeric",  "Start/Stop Grade",   6, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("StartStopSrpGrade_Numeric","Start/Stop SRP Grade", 6, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_Edge",   "Edge",     6, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_RlRd",   "Rl/Rd",   8, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_SC",     "SC/CC",   6,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_MinEC",  "MinEC",   6,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_MOD",    "Mod/CMOD",6,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_Defect", "Def",     6,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_DCOD",   "DCD",     6,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_DEC",    "DEC",     6,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_LQZ",    "QZ-L",   5,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_RQZ",    "QZ-R",   5,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_HQZ",    "QZ-H",   5,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_MinQZ",  "QZ",     5,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("BWG_Percent","BWG",     5,  SymbologyGroup.Linear1D, fam1D, numberFormat: "0.0"));
        cols.Add(Col("BWG_Mil",    "BWG(MIL)",6, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Magnification","Magnification", 7, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("NominalXDim_1D","X Dim/Mag",   7, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("InspectionZoneHeight","Inspection Zone Height", 6, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("DecodableSymbolHeight","Decodable Symbol Height", 6, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Ratio",      "Ratio",   5,  SymbologyGroup.Linear1D, fam1D));

        // ── Block C: 2D Common Parameters (ISO 15415) ─────────────────────────
        var fam2D = new[]
        {
            SymbologyFamily.DataMatrix, SymbologyFamily.GS1DataMatrix,
            SymbologyFamily.RectangularDataMatrix, SymbologyFamily.DMRE,
            SymbologyFamily.QRCode, SymbologyFamily.GS1QRCode, SymbologyFamily.DotCode,
        };

        cols.Add(Col("UEC_Percent",       "UEC%",        6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("UEC_Grade_Numeric", "UEC Grade",   6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("SC_Percent",        "SC%",         6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("SC_RlRd",           "Rl/Rd",       8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("SC_Grade_Numeric",  "SC/CC",       6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("MOD_Grade_Numeric", "Mod/CMOD",    6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("RM_Grade_Numeric",  "RM Grade",    6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("ANU_Percent",       "ANU%",        6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("ANU_Grade_Numeric", "ANU Grade",   6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("GNU_Percent",       "GNU%",        6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("GNU_Grade_Numeric", "GNU Grade",   6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("FPD_Grade_Numeric", "FPD",         6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("DECODE_Grade_Numeric","DECODE Grade",6, SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("AG_Value",          "AG",          6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("AG_Grade_Numeric",  "AG/DDG",      6,  SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("AG_Grade_Letter",   "AG Grade",    6,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("OverallPassFail2D", "Overall Grade", 6, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("MatrixSize",        "Matrix Size", 8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("HorizontalBWG",     "Horizontal BWG", 6, SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("VerticalBWG",       "Vertical BWG",   6, SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("EncodedCharacters", "Encoded characters", 6, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("TotalCodewords",    "Total Codewords",    6, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("DataCodewords",     "Data Codewords",     6, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ErrorCorrectionBudget","Budget",     6, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ErrorsCorrected",   "Err",          5,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ErrorCapacityUsed", "Error Capacity Used", 6, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ErrorCorrectionType","EC Type",     6,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ImagePolarity",     "Image",        8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("NominalXDim_2D",    "Nominal X Dimension", 6, SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("PixelsPerModule",   "Pixels per Module",   6, SymbologyGroup.TwoDCommon, fam2D, numberFormat: "0.0"));
        cols.Add(Col("ContrastUniformity","CU",                  8, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("MRD",              "MRD Calculation",     12, SymbologyGroup.TwoDCommon, fam2D));

        // ── Block D: 2D Data Matrix Standard Parameters (≤26×26) ─────────────
        var famDM = new[]
        {
            SymbologyFamily.DataMatrix, SymbologyFamily.GS1DataMatrix,
            SymbologyFamily.RectangularDataMatrix, SymbologyFamily.DMRE,
        };

        cols.Add(Col("LLS_Grade_Numeric", "LLS",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("LLS_Grade_Letter",  "LLS Ltr", 5, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("BLS_Grade_Numeric", "BLS",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("BLS_Grade_Letter",  "BLS Ltr", 5, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("LQZ_Grade_Numeric", "LQZ",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("LQZ_Grade_Letter",  "LQZ Ltr", 5, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("BQZ_Grade_Numeric", "BQZ",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("BQZ_Grade_Letter",  "BQZ Ltr", 5, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("TQZ_Grade_Numeric", "TQZ",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("TQZ_Grade_Letter",  "TQZ Ltr", 5, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("RQZ_Grade_Numeric", "RQZ",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("RQZ_Grade_Letter",  "RQZ Ltr", 5, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("TTR_Percent",       "TTR%",    6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("TTR_Grade_Numeric", "TTR",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("RTR_Percent",       "RTR%",    6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("RTR_Grade_Numeric", "RTR",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("TCT_Grade_Numeric", "TCT",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("TCT_Grade_Letter",  "TCT Ltr", 5, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("RCT_Grade_Numeric", "RCT",     6, SymbologyGroup.TwoDDataMatrix, famDM, numberFormat: "0.0"));
        cols.Add(Col("RCT_Grade_Letter",  "RCT Ltr", 5, SymbologyGroup.TwoDDataMatrix, famDM));

        // ── Block E: 2D Data Matrix Quadrant-Expanded (≥32×32) ────────────────
        cols.Add(Col("ULQZ_Grade_Numeric", "ULQZ",    6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("URQZ_Grade_Numeric", "URQZ",    6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("RUQZ_Grade_Numeric", "RUQZ",    6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("RLQZ_Grade_Numeric", "RLQZ",    6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));

        cols.Add(Col("ULQTTR_Percent",     "ULQTTR%", 6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("ULQTTR_Grade_Numeric","ULQTTR",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("URQTTR_Percent",     "URQTTR%", 6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("URQTTR_Grade_Numeric","URQTTR",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LLQTTR_Percent",     "LLQTTR%", 6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LLQTTR_Grade_Numeric","LLQTTR",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LRQTTR_Percent",     "LRQTTR%", 6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LRQTTR_Grade_Numeric","LRQTTR",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));

        cols.Add(Col("ULQRTR_Percent",     "ULQRTR%", 6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("ULQRTR_Grade_Numeric","ULQRTR",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("URQRTR_Percent",     "URQRTR%", 6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("URQRTR_Grade_Numeric","URQRTR",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LLQRTR_Percent",     "LLQRTR%", 6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LLQRTR_Grade_Numeric","LLQRTR",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LRQRTR_Percent",     "LRQRTR%", 6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LRQRTR_Grade_Numeric","LRQRTR",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));

        cols.Add(Col("ULQTCT_Grade_Numeric","ULQTCT",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("URQTCT_Grade_Numeric","URQTCT",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LLQTCT_Grade_Numeric","LLQTCT",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LRQTCT_Grade_Numeric","LRQTCT",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));

        cols.Add(Col("ULQRCT_Grade_Numeric","ULQRCT",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("URQRCT_Grade_Numeric","URQRCT",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LLQRCT_Grade_Numeric","LLQRCT",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));
        cols.Add(Col("LRQRCT_Grade_Numeric","LRQRCT",  6, SymbologyGroup.TwoDDataMatrixQuadrant, famDM, numberFormat: "0.0"));

        // ── Block F: Military / Standards-Specific ────────────────────────────
        cols.Add(Col("UIDFormat",             "UID Format",            10, SymbologyGroup.MilStd));
        cols.Add(Col("MilStd130VersionLetter","MIL-130 Version Letter", 6, SymbologyGroup.MilStd));
        cols.Add(Col("AS9132_Grade_Numeric",  "AS9132 Grade",           6, SymbologyGroup.MilStd, numberFormat: "0.0"));
        cols.Add(Col("Rmax",                  "Rmax",                   6, SymbologyGroup.MilStd));
        cols.Add(Col("TargetRmax",            "Target Rmax",            7, SymbologyGroup.MilStd));
        cols.Add(Col("RmaxDeviation",         "Rmax Deviation",         7, SymbologyGroup.MilStd));
        cols.Add(Col("Rmin",                  "Rmin",                   6, SymbologyGroup.MilStd));
        cols.Add(Col("TargetRmin",            "Target Rmin",            7, SymbologyGroup.MilStd));
        cols.Add(Col("RminDeviation",         "Rmin Deviation",         7, SymbologyGroup.MilStd));

        // ── Block G: Vendor / Part Tracking ──────────────────────────────────
        cols.Add(Col("VendorName",  "Vendor",     12, SymbologyGroup.VendorPartTracking));
        cols.Add(Col("PartNumber",  "Part Num",   12, SymbologyGroup.VendorPartTracking));
        cols.Add(Col("SerialNumber","Serial Num", 12, SymbologyGroup.VendorPartTracking));

        // ── Block H: VTCCP Extensions (QR reserved, not yet written) ──────────
        cols.Add(Col("QR_Version",    "QR Version",  7, SymbologyGroup.TwoDQR,
            new[] { SymbologyFamily.QRCode, SymbologyFamily.GS1QRCode }));
        cols.Add(Col("QR_ECLevel",    "QR EC Level", 6, SymbologyGroup.TwoDQR,
            new[] { SymbologyFamily.QRCode, SymbologyFamily.GS1QRCode }));
        cols.Add(Col("QR_MaskPattern","QR Mask",     6, SymbologyGroup.TwoDQR,
            new[] { SymbologyFamily.QRCode, SymbologyFamily.GS1QRCode }));

        // Custom note
        cols.Add(Col("CustomNote", "Custom Note", 20, SymbologyGroup.Universal));

        // ── Block I: GS1 / Data Format Check ──────────────────────────────────
        var famGS1 = new[]
        {
            SymbologyFamily.GS1DataMatrix, SymbologyFamily.GS1QRCode,
        };

        cols.Add(Col("DFC_Standard", "DFC Standard", 22, SymbologyGroup.Universal, famGS1));
        for (int slot = 1; slot <= 8; slot++)
        {
            cols.Add(Col($"DFC_R{slot}_Name",  $"DFC R{slot} Name",  10, SymbologyGroup.Universal, famGS1));
            cols.Add(Col($"DFC_R{slot}_Data",  $"DFC R{slot} Data",  18, SymbologyGroup.Universal, famGS1));
            cols.Add(Col($"DFC_R{slot}_Check", $"DFC R{slot} Check",  8, SymbologyGroup.Universal, famGS1));
        }

        return cols;
    }

    private static ColumnDefinition Col(
        string fieldId,
        string displayName,
        double width,
        SymbologyGroup group,
        IEnumerable<SymbologyFamily>? families = null,
        string? numberFormat = null) =>
        new()
        {
            FieldId = fieldId,
            DisplayName = displayName,
            Width = width,
            Group = group,
            ApplicableFamilies = families?.ToArray() ?? [],
            NumberFormat = numberFormat,
        };
}
