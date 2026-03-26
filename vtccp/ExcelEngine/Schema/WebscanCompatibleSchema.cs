namespace ExcelEngine.Schema;

using ExcelEngine.Models;

/// <summary>
/// The WebscanCompatible column schema — replicates the column order used by Webscan TruCheck
/// as extracted from the Webscan_Data_Capture30.xls and CalCardProd.xls reference files.
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
/// VTCCP additions not in Webscan:
///   - SchemaVersion metadata (written to a fixed non-data cell, not a schema column)
///   - Reflectance Margin (RM) — new in ISO 15415:2024 firmware
///   - Error Correction Budget / Errors Corrected / Error Capacity Used (newer General Chars)
///   - Pixels per Module, MRD (newer General Chars)
/// </summary>
public static class WebscanCompatibleSchema
{
    public const string SchemaName = "WebscanCompatible";

    public static ColumnSchema Build() => new()
    {
        Name = SchemaName,
        Description = "Replicates the Webscan TruCheck column layout for drop-in file compatibility. " +
                      "VTCCP-only fields appended after the standard Webscan columns.",
        Columns = BuildColumns(),
    };

    private static IReadOnlyList<ColumnDefinition> BuildColumns()
    {
        var cols = new List<ColumnDefinition>();

        // ── Block A: Universal / Session ──────────────────────────────────────
        cols.Add(Col("Date",           "Date",            14, SymbologyGroup.Universal, numberFormat: "yyyy-mm-dd"));
        cols.Add(Col("Time",           "Time",            10, SymbologyGroup.Universal, numberFormat: "hh:mm:ss"));
        cols.Add(Col("OperatorId",     "Operator Number", 14, SymbologyGroup.Universal));
        cols.Add(Col("RollNumber",     "Roll Number",     10, SymbologyGroup.Universal));
        cols.Add(Col("JobName",        "Job Name",        20, SymbologyGroup.Universal));
        cols.Add(Col("BatchNumber",    "Batch",           16, SymbologyGroup.Universal));
        cols.Add(Col("CompanyName",    "Company",         22, SymbologyGroup.Universal));
        cols.Add(Col("ProductName",    "Product",         22, SymbologyGroup.Universal));
        cols.Add(Col("Symbology",      "Symbology",       18, SymbologyGroup.Universal));
        cols.Add(Col("DecodedData",    "Data",            40, SymbologyGroup.Universal));
        cols.Add(Col("FormalGrade",    "Formal Grade",    18, SymbologyGroup.Universal));
        cols.Add(Col("OverallLetter",  "ANSI Letter Grade", 14, SymbologyGroup.Universal));
        cols.Add(Col("OverallNumeric", "ANSI Numeric Grade", 14, SymbologyGroup.Universal));
        cols.Add(Col("CustomPassFail", "Custom",          10, SymbologyGroup.Universal));
        cols.Add(Col("User1",          "User 1",          14, SymbologyGroup.Universal));
        cols.Add(Col("User2",          "User 2",          14, SymbologyGroup.Universal));
        cols.Add(Col("DeviceSerial",   "Unit Serial",     18, SymbologyGroup.Universal));
        cols.Add(Col("DeviceName",     "Device Name",     16, SymbologyGroup.Universal));
        cols.Add(Col("FirmwareVersion","Firmware",        14, SymbologyGroup.Universal));
        cols.Add(Col("CalibrationDate","Last Calibrated", 16, SymbologyGroup.Universal, numberFormat: "yyyy-mm-dd hh:mm"));

        // ── Block B: 1D ISO 15416 Parameters ──────────────────────────────────
        // These are blank for 2D records.
        var fam1D = new[] { SymbologyFamily.Linear1D };

        cols.Add(Col("SymbolAnsiGrade_Numeric", "Symbol ANSI Grade", 14, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("StartStopGrade_Numeric",  "Start/Stop Grade",  14, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("StartStopSrpGrade_Numeric","Start/Stop SRP Grade", 16, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_Edge",   "Edge",   10, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_RlRd",   "Rl/Rd",  10, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_SC",     "SC",      8,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_MinEC",  "MinEC",   8,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_MOD",    "MOD",     8,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_Defect", "DEF",     8,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_DCOD",   "DCOD",    8,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_DEC",    "DEC",     8,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Avg_MinQZ",  "MinQZ",   8,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("BWG_Percent","BWG%",    8,  SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("BWG_Mil",    "BWG(MIL)",10, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Magnification","Magnification", 12, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("NominalXDim_1D","X Dim/Mag",   12, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("InspectionZoneHeight","Inspection Zone Height", 18, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("DecodableSymbolHeight","Decodable Symbol Height", 20, SymbologyGroup.Linear1D, fam1D));
        cols.Add(Col("Ratio",      "Ratio",   8,  SymbologyGroup.Linear1D, fam1D));

        // ── Block C: 2D Common Parameters (ISO 15415) ─────────────────────────
        var fam2D = new[]
        {
            SymbologyFamily.DataMatrix, SymbologyFamily.GS1DataMatrix,
            SymbologyFamily.RectangularDataMatrix, SymbologyFamily.DMRE,
            SymbologyFamily.QRCode, SymbologyFamily.GS1QRCode, SymbologyFamily.DotCode,
        };

        cols.Add(Col("UEC_Percent",       "UEC%",        8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("UEC_Grade_Numeric", "UEC Grade",   10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("SC_Percent",        "SC%",         8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("SC_RlRd",           "Rl/Rd",       10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("SC_Grade_Numeric",  "SC/CC",       10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("MOD_Grade_Numeric", "Mod/CMOD",    10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("RM_Grade_Numeric",  "RM Grade",    10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ANU_Percent",       "ANU%",        8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ANU_Grade_Numeric", "ANU Grade",   10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("GNU_Percent",       "GNU%",        8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("GNU_Grade_Numeric", "GNU Grade",   10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("FPD_Grade_Numeric", "FPD Grade",   10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("DECODE_Grade_Numeric","DECODE Grade",10,SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("AG_Value",          "AG",          8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("AG_Grade_Numeric",  "AG/DDG",      10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("AG_Grade_Letter",   "AG Grade",    10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("OverallPassFail2D", "Overall Grade", 12, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("MatrixSize",        "Matrix Size", 18, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("HorizontalBWG",     "Horizontal BWG", 14, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("VerticalBWG",       "Vertical BWG",   14, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("EncodedCharacters", "Encoded characters", 16, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("TotalCodewords",    "Total Codewords",   14, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("DataCodewords",     "Data Codewords",    14, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ErrorCorrectionBudget","Budget",    10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ErrorsCorrected",   "Err",          8,  SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ErrorCapacityUsed", "Error Capacity Used", 16, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ErrorCorrectionType","EC Type",    10, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ImagePolarity",     "Image",       16, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("NominalXDim_2D",    "Nominal X Dimension", 16, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("PixelsPerModule",   "Pixels per Module",   16, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("ContrastUniformity","Contrast Uniformity", 16, SymbologyGroup.TwoDCommon, fam2D));
        cols.Add(Col("MRD",              "MRD Calculation",      16, SymbologyGroup.TwoDCommon, fam2D));

        // ── Block D: 2D Data Matrix Standard Parameters (≤26×26) ─────────────
        var famDM = new[]
        {
            SymbologyFamily.DataMatrix, SymbologyFamily.GS1DataMatrix,
            SymbologyFamily.RectangularDataMatrix, SymbologyFamily.DMRE,
        };

        cols.Add(Col("LLS_Grade_Numeric", "LLS Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("LLS_Grade_Letter",  "LLS",         6,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("BLS_Grade_Numeric", "BLS Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("BLS_Grade_Letter",  "BLS",         6,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("LQZ_Grade_Numeric", "LQZ Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("LQZ_Grade_Letter",  "LQZ",         6,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("BQZ_Grade_Numeric", "BQZ Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("BQZ_Grade_Letter",  "BQZ",         6,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("TQZ_Grade_Numeric", "TQZ Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("TQZ_Grade_Letter",  "TQZ",         6,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("RQZ_Grade_Numeric", "RQZ Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("RQZ_Grade_Letter",  "RQZ",         6,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("TTR_Percent",       "TTR%",        8,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("TTR_Grade_Numeric", "TTR Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("RTR_Percent",       "RTR%",        8,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("RTR_Grade_Numeric", "RTR Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("TCT_Grade_Numeric", "TCT Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("TCT_Grade_Letter",  "TCT",         6,  SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("RCT_Grade_Numeric", "RCT Grade",  10, SymbologyGroup.TwoDDataMatrix, famDM));
        cols.Add(Col("RCT_Grade_Letter",  "RCT",         6,  SymbologyGroup.TwoDDataMatrix, famDM));

        // ── Block E: 2D Data Matrix Quadrant-Expanded (≥32×32) ────────────────
        cols.Add(Col("ULQZ_Grade_Numeric", "ULQZ",        8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("URQZ_Grade_Numeric", "URQZ",        8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("RUQZ_Grade_Numeric", "RUQZ",        8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("RLQZ_Grade_Numeric", "RLQZ",        8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));

        cols.Add(Col("ULQTTR_Percent",     "ULQTTR%",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("ULQTTR_Grade_Numeric","ULQTTR",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("URQTTR_Percent",     "URQTTR%",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("URQTTR_Grade_Numeric","URQTTR",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LLQTTR_Percent",     "LLQTTR%",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LLQTTR_Grade_Numeric","LLQTTR",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LRQTTR_Percent",     "LRQTTR%",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LRQTTR_Grade_Numeric","LRQTTR",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));

        cols.Add(Col("ULQRTR_Percent",     "ULQRTR%",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("ULQRTR_Grade_Numeric","ULQRTR",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("URQRTR_Percent",     "URQRTR%",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("URQRTR_Grade_Numeric","URQRTR",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LLQRTR_Percent",     "LLQRTR%",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LLQRTR_Grade_Numeric","LLQRTR",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LRQRTR_Percent",     "LRQRTR%",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LRQRTR_Grade_Numeric","LRQRTR",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));

        cols.Add(Col("ULQTCT_Grade_Numeric","ULQTCT",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("URQTCT_Grade_Numeric","URQTCT",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LLQTCT_Grade_Numeric","LLQTCT",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LRQTCT_Grade_Numeric","LRQTCT",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));

        cols.Add(Col("ULQRCT_Grade_Numeric","ULQRCT",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("URQRCT_Grade_Numeric","URQRCT",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LLQRCT_Grade_Numeric","LLQRCT",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));
        cols.Add(Col("LRQRCT_Grade_Numeric","LRQRCT",     8, SymbologyGroup.TwoDDataMatrixQuadrant, famDM));

        // ── Block F: Military / Standards-Specific ────────────────────────────
        cols.Add(Col("UIDFormat",          "UID Format",         14, SymbologyGroup.MilStd));
        cols.Add(Col("MilStd130VersionLetter","MIL-130 Version Letter", 18, SymbologyGroup.MilStd));
        cols.Add(Col("AS9132_Grade_Numeric","AS9132 Grade",       12, SymbologyGroup.MilStd));
        cols.Add(Col("Rmax",               "Rmax",               10, SymbologyGroup.MilStd));
        cols.Add(Col("TargetRmax",         "Target  Rmax",       12, SymbologyGroup.MilStd));
        cols.Add(Col("RmaxDeviation",      "Rmax Deviation",     12, SymbologyGroup.MilStd));
        cols.Add(Col("Rmin",               "Rmin",               10, SymbologyGroup.MilStd));
        cols.Add(Col("TargetRmin",         "Target Rmin",        12, SymbologyGroup.MilStd));
        cols.Add(Col("RminDeviation",      "Rmin Deviation",     12, SymbologyGroup.MilStd));

        // ── Block G: Vendor / Part Tracking ──────────────────────────────────
        cols.Add(Col("VendorName",         "Vendor",             16, SymbologyGroup.VendorPartTracking));
        cols.Add(Col("PartNumber",         "Part Num",           16, SymbologyGroup.VendorPartTracking));
        cols.Add(Col("SerialNumber",       "Serial Num",         16, SymbologyGroup.VendorPartTracking));

        // ── Block H: VTCCP Extensions (QR reserved, not yet written) ──────────
        // QR code version info — reserved column positions
        cols.Add(Col("QR_Version",         "QR Version",         12, SymbologyGroup.TwoDQR,
            new[] { SymbologyFamily.QRCode, SymbologyFamily.GS1QRCode }));
        cols.Add(Col("QR_ECLevel",         "QR EC Level",        10, SymbologyGroup.TwoDQR,
            new[] { SymbologyFamily.QRCode, SymbologyFamily.GS1QRCode }));
        cols.Add(Col("QR_MaskPattern",     "QR Mask",             8, SymbologyGroup.TwoDQR,
            new[] { SymbologyFamily.QRCode, SymbologyFamily.GS1QRCode }));

        // Custom note — appended at end (VTCCP addition)
        cols.Add(Col("CustomNote",         "Custom Note",         30, SymbologyGroup.Universal));

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
