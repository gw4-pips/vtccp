// VTCCP TestHarness -- Phase 1 Tasks 1+2
// Writes sample 2D Data Matrix verification records to both .xlsx and .xls

using ExcelEngine.Adapters;
using ExcelEngine.Models;
using ExcelEngine.Schema;
using ExcelEngine.Session;
using ExcelEngine.Writer;

// ─── Schema / session setup ────────────────────────────────────────────────────
OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

var schemaManager = new ColumnSchemaManager();
var schema = schemaManager.GetActive();

Console.WriteLine("VTCCP Excel Engine -- Phase 1 Tasks 1+2");
Console.WriteLine("=========================================");
Console.WriteLine($"Schema: {schema.Name}, {schema.Columns.Count} columns");

var errors = ColumnSchemaManager.Validate(schema);
Console.WriteLine(errors.Count == 0 ? "Schema: PASS" : $"Schema: FAIL ({errors.Count} errors)");

// ─── Session ────────────────────────────────────────────────────────────────────
var session = new SessionState
{
    JobName      = "CalCardProd",
    OperatorId   = "GW4",
    RollNumber   = 1,
    CompanyName  = "Product Identification and Processing Systems, Inc.",
    DeviceSerial = "1A2202PP003652",
    DeviceName   = "DM475-8E9364",
    FirmwareVersion = "6.1.16_sr3",
    CalibrationDate = new DateTime(2025, 8, 11, 16, 50, 0),
    OutputDirectory = Path.Combine(Path.GetTempPath(), "vtccp_test_output"),
    SessionStarted = new DateTime(2025, 8, 11),
};

Directory.CreateDirectory(session.OutputDirectory!);

// ─── Test Record 1: GRADE-4-A (all A grades, 4.0 numeric) ─────────────────────
// Source: _F1_010123456789012810GRADE-4-A-AI-INC ... PDF
var grade4 = new GradingResult { NumericGrade = 4.0m, LetterGrade = GradeLetterValue.A, PassFail = OverallPassFail.Pass };
var grade4Pass = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS");

var record1 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2025, 8, 11, 17, 48, 18),
    Symbology        = "GS1 DataMatrix",
    SymbologyFamily  = SymbologyFamily.GS1DataMatrix,
    DecodedData      = "<F1>010123456789012810GRADE-4-A-AI-INC",
    FormalGrade      = "4.0/16/660/45Q",
    OverallGrade     = grade4Pass,
    CustomPassFail   = OverallPassFail.Pass,
    OperatorId       = session.OperatorId,
    JobName          = session.JobName,
    RollNumber       = session.RollNumber,
    CompanyName      = session.CompanyName,
    CustomNote       = "Pre-shipment Conformance Challenge for DANCO Medical",
    DeviceSerial     = session.DeviceSerial,
    DeviceName       = session.DeviceName,
    FirmwareVersion  = session.FirmwareVersion,
    CalibrationDate  = session.CalibrationDate,
    Aperture         = 16,
    Wavelength       = 660,
    Lighting         = "45Q",
    Standard         = "ISO 15415:2011",

    // General Characteristics (from PDF)
    MatrixSize             = "22x22 (Data: 20x20)",
    HorizontalBWG          = -11,
    VerticalBWG            = -11,
    EncodedCharacters      = 35,
    TotalCodewords         = 50,
    DataCodewords          = 30,
    ErrorCorrectionBudget  = 20,
    ErrorsCorrected        = 0,
    ErrorCapacityUsed      = 0,
    ErrorCorrectionType    = "ECC 200",
    ImagePolarity          = ImagePolarity.WhiteOnBlack,
    NominalXDim_2D         = 19.8m,
    PixelsPerModule        = 33.24m,
    ContrastUniformity     = "72 at module(10,5)",
    MRD                    = "71% (77% - 6%)",

    // ISO 15415 Quality Parameters (from PDF)
    UEC_Percent      = 100,
    UEC_Grade        = grade4Pass,
    SC_Percent       = 84,
    SC_RlRd          = "89/4",
    SC_Grade         = grade4Pass,
    MOD_Grade        = grade4Pass,
    RM_Grade         = grade4Pass,
    ANU_Percent      = 0.1m,
    ANU_Grade        = grade4Pass,
    GNU_Percent      = 2.3m,
    GNU_Grade        = grade4Pass,
    FPD_Grade        = grade4Pass,
    LLS_Grade        = grade4Pass,
    BLS_Grade        = grade4Pass,
    LQZ_Grade        = grade4Pass,
    BQZ_Grade        = grade4Pass,
    TQZ_Grade        = grade4Pass,
    RQZ_Grade        = grade4Pass,
    TTR_Percent      = 0,
    TTR_Grade        = grade4Pass,
    RTR_Percent      = 0,
    RTR_Grade        = grade4Pass,
    TCT_Grade        = grade4Pass,
    RCT_Grade        = grade4Pass,
    AG_Value         = 4.0m,
    AG_Grade         = grade4Pass,
    DECODE_Grade     = grade4Pass,
};

// ─── Test Record 2: GRADE-1-ANU (ANU = 1.0 D FAIL, others A) ─────────────────
// Source: _F1_010123456789012810GRADE-1-ANU-AI-INC ... PDF
var grade1d = GradingResult.FromLetterAndNumeric("D", 1.0m, "FAIL");

var record2 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2025, 8, 11, 17, 51, 33),
    Symbology        = "GS1 DataMatrix",
    SymbologyFamily  = SymbologyFamily.GS1DataMatrix,
    DecodedData      = "<F1>010123456789012810GRADE-1-ANU-AI-INC",
    FormalGrade      = "1.0/17/660/45Q",
    OverallGrade     = GradingResult.FromLetterAndNumeric("D", 1.0m, "FAIL"),
    CustomPassFail   = OverallPassFail.Fail,
    OperatorId       = session.OperatorId,
    JobName          = session.JobName,
    RollNumber       = session.RollNumber,
    CompanyName      = session.CompanyName,
    CustomNote       = "Pre-shipment Conformance Challenge for DANCO Medical",
    DeviceSerial     = session.DeviceSerial,
    DeviceName       = session.DeviceName,
    FirmwareVersion  = session.FirmwareVersion,
    CalibrationDate  = session.CalibrationDate,
    Aperture         = 17,
    Wavelength       = 660,
    Lighting         = "45Q",
    Standard         = "ISO 15415:2011",

    MatrixSize             = "22x22 (Data: 20x20)",
    HorizontalBWG          = 7,
    VerticalBWG            = 7,
    EncodedCharacters      = 37,
    TotalCodewords         = 50,
    DataCodewords          = 30,
    ErrorCorrectionBudget  = 20,
    ErrorsCorrected        = 0,
    ErrorCapacityUsed      = 0,
    ErrorCorrectionType    = "ECC 200",
    ImagePolarity          = ImagePolarity.BlackOnWhite,
    NominalXDim_2D         = 20.9m,
    PixelsPerModule        = 35.26m,
    ContrastUniformity     = "75 at module(10,1)",
    MRD                    = "71% (78% - 7%)",

    UEC_Percent      = 100,
    UEC_Grade        = grade4Pass,
    SC_Percent       = 84,
    SC_RlRd          = "89/4",
    SC_Grade         = grade4Pass,
    MOD_Grade        = grade4Pass,
    RM_Grade         = grade4Pass,
    ANU_Percent      = 11.0m,
    ANU_Grade        = grade1d,       // <-- FAIL here
    GNU_Percent      = 2.8m,
    GNU_Grade        = grade4Pass,
    FPD_Grade        = grade4Pass,
    LLS_Grade        = grade4Pass,
    BLS_Grade        = grade4Pass,
    LQZ_Grade        = grade4Pass,
    BQZ_Grade        = grade4Pass,
    TQZ_Grade        = grade4Pass,
    RQZ_Grade        = grade4Pass,
    TTR_Percent      = 0,
    TTR_Grade        = grade4Pass,
    RTR_Percent      = 0,
    RTR_Grade        = grade4Pass,
    TCT_Grade        = grade4Pass,
    RCT_Grade        = grade4Pass,
    AG_Value         = 4.0m,
    AG_Grade         = grade4Pass,
    DECODE_Grade     = grade4Pass,
};

// ─── Test Record 3: Third PDF (DM475-866D76, with GS1 DFC) ────────────────────
// Source: 2025-08-11_13-46-46-051 PDF
var record3 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2025, 8, 11, 13, 46, 44),
    Symbology        = "GS1 DataMatrix",
    SymbologyFamily  = SymbologyFamily.GS1DataMatrix,
    DecodedData      = "<F1>010123456789012810GRADE-4-A-AI-INC",
    FormalGrade      = "4.0/16/660/45Q",
    OverallGrade     = grade4Pass,
    CustomPassFail   = OverallPassFail.Pass,
    OperatorId       = session.OperatorId,
    JobName          = session.JobName,
    RollNumber       = session.RollNumber,
    BatchNumber      = "GRADE-4-A-AI-INC",
    CompanyName      = session.CompanyName,
    DeviceSerial     = session.DeviceSerial,
    DeviceName       = "DM475-866D76",
    FirmwareVersion  = session.FirmwareVersion,
    CalibrationDate  = new DateTime(2025, 8, 5, 18, 3, 0),
    Aperture         = 16,
    Wavelength       = 660,
    Lighting         = "45Q",
    Standard         = "ISO 15415:2011",

    MatrixSize             = "22x22 (Data: 20x20)",
    HorizontalBWG          = -11,
    VerticalBWG            = -11,
    EncodedCharacters      = 35,
    TotalCodewords         = 50,
    DataCodewords          = 30,
    ErrorCorrectionBudget  = 20,
    ErrorsCorrected        = 0,
    ErrorCapacityUsed      = 0,
    ErrorCorrectionType    = "ECC 200",
    ImagePolarity          = ImagePolarity.WhiteOnBlack,
    NominalXDim_2D         = 19.6m,
    PixelsPerModule        = 33.17m,
    ContrastUniformity     = "73 at module(10,5)",
    MRD                    = "78% (87% - 9%)",

    UEC_Percent      = 100,
    UEC_Grade        = grade4Pass,
    SC_Percent       = 94,
    SC_RlRd          = "100/6",
    SC_Grade         = grade4Pass,
    MOD_Grade        = grade4Pass,
    RM_Grade         = grade4Pass,
    ANU_Percent      = 0.2m,
    ANU_Grade        = grade4Pass,
    GNU_Percent      = 2.1m,
    GNU_Grade        = grade4Pass,
    FPD_Grade        = grade4Pass,
    LLS_Grade        = grade4Pass,
    BLS_Grade        = grade4Pass,
    LQZ_Grade        = grade4Pass,
    BQZ_Grade        = grade4Pass,
    TQZ_Grade        = grade4Pass,
    RQZ_Grade        = grade4Pass,
    TTR_Percent      = 0,
    TTR_Grade        = grade4Pass,
    RTR_Percent      = 0,
    RTR_Grade        = grade4Pass,
    TCT_Grade        = grade4Pass,
    RCT_Grade        = grade4Pass,
    AG_Value         = 4.0m,
    AG_Grade         = grade4Pass,
    DECODE_Grade     = grade4Pass,

    DataFormatCheck  = new ExcelEngine.Models.DataFormatCheckResult
    {
        Overall  = OverallPassFail.Pass,
        Standard = "GS1 Application Data Format",
        Rows     =
        [
            new() { Name = "GS1 Header",  Data = "<F1>",             Check = "PASS" },
            new() { Name = "AI:GTIN",     Data = "01",               Check = "PASS" },
            new() { Name = "GTIN",        Data = "0123456789012",    Check = "PASS" },
            new() { Name = "Chk Digit",   Data = "8",                Check = "PASS" },
            new() { Name = "AI:BatchLot", Data = "10",               Check = "PASS" },
            new() { Name = "BatchLot",    Data = "GRADE-4-A-AI-INC", Check = "PASS" },
        ],
    },
};

var allRecords = new[] { record1, record2, record3 };

// ─── Write XLSX ────────────────────────────────────────────────────────────────
var xlsxPath = Path.Combine(session.OutputDirectory!, ExcelFileManager.GenerateFileName(session, OutputFormat.Xlsx));
if (File.Exists(xlsxPath)) File.Delete(xlsxPath);   // always start fresh (avoid stale column layout)
Console.WriteLine($"\nWriting XLSX: {xlsxPath}");

using (var adapter = new XlsxAdapter())
using (var writer  = new ExcelWriter(adapter, schema, session))
{
    writer.Open(xlsxPath);
    foreach (var rec in allRecords)
        writer.AppendRecord(rec);
    writer.Save();
}

var xlsxInfo = new FileInfo(xlsxPath);
Console.WriteLine($"  Done. Size: {xlsxInfo.Length:N0} bytes, rows: {allRecords.Length} data rows");

// ─── Write XLS ────────────────────────────────────────────────────────────────
var xlsPath = Path.Combine(session.OutputDirectory!, ExcelFileManager.GenerateFileName(session, OutputFormat.Xls));
if (File.Exists(xlsPath)) File.Delete(xlsPath);   // always start fresh
Console.WriteLine($"\nWriting XLS:  {xlsPath}");

using (var adapter = new XlsAdapter())
using (var writer  = new ExcelWriter(adapter, schema, session))
{
    writer.Open(xlsPath);
    foreach (var rec in allRecords)
        writer.AppendRecord(rec);
    writer.Save();
}

var xlsInfo = new FileInfo(xlsPath);
Console.WriteLine($"  Done. Size: {xlsInfo.Length:N0} bytes, rows: {allRecords.Length} data rows");

// ─── Verify column count in written XLSX ──────────────────────────────────────
Console.WriteLine("\nVerifying XLSX column structure...");
OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
using (var pkg = new OfficeOpenXml.ExcelPackage(new FileInfo(xlsxPath)))
{
    var ws = pkg.Workbook.Worksheets["Main"];
    int headerCols = ws.Dimension?.Columns ?? 0;
    int rows = ws.Dimension?.Rows ?? 0;
    Console.WriteLine($"  Sheet 'Main': {rows} rows x {headerCols} columns");
    Console.WriteLine($"  Header row 2, col 1: '{ws.Cells[2,1].Text}'");
    Console.WriteLine($"  Header row 2, col 2: '{ws.Cells[2,2].Text}'");
    Console.WriteLine($"  Data row 3, col 9 (Symbology): '{ws.Cells[3,9].Text}'");
    Console.WriteLine($"  Data row 3, col 10 (Data): '{ws.Cells[3,10].Text}'");
    Console.WriteLine($"  Data row 3, col 11 (FormalGrade): '{ws.Cells[3,11].Text}'");
    Console.WriteLine($"  Data row 4, col 12 (OverallLetter): '{ws.Cells[4,12].Text}'");  // row4 = record2
    // headerCols >= schema.Columns.Count because SchemaVersionWriter adds cells past the last data col.
    Console.WriteLine($"  Columns >= schema: {(headerCols >= schema.Columns.Count ? "YES" : $"NO ({headerCols} vs {schema.Columns.Count})")}");
}

// ─── File naming test ──────────────────────────────────────────────────────────
Console.WriteLine("\nFile naming tests:");
Console.WriteLine($"  Xlsx name: {ExcelFileManager.GenerateFileName(session, OutputFormat.Xlsx)}");
Console.WriteLine($"  Xls name:  {ExcelFileManager.GenerateFileName(session, OutputFormat.Xls)}");
var noJob = new SessionState { OperatorId = "GW4", SessionStarted = new DateTime(2025, 8, 11) };
Console.WriteLine($"  No-job xlsx: {ExcelFileManager.GenerateFileName(noJob, OutputFormat.Xlsx)}");
var noJobNoOp = new SessionState { SessionStarted = new DateTime(2025, 8, 11) };
Console.WriteLine($"  Fallback:    {ExcelFileManager.GenerateFileName(noJobNoOp, OutputFormat.Xlsx)}");

Console.WriteLine("\nTask 2 complete. (2D Data Matrix records verified)");

// ─── DFC column verification ───────────────────────────────────────────────────
Console.WriteLine("\nVerifying GS1 DFC columns...");
using (var pkg2 = new OfficeOpenXml.ExcelPackage(new FileInfo(xlsxPath)))
{
    var ws2 = pkg2.Workbook.Worksheets["Main"];
    // Find DFC_Standard column by scanning header row
    int dfcCol = -1;
    for (int c = 1; c <= ws2.Dimension.Columns; c++)
        if (ws2.Cells[2, c].Text == "DFC Standard") { dfcCol = c; break; }

    Console.WriteLine($"  DFC_Standard column: {dfcCol}");
    // record3 = row 5 (row 1=title, 2=header, 3=rec1, 4=rec2, 5=rec3)
    Console.WriteLine($"  DFC_Standard (rec3, has DFC): '{ws2.Cells[5, dfcCol].Text}'");
    Console.WriteLine($"  DFC R1 Name  (rec3): '{ws2.Cells[5, dfcCol+1].Text}'");
    Console.WriteLine($"  DFC R1 Data  (rec3): '{ws2.Cells[5, dfcCol+2].Text}'");
    Console.WriteLine($"  DFC R1 Check (rec3): '{ws2.Cells[5, dfcCol+3].Text}'");
    Console.WriteLine($"  DFC R6 Name  (rec3): '{ws2.Cells[5, dfcCol+16].Text}'");
    Console.WriteLine($"  DFC R6 Data  (rec3): '{ws2.Cells[5, dfcCol+17].Text}'");
    Console.WriteLine($"  DFC_Standard (rec1, no DFC): '{ws2.Cells[3, dfcCol].Text}'");

    // Col offsets from dfcStdCol:
    //  +0  = DFC_Standard
    //  +1  = DFC_R1_Name, +2 = DFC_R1_Data, +3 = DFC_R1_Check
    //  ...
    //  +13 = DFC_R5_Name (AI:BatchLot), +14 = DFC_R5_Data, +15 = DFC_R5_Check
    //  +16 = DFC_R6_Name (BatchLot),    +17 = DFC_R6_Data, +18 = DFC_R6_Check
    bool dfcPass = ws2.Cells[5, dfcCol].Text.Contains("GS1 Application Data Format") &&
                   ws2.Cells[5, dfcCol].Text.Contains("PASS") &&
                   ws2.Cells[5, dfcCol+1].Text == "GS1 Header" &&
                   ws2.Cells[5, dfcCol+2].Text == "<F1>" &&
                   ws2.Cells[5, dfcCol+3].Text == "PASS" &&
                   ws2.Cells[5, dfcCol+13].Text == "AI:BatchLot" &&
                   ws2.Cells[5, dfcCol+16].Text == "BatchLot" &&
                   ws2.Cells[5, dfcCol+17].Text == "GRADE-4-A-AI-INC" &&
                   string.IsNullOrEmpty(ws2.Cells[3, dfcCol].Text);
    Console.WriteLine($"  DFC verification: {(dfcPass ? "PASS" : "FAIL")}");
}

// ═══════════════════════════════════════════════════════════════════════════════
// Task 3 — 1D ISO 15416 Excel Engine
// Test: UPC-A Master + EAN-13 Master from OMNI Wide Angle PDFs, plus mixed file
// ═══════════════════════════════════════════════════════════════════════════════

Console.WriteLine("\n─────────────────────────────────────────────────────────");
Console.WriteLine("Task 3: 1D ISO 15416 Engine (UPC-A + EAN-13 + mixed file)");
Console.WriteLine("─────────────────────────────────────────────────────────");

var grade4_1D = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS");

// ── Element width data helpers ────────────────────────────────────────────────
// UPC/EAN column layout:
//   CHAR | SPACE | BAR | SPACE (left half) | SPACE | BAR | SPACE (right half)
// i.e. 6 value columns after the row-label column
var ew1DHeaders = new[] { "SPACE", "BAR", "SPACE", "SPACE", "BAR", "SPACE" };

// UPC-A Master element sizes (from OMNI_Wide_Angle_UPC-A_Master PDF)
var upcaEwSizes = new ExcelEngine.Models.ElementWidthData
{
    RecordLabel  = "UPC-A | 012345678905 | OMNI Wide Angle | 2023-06-21",
    ColumnHeaders = ew1DHeaders,
    // 6 values per row, matching ew1DHeaders: SPACE | BAR | SPACE | SPACE | BAR | SPACE
    ElementSizes =
    [
        new() { ElementName = "LGB",  Values = [null, 14m,  null, null, 12m,  null]  },
        new() { ElementName = "0",    Values = [38m,  27m,  null, null, 12m,  null]  },
        new() { ElementName = "1",    Values = [26m,  26m,  null, null, 26m,  null]  },
        new() { ElementName = "2",    Values = [26m,  13m,  null, null, 26m,  null]  },
        new() { ElementName = "3",    Values = [12m,  52m,  null, null, 12m,  null]  },
        new() { ElementName = "4",    Values = [12m,  14m,  null, null, 38m,  null]  },
        new() { ElementName = "5",    Values = [12m,  27m,  null, null, 38m,  null]  },
        new() { ElementName = "CGB",  Values = [12m,  14m,  null, 12m,  14m,  12m]  },
        new() { ElementName = "6",    Values = [null, 14m,  null, null, 12m,  null]  },
        new() { ElementName = "7",    Values = [null, 13m,  null, null, 38m,  null]  },
        new() { ElementName = "8",    Values = [null, 13m,  null, null, 26m,  null]  },
        new() { ElementName = "9",    Values = [null, 40m,  null, null, 12m,  null]  },
        new() { ElementName = "0",    Values = [null, 40m,  null, null, 25m,  null]  },
        new() { ElementName = "5",    Values = [null, 14m,  null, null, 25m,  null]  },
        new() { ElementName = "RGB",  Values = [null, 14m,  null, null, 12m,  null]  },
    ],
    ElementDeviations =
    [
        new() { ElementName = "LGB",  Values = [null,  0.9m, null, null, -0.7m, null]  },
        new() { ElementName = "0",    Values = [-1.1m, 0.9m, null, null, -0.6m, null]  },
        new() { ElementName = "1",    Values = [-0.8m, 1.1m, null, null, -1.0m, null]  },
        new() { ElementName = "2",    Values = [-0.9m, 1.4m, null, null, -1.0m, null]  },
        new() { ElementName = "3",    Values = [-0.5m, 0.5m, null, null, -0.5m, null]  },
        new() { ElementName = "4",    Values = [-0.7m, 1.1m, null, null, -1.2m, null]  },
        new() { ElementName = "5",    Values = [-0.6m, 0.9m, null, null, -1.2m, null]  },
        new() { ElementName = "CGB",  Values = [-0.8m, 0.9m, null, -0.7m, 0.7m, -0.6m] },
        new() { ElementName = "6",    Values = [null,  0.8m, null, null, -0.7m, null]  },
        new() { ElementName = "7",    Values = [null,  1.4m, null, null, -1.3m, null]  },
        new() { ElementName = "8",    Values = [null,  1.3m, null, null, -1.1m, null]  },
        new() { ElementName = "9",    Values = [null,  0.8m, null, null, -0.5m, null]  },
        new() { ElementName = "0",    Values = [null,  0.9m, null, null, -1.0m, null]  },
        new() { ElementName = "5",    Values = [null,  1.0m, null, null, -1.0m, null]  },
        new() { ElementName = "RGB",  Values = [null,  0.6m, null, null, -0.6m, null]  },
    ],
};

// EAN-13 Master element sizes (from OMNI_Wide_Angle_EAN-13_Master PDF)
var ean13EwSizes = new ExcelEngine.Models.ElementWidthData
{
    RecordLabel  = "EAN-13 | 5012345678900 | OMNI Wide Angle | 2023-06-21",
    ColumnHeaders = ew1DHeaders,
    // 6 values per row, matching ew1DHeaders: SPACE | BAR | SPACE | SPACE | BAR | SPACE
    ElementSizes =
    [
        new() { ElementName = "LGB",  Values = [null, 14m,  null, null, 12m,  null]  },
        new() { ElementName = "0",    Values = [37m,  27m,  null, null, 12m,  null]  },
        new() { ElementName = "1",    Values = [11m,  28m,  null, null, 24m,  null]  },
        new() { ElementName = "2",    Values = [23m,  28m,  null, null, 11m,  null]  },
        new() { ElementName = "3",    Values = [12m,  52m,  null, null, 12m,  null]  },
        new() { ElementName = "4",    Values = [12m,  14m,  null, null, 37m,  null]  },
        new() { ElementName = "5",    Values = [13m,  39m,  null, null, 25m,  null]  },
        new() { ElementName = "CGB",  Values = [12m,  14m,  null, 12m,  14m,  12m]  },
        new() { ElementName = "6",    Values = [null, 14m,  null, null, 12m,  null]  },
        new() { ElementName = "7",    Values = [null, 13m,  null, null, 39m,  null]  },
        new() { ElementName = "8",    Values = [null, 13m,  null, null, 26m,  null]  },
        new() { ElementName = "9",    Values = [null, 40m,  null, null, 12m,  null]  },
        new() { ElementName = "0",    Values = [null, 40m,  null, null, 25m,  null]  },
        new() { ElementName = "0",    Values = [null, 40m,  null, null, 25m,  null]  },
        new() { ElementName = "RGB",  Values = [null, 14m,  null, null, 12m,  null]  },
    ],
    ElementDeviations =
    [
        new() { ElementName = "LGB",  Values = [null,  0.9m, null, null, -0.7m, null]  },
        new() { ElementName = "0",    Values = [-1.3m, 1.2m, null, null, -0.5m, null]  },
        new() { ElementName = "1",    Values = [-0.9m, 1.1m, null, null, -1.2m, null]  },
        new() { ElementName = "2",    Values = [-1.3m, 1.3m, null, null, -0.5m, null]  },
        new() { ElementName = "3",    Values = [-0.7m, 0.7m, null, null, -0.5m, null]  },
        new() { ElementName = "4",    Values = [-0.9m, 1.1m, null, null, -1.1m, null]  },
        new() { ElementName = "5",    Values = [-0.4m, 0.6m, null, null, -0.9m, null]  },
        new() { ElementName = "CGB",  Values = [-0.6m, 0.6m, null, -0.6m, 0.7m, -0.6m] },
        new() { ElementName = "6",    Values = [null,  0.5m, null, null, -0.9m, null]  },
        new() { ElementName = "7",    Values = [null,  1.3m, null, null, -0.8m, null]  },
        new() { ElementName = "8",    Values = [null,  1.1m, null, null, -0.8m, null]  },
        new() { ElementName = "9",    Values = [null,  1.0m, null, null, -0.4m, null]  },
        new() { ElementName = "0",    Values = [null,  1.2m, null, null, -1.0m, null]  },
        new() { ElementName = "0",    Values = [null,  1.0m, null, null, -0.9m, null]  },
        new() { ElementName = "RGB",  Values = [null,  0.7m, null, null, -0.6m, null]  },
    ],
};

// ── Test Record 1D-1: UPC-A Master (all A/4.0) ────────────────────────────────
// Source: OMNI_Wide_Angle_UPC-A_Master_23-06-21_15_15_01 PDF
var upcaSession = new SessionState
{
    JobName      = "OMNI Wide Angle",
    OperatorId   = "GW4",
    RollNumber   = 1,
    CompanyName  = "Product Identification and Processing Systems, Inc.",
    DeviceSerial = "TC-833-0620-018",
    DeviceName   = "TruCheck USB",
    FirmwareVersion = "3.03.66",
    OutputDirectory = Path.Combine(Path.GetTempPath(), "vtccp_test_output"),
    SessionStarted = new DateTime(2023, 6, 21),
};

var upcaRecord = new VerificationRecord
{
    VerificationDateTime = new DateTime(2023, 6, 21, 15, 15, 1),
    Symbology       = "UPCA",
    SymbologyFamily = SymbologyFamily.Linear1D,
    DecodedData     = "012345678905",
    FormalGrade     = "4.0/06/660",
    OverallGrade    = grade4_1D,
    CustomPassFail  = OverallPassFail.Pass,
    OperatorId      = upcaSession.OperatorId,
    JobName         = upcaSession.JobName,
    BatchNumber     = "UPC-A Master",
    CompanyName     = upcaSession.CompanyName,
    ProductName     = "Elmer Chocolate, Post-calibration Conformance Challenge",
    DeviceSerial    = upcaSession.DeviceSerial,
    DeviceName      = upcaSession.DeviceName,
    FirmwareVersion = upcaSession.FirmwareVersion,
    CalibrationDate = new DateTime(2023, 6, 21, 15, 9, 56),
    Aperture        = 6,
    Wavelength      = 660,
    Standard        = "ANSI/ISO",

    // SymbolAnsiGrade — the overall grade for 1D is A/4.0
    SymbolAnsiGrade = grade4_1D,

    // Per-scan results: 10 scans, all 4.0 for all parameters
    ScanResults = Enumerable.Range(1, 10).Select(i => new ScanResult1D
    {
        ScanNumber   = i,
        Edge         = 4.0m,
        Reflectance  = "87/3",
        SC           = 4.0m,
        MinEC        = 4.0m,
        MOD          = 4.0m,
        Defect       = 4.0m,
        DCOD         = "10/10",
        DEC          = 4.0m,
        LQZ          = 4.0m,
        RQZ          = 4.0m,
        HQZ          = 4.0m,
        PerScanGrade = grade4_1D,
    }).ToList(),

    // Summary averages (from PDF)
    Avg_Edge    = 59m,
    Avg_RlRd    = "87/3",
    Avg_SC      = 84m,
    Avg_MinEC   = 71m,
    Avg_MOD     = 84m,
    Avg_Defect  = 0m,
    Avg_DCOD    = "10/10",
    Avg_DEC     = 82m,
    Avg_LQZ     = 10m,
    Avg_RQZ     = 11m,
    Avg_HQZ     = null,   // UPC-A has no horizontal quiet zone measurement
    Avg_MinQZ   = 10m,    // min(10, 11) = 10

    // General Characteristics
    BWG_Percent           = 8m,
    BWG_Mil               = 1.0m,
    Magnification         = 102m,
    NominalXDim_1D        = 13.2m,
    InspectionZoneHeight  = 293m,
    DecodableSymbolHeight = 369.7m,

    // GS1 DFC
    DataFormatCheck = new ExcelEngine.Models.DataFormatCheckResult
    {
        Overall  = OverallPassFail.Pass,
        Standard = "GS1 Application Data Format",
        Rows =
        [
            new() { Name = "GTIN",      Data = "01234567890",  Check = "PASS" },
            new() { Name = "Chk Digit", Data = "5",            Check = "PASS" },
        ],
    },

    // Element widths
    ElementWidths = upcaEwSizes,
};

// ── Test Record 1D-2: EAN-13 Master (all A/4.0) ───────────────────────────────
// Source: OMNI_Wide_Angle_EAN-13_Master_23-06-21_15_13_47 PDF
var ean13Record = new VerificationRecord
{
    VerificationDateTime = new DateTime(2023, 6, 21, 15, 13, 47),
    Symbology       = "EAN13",
    SymbologyFamily = SymbologyFamily.Linear1D,
    DecodedData     = "5012345678900",
    FormalGrade     = "4.0/06/660",
    OverallGrade    = grade4_1D,
    CustomPassFail  = OverallPassFail.Pass,
    OperatorId      = upcaSession.OperatorId,
    JobName         = upcaSession.JobName,
    BatchNumber     = "EAN-13 Master",
    CompanyName     = upcaSession.CompanyName,
    ProductName     = "Elmer Chocolate, Post-calibration Conformance Challenge",
    DeviceSerial    = upcaSession.DeviceSerial,
    DeviceName      = upcaSession.DeviceName,
    FirmwareVersion = upcaSession.FirmwareVersion,
    CalibrationDate = new DateTime(2023, 6, 21, 15, 9, 56),
    Aperture        = 6,
    Wavelength      = 660,
    Standard        = "ANSI/ISO",

    SymbolAnsiGrade = grade4_1D,

    ScanResults = Enumerable.Range(1, 10).Select(i => new ScanResult1D
    {
        ScanNumber   = i,
        Edge         = 4.0m,
        Reflectance  = "86/3",
        SC           = 4.0m,
        MinEC        = 4.0m,
        MOD          = 4.0m,
        Defect       = 4.0m,
        DCOD         = "10/10",
        DEC          = 4.0m,
        LQZ          = 4.0m,
        RQZ          = 4.0m,
        HQZ          = 4.0m,
        PerScanGrade = grade4_1D,
    }).ToList(),

    Avg_Edge    = 59m,
    Avg_RlRd    = "86/3",
    Avg_SC      = 84m,
    Avg_MinEC   = 69m,
    Avg_MOD     = 83m,
    Avg_Defect  = 0m,
    Avg_DCOD    = "10/10",
    Avg_DEC     = 84m,
    Avg_LQZ     = 8m,
    Avg_RQZ     = 9m,
    Avg_HQZ     = null,   // EAN-13 has no horizontal quiet zone measurement
    Avg_MinQZ   = 8m,     // min(8, 9) = 8

    BWG_Percent           = 8m,
    BWG_Mil               = 1.0m,
    Magnification         = 102m,
    NominalXDim_1D        = 13.4m,
    InspectionZoneHeight  = 281m,
    DecodableSymbolHeight = 352.5m,

    DataFormatCheck = new ExcelEngine.Models.DataFormatCheckResult
    {
        Overall  = OverallPassFail.Pass,
        Standard = "GS1 Application Data Format",
        Rows =
        [
            new() { Name = "GTIN",      Data = "501234567890",  Check = "PASS" },
            new() { Name = "Chk Digit", Data = "0",             Check = "PASS" },
        ],
    },

    ElementWidths = ean13EwSizes,
};

var records1D = new[] { upcaRecord, ean13Record };

// ── Write 1D-only XLSX ────────────────────────────────────────────────────────
var xlsx1DPath = Path.Combine(session.OutputDirectory!,
    ExcelFileManager.GenerateFileName(upcaSession, OutputFormat.Xlsx));
// Delete before writing so each test run is deterministic (prevents append row accumulation).
if (File.Exists(xlsx1DPath)) File.Delete(xlsx1DPath);
Console.WriteLine($"\nWriting 1D XLSX: {xlsx1DPath}");

using (var adapter = new XlsxAdapter())
using (var writer  = new ExcelWriter(adapter, schema, upcaSession))
{
    writer.Open(xlsx1DPath);
    foreach (var rec in records1D)
        writer.AppendRecord(rec);
    writer.Save();
}
var info1D = new FileInfo(xlsx1DPath);
Console.WriteLine($"  Done. Size: {info1D.Length:N0} bytes, rows: {records1D.Length} data rows");

// ── Write mixed-symbology XLSX (2D records + 1D records in same file) ─────────
var mixedPath = Path.Combine(session.OutputDirectory!, "Mixed_2D_and_1D_2025-08-11.xlsx");
// Delete before writing so each test run is deterministic (prevents append row accumulation).
if (File.Exists(mixedPath)) File.Delete(mixedPath);
Console.WriteLine($"\nWriting mixed XLSX: {mixedPath}");

using (var adapter = new XlsxAdapter())
using (var writer  = new ExcelWriter(adapter, schema, session))
{
    writer.Open(mixedPath);
    writer.AppendRecord(record1);   // 2D record
    writer.AppendRecord(record2);   // 2D record
    writer.AppendRecord(upcaRecord); // 1D record
    writer.AppendRecord(ean13Record);// 1D record
    writer.Save();
}
var mixedInfo = new FileInfo(mixedPath);
Console.WriteLine($"  Done. Size: {mixedInfo.Length:N0} bytes, rows: 4 data rows (2 x 2D + 2 x 1D)");

// ── Verify 1D XLSX structure ──────────────────────────────────────────────────
Console.WriteLine("\nVerifying 1D XLSX structure...");
OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
using (var pkg3 = new OfficeOpenXml.ExcelPackage(new FileInfo(xlsx1DPath)))
{
    var wsMain = pkg3.Workbook.Worksheets["Main"];
    var wsEW   = pkg3.Workbook.Worksheets["Element Widths"];

    int mainRows = wsMain?.Dimension?.Rows ?? 0;
    int mainCols = wsMain?.Dimension?.Columns ?? 0;
    Console.WriteLine($"  Main sheet: {mainRows} rows x {mainCols} columns");
    // mainCols may be > schema.Columns.Count now that SchemaVersionWriter adds cells in row 1
    // past the last data column; those extra cells are intentional metadata and not data columns.
    Console.WriteLine($"  Columns >= schema: {(mainCols >= schema.Columns.Count ? "YES" : $"NO ({mainCols} vs {schema.Columns.Count})")}");

    // Row 3 = UPC-A record
    Console.WriteLine($"  Row 3 Symbology: '{wsMain?.Cells[3, 9].Text}'");
    Console.WriteLine($"  Row 3 Data:      '{wsMain?.Cells[3, 10].Text}'");
    Console.WriteLine($"  Row 3 FormalGrade: '{wsMain?.Cells[3, 11].Text}'");

    // Find 1D summary columns by header labels
    int avgEdgeCol = -1, avgScCol = -1, bwgPctCol = -1, ratioCol = -1;
    for (int c = 1; c <= mainCols; c++)
    {
        var hdr = wsMain?.Cells[2, c].Text ?? "";
        if (hdr == "Edge")   avgEdgeCol = c;
        if (hdr == "SC")     avgScCol   = c;
        if (hdr == "BWG%")   bwgPctCol  = c;
        if (hdr == "Ratio")  ratioCol   = c;
    }
    Console.WriteLine($"  Avg_Edge col: {avgEdgeCol}, value (row 3): '{wsMain?.Cells[3, avgEdgeCol].Text}'");
    Console.WriteLine($"  Avg_SC col:   {avgScCol},   value (row 3): '{wsMain?.Cells[3, avgScCol].Text}'");
    Console.WriteLine($"  BWG% col:     {bwgPctCol},  value (row 3): '{wsMain?.Cells[3, bwgPctCol].Text}'");
    Console.WriteLine($"  Ratio col:    {ratioCol},   value (row 3): '{wsMain?.Cells[3, ratioCol].Text}' (should be blank for UPC-A)");

    // Verify 2D columns are blank for 1D record
    int uecPctCol = -1;
    for (int c = 1; c <= mainCols; c++)
        if ((wsMain?.Cells[2, c].Text ?? "") == "UEC%") { uecPctCol = c; break; }
    string uecVal = wsMain?.Cells[3, uecPctCol].Text ?? "";
    Console.WriteLine($"  UEC% (2D col, should be blank for 1D): '{uecVal}'");

    // Element Widths sheet
    bool ewSheetExists = wsEW is not null;
    int ewRows = wsEW?.Dimension?.Rows ?? 0;
    Console.WriteLine($"\n  Element Widths sheet exists: {ewSheetExists}");
    Console.WriteLine($"  Element Widths rows: {ewRows}");
    if (wsEW is not null && ewRows > 0)
    {
        Console.WriteLine($"  EW row 1 (label): '{wsEW.Cells[1, 1].Text}'");
        Console.WriteLine($"  EW row 2 (section): '{wsEW.Cells[2, 1].Text}'");
        Console.WriteLine($"  EW row 3 (col hdr): '{wsEW.Cells[3, 1].Text}' | '{wsEW.Cells[3, 2].Text}'");
        Console.WriteLine($"  EW row 4 (LGB): '{wsEW.Cells[4, 1].Text}' | BAR={wsEW.Cells[4, 3].Text}");
    }

    // ── Per-scan sub-table verification ───────────────────────────────────────
    // Layout: row 3 = UPC-A summary, rows 4..13 = 10 scan sub-rows, row 14 = EAN-13 summary
    int scanLabelCol = 1;   // "Scan N" written to Date column (col 1)

    // Find LQZ/RQZ/HQZ/MinQZ column positions from header row
    int avgLqzCol = -1, avgRqzCol = -1, avgHqzCol = -1, avgMinQzCol = -1;
    for (int c = 1; c <= mainCols; c++)
    {
        var hdr = wsMain?.Cells[2, c].Text ?? "";
        if (hdr == "QZ-L")  avgLqzCol   = c;
        if (hdr == "QZ-R")  avgRqzCol   = c;
        if (hdr == "QZ-H")  avgHqzCol   = c;
        if (hdr == "MinQZ") avgMinQzCol = c;
    }
    Console.WriteLine($"\n  Per-scan sub-table verification:");
    Console.WriteLine($"  LQZ col: {avgLqzCol}, RQZ col: {avgRqzCol}, HQZ col: {avgHqzCol}, MinQZ col: {avgMinQzCol}");
    Console.WriteLine($"  Row 3 UPC-A summary: LQZ='{wsMain?.Cells[3, avgLqzCol].Text}' RQZ='{wsMain?.Cells[3, avgRqzCol].Text}' MinQZ='{wsMain?.Cells[3, avgMinQzCol].Text}'");
    Console.WriteLine($"  Row 4 col 1 (Scan 1 label):  '{wsMain?.Cells[4, scanLabelCol].Text}'");
    Console.WriteLine($"  Row 4 Edge value (col {avgEdgeCol}): '{wsMain?.Cells[4, avgEdgeCol].Text}' (expect 4)");
    Console.WriteLine($"  Row 4 SC value (col {avgScCol}):   '{wsMain?.Cells[4, avgScCol].Text}' (expect 4)");
    Console.WriteLine($"  Row 4 LQZ (col {avgLqzCol}): '{wsMain?.Cells[4, avgLqzCol].Text}' (expect 4)");
    Console.WriteLine($"  Row 4 RQZ (col {avgRqzCol}): '{wsMain?.Cells[4, avgRqzCol].Text}' (expect 4)");
    Console.WriteLine($"  Row 4 HQZ (col {avgHqzCol}): '{wsMain?.Cells[4, avgHqzCol].Text}' (expect 4)");
    Console.WriteLine($"  Row 4 MinQZ (col {avgMinQzCol}): '{wsMain?.Cells[4, avgMinQzCol].Text}' (expect 4 = min(4,4,4))");
    Console.WriteLine($"  Row 13 col 1 (Scan 10 label): '{wsMain?.Cells[13, scanLabelCol].Text}'");
    Console.WriteLine($"  Row 14 Symbology (EAN-13):    '{wsMain?.Cells[14, 9].Text}' (expect EAN13)");
    Console.WriteLine($"  Row 14 Data:                  '{wsMain?.Cells[14, 10].Text}'");
    // UPC-A summary row: LQZ=10, RQZ=11, MinQZ=10
    Console.WriteLine($"  UPC-A summary Avg_LQZ: '{wsMain?.Cells[3, avgLqzCol].Text}' (expect 10)");
    Console.WriteLine($"  UPC-A summary Avg_RQZ: '{wsMain?.Cells[3, avgRqzCol].Text}' (expect 11)");
    Console.WriteLine($"  UPC-A summary Avg_MinQZ: '{wsMain?.Cells[3, avgMinQzCol].Text}' (expect 10)");

    // 1D verification assertions
    bool scan1Label    = wsMain?.Cells[4, scanLabelCol].Text == "Scan 1";
    bool scan1Edge     = wsMain?.Cells[4, avgEdgeCol].Text == "4";
    bool scan1Lqz      = avgLqzCol > 0 && wsMain?.Cells[4, avgLqzCol].Text == "4";
    bool scan1Rqz      = avgRqzCol > 0 && wsMain?.Cells[4, avgRqzCol].Text == "4";
    bool scan1Hqz      = avgHqzCol > 0 && wsMain?.Cells[4, avgHqzCol].Text == "4";
    bool scan1MinQz    = avgMinQzCol > 0 && wsMain?.Cells[4, avgMinQzCol].Text == "4";
    bool scan10Label   = wsMain?.Cells[13, scanLabelCol].Text == "Scan 10";
    bool ean13Row      = wsMain?.Cells[14, 9].Text == "EAN13";
    bool upcaAvgLqz    = avgLqzCol > 0 && wsMain?.Cells[3, avgLqzCol].Text == "10";
    bool upcaAvgRqz    = avgRqzCol > 0 && wsMain?.Cells[3, avgRqzCol].Text == "11";
    bool upcaAvgMinQz  = avgMinQzCol > 0 && wsMain?.Cells[3, avgMinQzCol].Text == "10";

    bool oneDPass =
        mainCols >= schema.Columns.Count &&
        wsMain!.Cells[3, 9].Text == "UPCA" &&
        wsMain.Cells[3, 10].Text == "012345678905" &&
        avgEdgeCol > 0 && wsMain.Cells[3, avgEdgeCol].Text == "59" &&
        avgScCol   > 0 && wsMain.Cells[3, avgScCol].Text   == "84" &&
        bwgPctCol  > 0 && wsMain.Cells[3, bwgPctCol].Text  == "8" &&
        string.IsNullOrEmpty(uecVal) &&
        scan1Label && scan1Edge &&
        scan1Lqz && scan1Rqz && scan1Hqz && scan1MinQz &&
        scan10Label && ean13Row &&
        upcaAvgLqz && upcaAvgRqz && upcaAvgMinQz &&
        ewSheetExists &&
        wsEW!.Cells[4, 1].Text == "LGB";  // first data row element name

    Console.WriteLine($"\n  1D verification: {(oneDPass ? "PASS" : "FAIL")}");
    if (!scan1Label)   Console.WriteLine($"    FAIL: scan 1 label = '{wsMain?.Cells[4, scanLabelCol].Text}'");
    if (!scan1Edge)    Console.WriteLine($"    FAIL: scan 1 edge  = '{wsMain?.Cells[4, avgEdgeCol].Text}'");
    if (!scan1Lqz)     Console.WriteLine($"    FAIL: scan 1 LQZ   = '{wsMain?.Cells[4, avgLqzCol].Text}'");
    if (!scan1Rqz)     Console.WriteLine($"    FAIL: scan 1 RQZ   = '{wsMain?.Cells[4, avgRqzCol].Text}'");
    if (!scan1Hqz)     Console.WriteLine($"    FAIL: scan 1 HQZ   = '{wsMain?.Cells[4, avgHqzCol].Text}'");
    if (!scan1MinQz)   Console.WriteLine($"    FAIL: scan 1 MinQZ = '{wsMain?.Cells[4, avgMinQzCol].Text}'");
    if (!scan10Label)  Console.WriteLine($"    FAIL: scan 10 label= '{wsMain?.Cells[13, scanLabelCol].Text}'");
    if (!ean13Row)     Console.WriteLine($"    FAIL: EAN-13 at row 14 = '{wsMain?.Cells[14, 9].Text}'");
    if (!upcaAvgLqz)   Console.WriteLine($"    FAIL: UPC-A Avg_LQZ  = '{wsMain?.Cells[3, avgLqzCol].Text}'");
    if (!upcaAvgRqz)   Console.WriteLine($"    FAIL: UPC-A Avg_RQZ  = '{wsMain?.Cells[3, avgRqzCol].Text}'");
    if (!upcaAvgMinQz) Console.WriteLine($"    FAIL: UPC-A Avg_MinQZ= '{wsMain?.Cells[3, avgMinQzCol].Text}'");

    // Mixed-symbology file check
    // Row layout with scan sub-rows:
    //   Row 1 = Title, Row 2 = Header
    //   Row 3 = 2D record 1 (GS1 DataMatrix)
    //   Row 4 = 2D record 2 (GS1 DataMatrix)
    //   Row 5 = UPC-A summary, Rows 6-15 = UPC-A scans (10)
    //   Row 16 = EAN-13 summary, Rows 17-26 = EAN-13 scans (10)
    using var pkgMixed = new OfficeOpenXml.ExcelPackage(new FileInfo(mixedPath));
    var wsMixed = pkgMixed.Workbook.Worksheets["Main"];
    int mixedRows = wsMixed?.Dimension?.Rows ?? 0;
    int mixedCols = wsMixed?.Dimension?.Columns ?? 0;
    Console.WriteLine($"\n  Mixed file Main: {mixedRows} rows x {mixedCols} columns");
    Console.WriteLine($"  Row 3 (2D DM rec1):  '{wsMixed?.Cells[3, 9].Text}' (expect 'GS1 DataMatrix')");
    Console.WriteLine($"  Row 4 (2D DM rec2):  '{wsMixed?.Cells[4, 9].Text}' (expect 'GS1 DataMatrix')");
    Console.WriteLine($"  Row 5 (1D UPC-A):    '{wsMixed?.Cells[5, 9].Text}' (expect 'UPCA')");
    Console.WriteLine($"  Row 6 (UPC-A scan1): '{wsMixed?.Cells[6, scanLabelCol].Text}' (expect 'Scan 1')");
    Console.WriteLine($"  Row 16 (1D EAN-13):  '{wsMixed?.Cells[16, 9].Text}' (expect 'EAN13')");

    bool mixedPass =
        mixedRows == 26 &&  // title + header + 2 2D rows + 2 UPC-A (summary+10scans) + 2 EAN-13 (summary+10scans)
        wsMixed!.Cells[3, 9].Text  == "GS1 DataMatrix" &&
        wsMixed.Cells[4, 9].Text   == "GS1 DataMatrix" &&
        wsMixed.Cells[5, 9].Text   == "UPCA" &&
        wsMixed.Cells[6, scanLabelCol].Text == "Scan 1" &&
        wsMixed.Cells[16, 9].Text  == "EAN13";

    Console.WriteLine($"  Mixed-symbology verification: {(mixedPass ? "PASS" : "FAIL")}");
}

Console.WriteLine("\nTask 3 complete.");

// ═══════════════════════════════════════════════════════════════════════════════
// Task 4 — Session management & test harness
// Tests: SessionManager lifecycle (start / add 6 records / close),
//        JSON sidecar written + cleaned up, SchemaVersion header in row 1,
//        filename sanitization, and a 6-record full harness run.
// ═══════════════════════════════════════════════════════════════════════════════

Console.WriteLine("\n─────────────────────────────────────────────────────────");
Console.WriteLine("Task 4: Session management & full 6-record harness");
Console.WriteLine("─────────────────────────────────────────────────────────");

var t4Session = new SessionState
{
    JobName         = "T4 Harness & Co.",   // contains '.' — tests sanitization
    OperatorId      = "GW4",
    RollNumber      = 1,
    BatchNumber     = "BATCH-T4",
    CompanyName     = "Product Identification and Processing Systems, Inc.",
    ProductName     = "Phase 1 Integration Test",
    DeviceSerial    = "1A2202PP003652",
    DeviceName      = "DM475-8E9364",
    FirmwareVersion = "6.1.16_sr3",
    CalibrationDate = new DateTime(2025, 8, 11, 16, 50, 0),
    OutputDirectory = Path.Combine(Path.GetTempPath(), "vtccp_t4_output"),
    OutputFormat    = OutputFormat.Xlsx,
    SessionStarted  = new DateTime(2026, 1, 15),
};

// ── Filename sanitization check ───────────────────────────────────────────────
// Uses explicit character set (not Path.GetInvalidFileNameChars) for platform-
// independent determinism.  '&' is legal in Windows filenames and preserved.
Console.WriteLine("\nFilename sanitization tests:");
Console.WriteLine($"  'My Job & Co.' -> '{ExcelFileManager.SanitizeFileName("My Job & Co.")}'   (& preserved — legal in filenames)");
Console.WriteLine($"  'Roll/3'       -> '{ExcelFileManager.SanitizeFileName("Roll/3")}'");
Console.WriteLine($"  'A B C'        -> '{ExcelFileManager.SanitizeFileName("A B C")}'");
Console.WriteLine($"  'x[1].xlsx'    -> '{ExcelFileManager.SanitizeFileName("x[1].xlsx")}'");
Console.WriteLine($"  'a*b?c'        -> '{ExcelFileManager.SanitizeFileName("a*b?c")}'");

// Verify specific sanitization expectations
bool sanSlash     = ExcelFileManager.SanitizeFileName("Roll/3")    == "Roll_3";
bool sanSpace     = ExcelFileManager.SanitizeFileName("A B C")     == "A_B_C";
bool sanBracket   = ExcelFileManager.SanitizeFileName("x[1].xlsx") == "x_1_.xlsx";
bool sanGlob      = ExcelFileManager.SanitizeFileName("a*b?c")     == "a_b_c";
bool sanAmpersand = ExcelFileManager.SanitizeFileName("My & Co.")  == "My_&_Co."; // & is legal
Console.WriteLine($"  Sanitization assertions: slash={sanSlash} space={sanSpace} bracket={sanBracket} glob={sanGlob} ampersand={sanAmpersand}");

// Default filename for T4 session (uses JobName + date)
Console.WriteLine($"  T4 session file (default): {ExcelFileManager.GenerateFileName(t4Session, OutputFormat.Xlsx)}");

// Custom filename pattern test
var patternSession = new SessionState
{
    JobName = "CalCard Prod", OperatorId = "GW4", RollNumber = 3,
    SessionStarted = new DateTime(2026, 1, 15),
    FileNamePattern = "{Job}_{Op}_Roll{Roll}_{Date}",
    OutputDirectory = t4Session.OutputDirectory,
    OutputFormat = OutputFormat.Xlsx,
};
string customFileName = ExcelFileManager.GenerateFileName(patternSession, OutputFormat.Xlsx);
bool patternOk = customFileName == "CalCard_Prod_GW4_Roll3_2026-01-15.xlsx";
Console.WriteLine($"  Custom pattern file: '{customFileName}' — {(patternOk ? "PASS" : "FAIL")}");

// ── Build the 6 test records ──────────────────────────────────────────────────
// 3 × GS1 DataMatrix  (records 1–3)
// 1 × plain DataMatrix
// 1 × UPC-A (1D)
// 1 × EAN-13 (1D)

var t4Pass4A = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS");
var t4Fail1D = GradingResult.FromLetterAndNumeric("D", 1.0m, "FAIL");

// Record T4-1: GS1 DataMatrix grade A
var t4Rec1 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 1, 15, 9, 0, 0),
    Symbology       = "GS1 DataMatrix",
    SymbologyFamily = SymbologyFamily.GS1DataMatrix,
    DecodedData     = "<F1>0101234567890128",
    FormalGrade     = "4.0/16/660/45Q",
    OverallGrade    = t4Pass4A, CustomPassFail = OverallPassFail.Pass,
    OperatorId      = t4Session.OperatorId, JobName = t4Session.JobName,
    RollNumber      = t4Session.RollNumber, CompanyName = t4Session.CompanyName,
    DeviceSerial    = t4Session.DeviceSerial, DeviceName = t4Session.DeviceName,
    FirmwareVersion = t4Session.FirmwareVersion, CalibrationDate = t4Session.CalibrationDate,
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    MatrixSize = "22x22", HorizontalBWG = -8, VerticalBWG = -8,
    EncodedCharacters = 15, TotalCodewords = 50, DataCodewords = 30,
    ErrorCorrectionBudget = 20, ErrorsCorrected = 0, ErrorCapacityUsed = 0,
    ErrorCorrectionType = "ECC 200", ImagePolarity = ImagePolarity.WhiteOnBlack,
    NominalXDim_2D = 19.8m, PixelsPerModule = 33.0m,
    UEC_Percent = 100, UEC_Grade = t4Pass4A, SC_Percent = 90, SC_Grade = t4Pass4A,
    MOD_Grade = t4Pass4A, RM_Grade = t4Pass4A, ANU_Percent = 0.1m, ANU_Grade = t4Pass4A,
    GNU_Percent = 1.5m, GNU_Grade = t4Pass4A, FPD_Grade = t4Pass4A,
    LLS_Grade = t4Pass4A, BLS_Grade = t4Pass4A, LQZ_Grade = t4Pass4A,
    BQZ_Grade = t4Pass4A, TQZ_Grade = t4Pass4A, RQZ_Grade = t4Pass4A,
    TTR_Percent = 0, TTR_Grade = t4Pass4A, RTR_Percent = 0, RTR_Grade = t4Pass4A,
    TCT_Grade = t4Pass4A, RCT_Grade = t4Pass4A, AG_Value = 4.0m, AG_Grade = t4Pass4A,
    DECODE_Grade = t4Pass4A,
    DataFormatCheck = new ExcelEngine.Models.DataFormatCheckResult
    {
        Overall = OverallPassFail.Pass, Standard = "GS1 Application Data Format",
        Rows =
        [
            new() { Name = "GS1 Header", Data = "<F1>", Check = "PASS" },
            new() { Name = "AI:GTIN",    Data = "01",   Check = "PASS" },
            new() { Name = "GTIN",       Data = "0123456789012", Check = "PASS" },
            new() { Name = "Chk Digit",  Data = "8",   Check = "PASS" },
        ],
    },
};

// Record T4-2: GS1 DataMatrix grade D (ANU fail)
var t4Rec2 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 1, 15, 9, 1, 0),
    Symbology       = "GS1 DataMatrix",
    SymbologyFamily = SymbologyFamily.GS1DataMatrix,
    DecodedData     = "<F1>0101234567890128BATCHT42",
    FormalGrade     = "1.0/17/660/45Q",
    OverallGrade    = t4Fail1D, CustomPassFail = OverallPassFail.Fail,
    OperatorId      = t4Session.OperatorId, JobName = t4Session.JobName,
    RollNumber      = t4Session.RollNumber, CompanyName = t4Session.CompanyName,
    DeviceSerial    = t4Session.DeviceSerial, DeviceName = t4Session.DeviceName,
    FirmwareVersion = t4Session.FirmwareVersion, CalibrationDate = t4Session.CalibrationDate,
    Aperture = 17, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    MatrixSize = "22x22", HorizontalBWG = 7, VerticalBWG = 7,
    EncodedCharacters = 20, TotalCodewords = 50, DataCodewords = 30,
    ErrorCorrectionBudget = 20, ErrorsCorrected = 0, ErrorCapacityUsed = 0,
    ErrorCorrectionType = "ECC 200", ImagePolarity = ImagePolarity.BlackOnWhite,
    NominalXDim_2D = 20.9m, PixelsPerModule = 35.0m,
    UEC_Percent = 100, UEC_Grade = t4Pass4A, SC_Percent = 84, SC_Grade = t4Pass4A,
    MOD_Grade = t4Pass4A, RM_Grade = t4Pass4A, ANU_Percent = 11.0m, ANU_Grade = t4Fail1D,
    GNU_Percent = 2.8m, GNU_Grade = t4Pass4A, FPD_Grade = t4Pass4A,
    LLS_Grade = t4Pass4A, BLS_Grade = t4Pass4A, LQZ_Grade = t4Pass4A,
    BQZ_Grade = t4Pass4A, TQZ_Grade = t4Pass4A, RQZ_Grade = t4Pass4A,
    TTR_Percent = 0, TTR_Grade = t4Pass4A, RTR_Percent = 0, RTR_Grade = t4Pass4A,
    TCT_Grade = t4Pass4A, RCT_Grade = t4Pass4A, AG_Value = 4.0m, AG_Grade = t4Pass4A,
    DECODE_Grade = t4Pass4A,
};

// Record T4-3: GS1 DataMatrix grade A (another device)
var t4Rec3 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 1, 15, 9, 2, 0),
    Symbology       = "GS1 DataMatrix",
    SymbologyFamily = SymbologyFamily.GS1DataMatrix,
    DecodedData     = "<F1>0101234567890128BATCHT43",
    FormalGrade     = "4.0/16/660/45Q",
    OverallGrade    = t4Pass4A, CustomPassFail = OverallPassFail.Pass,
    OperatorId      = t4Session.OperatorId, JobName = t4Session.JobName,
    RollNumber      = t4Session.RollNumber, BatchNumber = "BATCH-T4C",
    CompanyName = t4Session.CompanyName, DeviceSerial = t4Session.DeviceSerial,
    DeviceName = "DM475-866D76", FirmwareVersion = t4Session.FirmwareVersion,
    CalibrationDate = t4Session.CalibrationDate,
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    MatrixSize = "22x22", HorizontalBWG = -5, VerticalBWG = -5,
    EncodedCharacters = 20, TotalCodewords = 50, DataCodewords = 30,
    ErrorCorrectionBudget = 20, ErrorsCorrected = 0, ErrorCapacityUsed = 0,
    ErrorCorrectionType = "ECC 200", ImagePolarity = ImagePolarity.WhiteOnBlack,
    NominalXDim_2D = 19.6m, PixelsPerModule = 33.0m,
    UEC_Percent = 100, UEC_Grade = t4Pass4A, SC_Percent = 94, SC_Grade = t4Pass4A,
    MOD_Grade = t4Pass4A, RM_Grade = t4Pass4A, ANU_Percent = 0.2m, ANU_Grade = t4Pass4A,
    GNU_Percent = 2.1m, GNU_Grade = t4Pass4A, FPD_Grade = t4Pass4A,
    LLS_Grade = t4Pass4A, BLS_Grade = t4Pass4A, LQZ_Grade = t4Pass4A,
    BQZ_Grade = t4Pass4A, TQZ_Grade = t4Pass4A, RQZ_Grade = t4Pass4A,
    TTR_Percent = 0, TTR_Grade = t4Pass4A, RTR_Percent = 0, RTR_Grade = t4Pass4A,
    TCT_Grade = t4Pass4A, RCT_Grade = t4Pass4A, AG_Value = 4.0m, AG_Grade = t4Pass4A,
    DECODE_Grade = t4Pass4A,
};

// Record T4-4: plain Data Matrix (not GS1)
var t4Rec4 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 1, 15, 9, 3, 0),
    Symbology       = "DataMatrix",
    SymbologyFamily = SymbologyFamily.DataMatrix,
    DecodedData     = "PLAIN-DM-T4-TEST",
    FormalGrade     = "3.5/16/660/45Q",
    OverallGrade    = GradingResult.FromLetterAndNumeric("B", 3.5m, "PASS"),
    CustomPassFail  = OverallPassFail.Pass,
    OperatorId      = t4Session.OperatorId, JobName = t4Session.JobName,
    RollNumber      = t4Session.RollNumber, CompanyName = t4Session.CompanyName,
    DeviceSerial    = t4Session.DeviceSerial, DeviceName = t4Session.DeviceName,
    FirmwareVersion = t4Session.FirmwareVersion, CalibrationDate = t4Session.CalibrationDate,
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    MatrixSize = "16x16", HorizontalBWG = -3, VerticalBWG = -3,
    EncodedCharacters = 12, TotalCodewords = 24, DataCodewords = 12,
    ErrorCorrectionBudget = 12, ErrorsCorrected = 0, ErrorCapacityUsed = 0,
    ErrorCorrectionType = "ECC 200", ImagePolarity = ImagePolarity.BlackOnWhite,
    NominalXDim_2D = 25.0m, PixelsPerModule = 41.5m,
    UEC_Percent = 100, UEC_Grade = t4Pass4A,
    SC_Percent = 80, SC_Grade = GradingResult.FromLetterAndNumeric("B", 3.5m, "PASS"),
    MOD_Grade = t4Pass4A, RM_Grade = t4Pass4A, ANU_Percent = 0.5m, ANU_Grade = t4Pass4A,
    GNU_Percent = 1.0m, GNU_Grade = t4Pass4A, FPD_Grade = t4Pass4A,
    LLS_Grade = t4Pass4A, BLS_Grade = t4Pass4A, LQZ_Grade = t4Pass4A,
    BQZ_Grade = t4Pass4A, TQZ_Grade = t4Pass4A, RQZ_Grade = t4Pass4A,
    TTR_Percent = 0, TTR_Grade = t4Pass4A, RTR_Percent = 0, RTR_Grade = t4Pass4A,
    TCT_Grade = t4Pass4A, RCT_Grade = t4Pass4A,
    AG_Value = 3.5m, AG_Grade = GradingResult.FromLetterAndNumeric("B", 3.5m, "PASS"),
    DECODE_Grade = t4Pass4A,
};

// Record T4-5: UPC-A (1D, 5 scans)
var t4Rec5 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 1, 15, 9, 4, 0),
    Symbology       = "UPCA",
    SymbologyFamily = SymbologyFamily.Linear1D,
    DecodedData     = "012345678905",
    FormalGrade     = "4.0/06/660",
    OverallGrade    = t4Pass4A, CustomPassFail = OverallPassFail.Pass,
    OperatorId      = t4Session.OperatorId, JobName = t4Session.JobName,
    RollNumber      = t4Session.RollNumber, CompanyName = t4Session.CompanyName,
    DeviceSerial    = t4Session.DeviceSerial, DeviceName = t4Session.DeviceName,
    FirmwareVersion = t4Session.FirmwareVersion, CalibrationDate = t4Session.CalibrationDate,
    Aperture = 6, Wavelength = 660, Standard = "ANSI/ISO",
    SymbolAnsiGrade = t4Pass4A,
    ScanResults = Enumerable.Range(1, 5).Select(i => new ScanResult1D
    {
        ScanNumber = i, Edge = 4.0m, Reflectance = "87/3",
        SC = 4.0m, MinEC = 4.0m, MOD = 4.0m, Defect = 4.0m,
        DCOD = "10/10", DEC = 4.0m, LQZ = 4.0m, RQZ = 4.0m, HQZ = null,
        PerScanGrade = t4Pass4A,
    }).ToList(),
    Avg_Edge = 59m, Avg_RlRd = "87/3", Avg_SC = 84m, Avg_MinEC = 71m,
    Avg_MOD = 84m, Avg_Defect = 0m, Avg_DCOD = "10/10", Avg_DEC = 82m,
    Avg_LQZ = 10m, Avg_RQZ = 11m, Avg_HQZ = null, Avg_MinQZ = 10m,
    BWG_Percent = 8m, BWG_Mil = 1.0m, Magnification = 102m,
    NominalXDim_1D = 13.2m, InspectionZoneHeight = 293m, DecodableSymbolHeight = 369.7m,
};

// Record T4-6: EAN-13 (1D, 5 scans)
var t4Rec6 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 1, 15, 9, 5, 0),
    Symbology       = "EAN13",
    SymbologyFamily = SymbologyFamily.Linear1D,
    DecodedData     = "5012345678900",
    FormalGrade     = "4.0/06/660",
    OverallGrade    = t4Pass4A, CustomPassFail = OverallPassFail.Pass,
    OperatorId      = t4Session.OperatorId, JobName = t4Session.JobName,
    RollNumber      = t4Session.RollNumber, CompanyName = t4Session.CompanyName,
    DeviceSerial    = t4Session.DeviceSerial, DeviceName = t4Session.DeviceName,
    FirmwareVersion = t4Session.FirmwareVersion, CalibrationDate = t4Session.CalibrationDate,
    Aperture = 6, Wavelength = 660, Standard = "ANSI/ISO",
    SymbolAnsiGrade = t4Pass4A,
    ScanResults = Enumerable.Range(1, 5).Select(i => new ScanResult1D
    {
        ScanNumber = i, Edge = 4.0m, Reflectance = "86/3",
        SC = 4.0m, MinEC = 4.0m, MOD = 4.0m, Defect = 4.0m,
        DCOD = "10/10", DEC = 4.0m, LQZ = 4.0m, RQZ = 4.0m, HQZ = 4.0m,
        PerScanGrade = t4Pass4A,
    }).ToList(),
    Avg_Edge = 59m, Avg_RlRd = "86/3", Avg_SC = 84m, Avg_MinEC = 69m,
    Avg_MOD = 83m, Avg_Defect = 0m, Avg_DCOD = "10/10", Avg_DEC = 84m,
    Avg_LQZ = 8m, Avg_RQZ = 9m, Avg_HQZ = null, Avg_MinQZ = 8m,
    BWG_Percent = 8m, BWG_Mil = 1.0m, Magnification = 102m,
    NominalXDim_1D = 13.4m, InspectionZoneHeight = 281m, DecodableSymbolHeight = 352.5m,
};

var t4Records = new[] { t4Rec1, t4Rec2, t4Rec3, t4Rec4, t4Rec5, t4Rec6 };

// ── Clean up any leftover test file from previous runs ────────────────────────
var t4OutputPath = ExcelFileManager.ResolveOutputPath(t4Session, OutputFormat.Xlsx);
var t4SidecarPath = SessionManager.GetSidecarPath(t4OutputPath);
if (File.Exists(t4OutputPath)) File.Delete(t4OutputPath);
if (File.Exists(t4SidecarPath)) File.Delete(t4SidecarPath);

// ── SessionManager lifecycle test ─────────────────────────────────────────────
Console.WriteLine("\nSessionManager lifecycle:");
using var t4Mgr = new SessionManager(schema);

string openedPath = t4Mgr.StartSession(t4Session);
Console.WriteLine($"  Session opened: {openedPath}");

// Sidecar must exist immediately after StartSession.
bool sidecarCreated = File.Exists(t4SidecarPath);
Console.WriteLine($"  Sidecar created on StartSession: {sidecarCreated}");

foreach (var rec in t4Records)
    t4Mgr.AddRecord(rec);

int recordsWrittenBeforeClose = t4Mgr.RecordsWritten;
Console.WriteLine($"  Records written: {recordsWrittenBeforeClose} (expect 6)");

// Sidecar must still exist while session is open.
bool sidecarStillPresent = File.Exists(t4SidecarPath);
Console.WriteLine($"  Sidecar present while open:     {sidecarStillPresent}");

// ── Roll semantics test ───────────────────────────────────────────────────────
// Rule 1: New session — roll stays at whatever caller supplied.
int rollBefore = t4Mgr.CurrentSession!.RollNumber;
Console.WriteLine($"  Roll before SetNewOperatorAndRoll: {rollBefore} (expect 1)");

// Rule 2: Manual mode — caller must supply the new roll value explicitly.
t4Mgr.SetNewOperatorAndRoll("OP2", manualRoll: 2);
int rollAfterSet = t4Mgr.CurrentSession!.RollNumber;
string opAfterSet = t4Mgr.CurrentSession!.OperatorId ?? "";
Console.WriteLine($"  Roll after SetNewOperatorAndRoll:  {rollAfterSet} (expect 2), Operator: '{opAfterSet}'");

bool rollNewSession = rollBefore == 1;
bool rollIncrement  = rollAfterSet == 2;
bool opChanged      = opAfterSet == "OP2";
Console.WriteLine($"  Roll semantics: new-session={rollNewSession} increment={rollIncrement} op-changed={opChanged}");

t4Mgr.CloseSession();

// Sidecar must be deleted after clean close.
bool sidecarDeleted = !File.Exists(t4SidecarPath);
Console.WriteLine($"  Sidecar deleted after close:    {sidecarDeleted}");

// ── Verify output file structure ──────────────────────────────────────────────
Console.WriteLine("\nVerifying Task 4 output file...");
var t4FileInfo = new FileInfo(t4OutputPath);
Console.WriteLine($"  File size: {t4FileInfo.Length:N0} bytes");

using var t4Pkg = new OfficeOpenXml.ExcelPackage(t4FileInfo);
var t4Main = t4Pkg.Workbook.Worksheets["Main"];
int t4Rows = t4Main?.Dimension?.Rows ?? 0;
int t4Cols = t4Main?.Dimension?.Columns ?? 0;
Console.WriteLine($"  Main sheet: {t4Rows} rows x {t4Cols} columns");

// Row layout:
//   Row 1 = Title + SchemaVersion
//   Row 2 = Header
//   Rows 3-6 = 4× 2D records (T4-1..T4-4)
//   Row 7 = UPC-A summary, Rows 8..12 = 5 scan sub-rows
//   Row 13 = EAN-13 summary, Rows 14..18 = 5 scan sub-rows
// Total = 18 rows
Console.WriteLine($"  Expected 18 rows (2 hdr + 4 2D + 1+5 UPC-A + 1+5 EAN-13)");

// SchemaVersion header check (cells in title row past last schema column)
int svCol = schema.Columns.Count + 2; // startCol in SchemaVersionWriter
string svMarker  = t4Main?.Cells[1, svCol    ].Text ?? "";
string svName    = t4Main?.Cells[1, svCol + 1].Text ?? "";
string svVersion = t4Main?.Cells[1, svCol + 2].Text ?? "";
Console.WriteLine($"  SchemaVersion row 1 col {svCol}:   '{svMarker}' (expect 'VTCCP')");
Console.WriteLine($"  SchemaVersion row 1 col {svCol+1}: '{svName}'   (expect '{schema.Name}')");
Console.WriteLine($"  SchemaVersion row 1 col {svCol+2}: '{svVersion}' (expect '{schema.Version}')");

// Data rows spot-check
Console.WriteLine($"  Row 3 symbology (GS1 DM):    '{t4Main?.Cells[3,  9].Text}' (expect 'GS1 DataMatrix')");
Console.WriteLine($"  Row 4 symbology (GS1 DM):    '{t4Main?.Cells[4,  9].Text}' (expect 'GS1 DataMatrix')");
Console.WriteLine($"  Row 5 symbology (GS1 DM):    '{t4Main?.Cells[5,  9].Text}' (expect 'GS1 DataMatrix')");
Console.WriteLine($"  Row 6 symbology (DM):        '{t4Main?.Cells[6,  9].Text}' (expect 'DataMatrix')");
Console.WriteLine($"  Row 7 symbology (UPC-A):     '{t4Main?.Cells[7,  9].Text}' (expect 'UPCA')");
Console.WriteLine($"  Row 8  col 1 (UPC-A Scan 1): '{t4Main?.Cells[8,  1].Text}' (expect 'Scan 1')");
Console.WriteLine($"  Row 12 col 1 (UPC-A Scan 5): '{t4Main?.Cells[12, 1].Text}' (expect 'Scan 5')");
Console.WriteLine($"  Row 13 symbology (EAN-13):   '{t4Main?.Cells[13, 9].Text}' (expect 'EAN13')");
Console.WriteLine($"  Row 14 col 1 (EAN Scan 1):   '{t4Main?.Cells[14, 1].Text}' (expect 'Scan 1')");
Console.WriteLine($"  Row 18 col 1 (EAN Scan 5):   '{t4Main?.Cells[18, 1].Text}' (expect 'Scan 5')");

// ── Sidecar full-state resume test ────────────────────────────────────────────
// Simulate a resume: write a sidecar with full context, then call StartSession
// with a nearly-empty state; verify all fields are restored.
var resumeTestSession = new SessionState
{
    OutputDirectory = t4Session.OutputDirectory,
    OutputFormat    = OutputFormat.Xlsx,
    SessionStarted  = new DateTime(2026, 1, 15),
    // Leave context fields null — they should be restored from sidecar.
};
var resumeTestOutputPath  = ExcelFileManager.ResolveOutputPath(resumeTestSession, OutputFormat.Xlsx);
var resumeTestSidecarPath = SessionManager.GetSidecarPath(resumeTestOutputPath);

// Clean up any leftover files from prior test runs before setting up the resume test.
if (File.Exists(resumeTestOutputPath))  File.Delete(resumeTestOutputPath);
if (File.Exists(resumeTestSidecarPath)) File.Delete(resumeTestSidecarPath);

// Write a synthetic sidecar mimicking a prior session.
var priorState = new SessionState
{
    JobName         = "ResumeTestJob",
    OperatorId      = "OP99",
    RollNumber      = 7,
    BatchNumber     = "BTC-RT",
    CompanyName     = "Resume Corp",
    ProductName     = "ResumeProduct",
    CustomNote      = "Resumed note",
    User1           = "U1Val",
    User2           = "U2Val",
    DeviceSerial    = "SN-RT-001",
    DeviceName      = "DM-RT",
    FirmwareVersion = "1.2.3",
    CalibrationDate = new DateTime(2026, 1, 10),
    OutputDirectory = t4Session.OutputDirectory,
    OutputFormat    = OutputFormat.Xlsx,
    SessionStarted  = new DateTime(2026, 1, 15),
    RecordCount     = 42,
};
// Create a real (minimal) Excel file at resumeTestOutputPath so the resume path
// has a valid xlsx to open. This simulates a crash scenario.
using (var resumeSetupAdapter = new XlsxAdapter())
{
    var resumeSetupState = new SessionState
    {
        SessionStarted = new DateTime(2026, 1, 15),
        OutputDirectory = t4Session.OutputDirectory, OutputFormat = OutputFormat.Xlsx,
    };
    using var resumeSetupWriter = new ExcelWriter(resumeSetupAdapter, schema, resumeSetupState);
    resumeSetupWriter.Open(resumeTestOutputPath);  // opens the exact target path
    resumeSetupWriter.Save();
}
// Write the sidecar directly as JSON to simulate what SessionManager persists.
string syntheticSidecarJson = System.Text.Json.JsonSerializer.Serialize(new
{
    priorState.JobName, priorState.OperatorId, priorState.RollNumber,
    priorState.BatchNumber, priorState.CompanyName, priorState.ProductName,
    priorState.CustomNote, priorState.User1, priorState.User2,
    priorState.DeviceSerial, priorState.DeviceName, priorState.FirmwareVersion,
    priorState.CalibrationDate,
    OutputFormat    = priorState.OutputFormat.ToString(),
    priorState.OutputDirectory,
    FileNamePattern = (string?)null,
    priorState.SessionStarted,
    priorState.RecordCount,
    IsNewRollSaved  = false,
}, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
File.WriteAllText(resumeTestSidecarPath, syntheticSidecarJson);

// Now do a real StartSession with empty context — should restore from sidecar.
// Use t4Session output path so the real Excel adapter can open it.
var resumeCheckState = new SessionState
{
    OutputDirectory = t4Session.OutputDirectory,
    OutputFormat    = OutputFormat.Xlsx,
    SessionStarted  = new DateTime(2026, 1, 15),
};
using (var resumeMgr = new SessionManager(schema))
{
    resumeMgr.StartSession(resumeCheckState);
    resumeMgr.CloseSession();
}
// The sidecar is deleted by CloseSession; verify state was restored before close.
// (We check resumeCheckState directly since it's the live reference.)
bool resumeJobName     = resumeCheckState.JobName         == "ResumeTestJob";
bool resumeOperator    = resumeCheckState.OperatorId      == "OP99";
bool resumeRoll        = resumeCheckState.RollNumber      == 7;
bool resumeCompany     = resumeCheckState.CompanyName     == "Resume Corp";
bool resumeProduct     = resumeCheckState.ProductName     == "ResumeProduct";
bool resumeNote        = resumeCheckState.CustomNote      == "Resumed note";
bool resumeUser1       = resumeCheckState.User1           == "U1Val";
bool resumeDevice      = resumeCheckState.DeviceSerial    == "SN-RT-001";
bool resumeCount       = resumeCheckState.RecordCount     == 42;
bool resumeAll = resumeJobName && resumeOperator && resumeRoll && resumeCompany &&
                 resumeProduct && resumeNote && resumeUser1 && resumeDevice && resumeCount;
Console.WriteLine($"\nSidecar full-state resume: {(resumeAll ? "PASS" : "FAIL")}");
if (!resumeAll)
{
    if (!resumeJobName)  Console.WriteLine($"  JobName   = '{resumeCheckState.JobName}'");
    if (!resumeOperator) Console.WriteLine($"  OperatorId= '{resumeCheckState.OperatorId}'");
    if (!resumeRoll)     Console.WriteLine($"  RollNumber= {resumeCheckState.RollNumber}");
    if (!resumeCompany)  Console.WriteLine($"  CompanyName= '{resumeCheckState.CompanyName}'");
    if (!resumeNote)     Console.WriteLine($"  CustomNote= '{resumeCheckState.CustomNote}'");
    if (!resumeUser1)    Console.WriteLine($"  User1     = '{resumeCheckState.User1}'");
    if (!resumeDevice)   Console.WriteLine($"  DeviceSerial= '{resumeCheckState.DeviceSerial}'");
    if (!resumeCount)    Console.WriteLine($"  RecordCount= {resumeCheckState.RecordCount}");
}
// Clean up resume test files.
if (File.Exists(resumeTestOutputPath))  File.Delete(resumeTestOutputPath);
if (File.Exists(resumeTestSidecarPath)) File.Delete(resumeTestSidecarPath);

// ── Resume append-correctness test ───────────────────────────────────────────
// Verifies that after a crash+resume the session appends records to the SAME
// physical file (not a new one), and that the final row count is correct.
//
// Sequence:
//   Phase 1: write 2 records, close cleanly → xlsx has 2 data rows on disk.
//   Phase 2: inject a sidecar mimicking a crash (RecordCount=2, context set).
//   Phase 3: resume with a fresh SessionManager → must open the same xlsx.
//            write 2 more records, close.
//   Verify : xlsx contains exactly 2+2 = 4 data rows (+ 2 header rows = 6 total).

Console.WriteLine("\nResume append-correctness test:");
var appendOutputDir = Path.Combine(Path.GetTempPath(), "vtccp_append_test");
Directory.CreateDirectory(appendOutputDir);

// Phase 1: initial session — write 2 records, close cleanly.
var appendSession1 = new SessionState
{
    JobName         = "AppendTest",
    OperatorId      = "OP-APP",
    RollNumber      = 1,
    OutputDirectory = appendOutputDir,
    OutputFormat    = OutputFormat.Xlsx,
    SessionStarted  = new DateTime(2026, 2, 1),
};
using var appendMgr1 = new SessionManager(schema);
string appendPath1 = appendMgr1.StartSession(appendSession1);
var t4AppPass = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS");
var appRec1 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 2, 1, 9, 0, 0),
    Symbology = "DataMatrix", SymbologyFamily = SymbologyFamily.DataMatrix,
    DecodedData = "APPEND-REC-01", FormalGrade = "4.0/16/660/45Q",
    OverallGrade = t4AppPass, CustomPassFail = OverallPassFail.Pass,
    OperatorId = "OP-APP", JobName = "AppendTest", RollNumber = 1,
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    MatrixSize = "16x16", HorizontalBWG = 0, VerticalBWG = 0,
    EncodedCharacters = 12, TotalCodewords = 24, DataCodewords = 12,
    ErrorCorrectionBudget = 12, ErrorsCorrected = 0, ErrorCapacityUsed = 0,
    ErrorCorrectionType = "ECC 200", ImagePolarity = ImagePolarity.BlackOnWhite,
    NominalXDim_2D = 25.0m, PixelsPerModule = 41.5m,
    UEC_Percent = 100, UEC_Grade = t4AppPass,
    SC_Percent = 80, SC_Grade = t4AppPass,
    MOD_Grade = t4AppPass, RM_Grade = t4AppPass, ANU_Percent = 0m, ANU_Grade = t4AppPass,
    GNU_Percent = 0m, GNU_Grade = t4AppPass, FPD_Grade = t4AppPass,
    LLS_Grade = t4AppPass, BLS_Grade = t4AppPass, LQZ_Grade = t4AppPass,
    BQZ_Grade = t4AppPass, TQZ_Grade = t4AppPass, RQZ_Grade = t4AppPass,
    TTR_Percent = 0, TTR_Grade = t4AppPass, RTR_Percent = 0, RTR_Grade = t4AppPass,
    TCT_Grade = t4AppPass, RCT_Grade = t4AppPass,
    AG_Value = 4.0m, AG_Grade = t4AppPass, DECODE_Grade = t4AppPass,
};
var appRec2 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 2, 1, 9, 1, 0),
    Symbology = "DataMatrix", SymbologyFamily = SymbologyFamily.DataMatrix,
    DecodedData = "APPEND-REC-02", FormalGrade = "4.0/16/660/45Q",
    OverallGrade = t4AppPass, CustomPassFail = OverallPassFail.Pass,
    OperatorId = "OP-APP", JobName = "AppendTest", RollNumber = 1,
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    MatrixSize = "16x16", HorizontalBWG = 0, VerticalBWG = 0,
    EncodedCharacters = 12, TotalCodewords = 24, DataCodewords = 12,
    ErrorCorrectionBudget = 12, ErrorsCorrected = 0, ErrorCapacityUsed = 0,
    ErrorCorrectionType = "ECC 200", ImagePolarity = ImagePolarity.BlackOnWhite,
    NominalXDim_2D = 25.0m, PixelsPerModule = 41.5m,
    UEC_Percent = 100, UEC_Grade = t4AppPass,
    SC_Percent = 80, SC_Grade = t4AppPass,
    MOD_Grade = t4AppPass, RM_Grade = t4AppPass, ANU_Percent = 0m, ANU_Grade = t4AppPass,
    GNU_Percent = 0m, GNU_Grade = t4AppPass, FPD_Grade = t4AppPass,
    LLS_Grade = t4AppPass, BLS_Grade = t4AppPass, LQZ_Grade = t4AppPass,
    BQZ_Grade = t4AppPass, TQZ_Grade = t4AppPass, RQZ_Grade = t4AppPass,
    TTR_Percent = 0, TTR_Grade = t4AppPass, RTR_Percent = 0, RTR_Grade = t4AppPass,
    TCT_Grade = t4AppPass, RCT_Grade = t4AppPass,
    AG_Value = 4.0m, AG_Grade = t4AppPass, DECODE_Grade = t4AppPass,
};
appendMgr1.AddRecord(appRec1);
appendMgr1.AddRecord(appRec2);
int appendPhase1RecordCount = appendMgr1.RecordsWritten;
appendMgr1.CloseSession();   // saves xlsx, deletes sidecar

string appendSidecarPath = SessionManager.GetSidecarPath(appendPath1);

// Verify phase 1 saved correctly.
using var appendPkg1 = new OfficeOpenXml.ExcelPackage(new FileInfo(appendPath1));
int appendRows1 = appendPkg1.Workbook.Worksheets["Main"]?.Dimension?.Rows ?? 0;
Console.WriteLine($"  Phase 1: {appendPhase1RecordCount} records written, xlsx has {appendRows1} rows (expect 4 = 2 hdr + 2 data)");
appendPkg1.Dispose();

// Phase 2: inject sidecar (simulates crash after Phase 1 save).
// Written as an anonymous object so we don't reference the internal SessionSidecar class.
var crashSidecarJson = new
{
    JobName           = "AppendTest",
    OperatorId        = "OP-APP",
    RollNumber        = 1,
    RollIncrementMode = "Manual",
    RollStartValue    = 1,
    RollTimestamp     = (string?)null,
    OutputDirectory   = appendOutputDir,
    OutputFormat      = "Xlsx",
    FileNamePattern   = (string?)null,
    SessionStarted    = new DateTime(2026, 2, 1),
    RecordCount       = appendPhase1RecordCount,
};
File.WriteAllText(appendSidecarPath,
    System.Text.Json.JsonSerializer.Serialize(crashSidecarJson,
        new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));

// Phase 3: resume session and append 2 more records.
// In a real crash-recovery scenario the caller supplies the job name (or uses
// a file picker in Phase 3 GUI) so the initial path resolves to the right file.
var appendSession2 = new SessionState
{
    JobName         = "AppendTest",  // caller provides job name for path resolution
    OutputDirectory = appendOutputDir,
    OutputFormat    = OutputFormat.Xlsx,
    SessionStarted  = new DateTime(2026, 2, 1),  // same date → same filename
    // All other context fields left null — sidecar restores them
};
using var appendMgr2 = new SessionManager(schema);
string appendPath2 = appendMgr2.StartSession(appendSession2);
bool appendSameFile = appendPath2 == appendPath1;
Console.WriteLine($"  Phase 3: resumed path same as original: {appendSameFile}");
Console.WriteLine($"  Phase 3: restored RecordCount = {appendMgr2.CurrentSession!.RecordCount} (expect {appendPhase1RecordCount})");

var appRec3 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 2, 1, 10, 0, 0),
    Symbology = "DataMatrix", SymbologyFamily = SymbologyFamily.DataMatrix,
    DecodedData = "APPEND-REC-03", FormalGrade = "4.0/16/660/45Q",
    OverallGrade = t4AppPass, CustomPassFail = OverallPassFail.Pass,
    OperatorId = "OP-APP", JobName = "AppendTest", RollNumber = 1,
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    MatrixSize = "16x16", HorizontalBWG = 0, VerticalBWG = 0,
    EncodedCharacters = 12, TotalCodewords = 24, DataCodewords = 12,
    ErrorCorrectionBudget = 12, ErrorsCorrected = 0, ErrorCapacityUsed = 0,
    ErrorCorrectionType = "ECC 200", ImagePolarity = ImagePolarity.BlackOnWhite,
    NominalXDim_2D = 25.0m, PixelsPerModule = 41.5m,
    UEC_Percent = 100, UEC_Grade = t4AppPass,
    SC_Percent = 80, SC_Grade = t4AppPass,
    MOD_Grade = t4AppPass, RM_Grade = t4AppPass, ANU_Percent = 0m, ANU_Grade = t4AppPass,
    GNU_Percent = 0m, GNU_Grade = t4AppPass, FPD_Grade = t4AppPass,
    LLS_Grade = t4AppPass, BLS_Grade = t4AppPass, LQZ_Grade = t4AppPass,
    BQZ_Grade = t4AppPass, TQZ_Grade = t4AppPass, RQZ_Grade = t4AppPass,
    TTR_Percent = 0, TTR_Grade = t4AppPass, RTR_Percent = 0, RTR_Grade = t4AppPass,
    TCT_Grade = t4AppPass, RCT_Grade = t4AppPass,
    AG_Value = 4.0m, AG_Grade = t4AppPass, DECODE_Grade = t4AppPass,
};
var appRec4 = new VerificationRecord
{
    VerificationDateTime = new DateTime(2026, 2, 1, 10, 1, 0),
    Symbology = "DataMatrix", SymbologyFamily = SymbologyFamily.DataMatrix,
    DecodedData = "APPEND-REC-04", FormalGrade = "4.0/16/660/45Q",
    OverallGrade = t4AppPass, CustomPassFail = OverallPassFail.Pass,
    OperatorId = "OP-APP", JobName = "AppendTest", RollNumber = 1,
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    MatrixSize = "16x16", HorizontalBWG = 0, VerticalBWG = 0,
    EncodedCharacters = 12, TotalCodewords = 24, DataCodewords = 12,
    ErrorCorrectionBudget = 12, ErrorsCorrected = 0, ErrorCapacityUsed = 0,
    ErrorCorrectionType = "ECC 200", ImagePolarity = ImagePolarity.BlackOnWhite,
    NominalXDim_2D = 25.0m, PixelsPerModule = 41.5m,
    UEC_Percent = 100, UEC_Grade = t4AppPass,
    SC_Percent = 80, SC_Grade = t4AppPass,
    MOD_Grade = t4AppPass, RM_Grade = t4AppPass, ANU_Percent = 0m, ANU_Grade = t4AppPass,
    GNU_Percent = 0m, GNU_Grade = t4AppPass, FPD_Grade = t4AppPass,
    LLS_Grade = t4AppPass, BLS_Grade = t4AppPass, LQZ_Grade = t4AppPass,
    BQZ_Grade = t4AppPass, TQZ_Grade = t4AppPass, RQZ_Grade = t4AppPass,
    TTR_Percent = 0, TTR_Grade = t4AppPass, RTR_Percent = 0, RTR_Grade = t4AppPass,
    TCT_Grade = t4AppPass, RCT_Grade = t4AppPass,
    AG_Value = 4.0m, AG_Grade = t4AppPass, DECODE_Grade = t4AppPass,
};
appendMgr2.AddRecord(appRec3);
appendMgr2.AddRecord(appRec4);
appendMgr2.CloseSession();

// Verify final row count: 2 headers + 2 original + 2 resumed = 6 total.
using var appendPkg2 = new OfficeOpenXml.ExcelPackage(new FileInfo(appendPath2));
int appendRowsFinal = appendPkg2.Workbook.Worksheets["Main"]?.Dimension?.Rows ?? 0;
appendPkg2.Dispose();
Console.WriteLine($"  Final row count: {appendRowsFinal} (expect 6 = 2 hdr + 2 original + 2 resumed)");

bool appendCorrect =
    appendSameFile &&
    appendRows1 == 4 &&
    appendRowsFinal == 6;

Console.WriteLine($"  Resume append-correctness: {(appendCorrect ? "PASS" : "FAIL")}");

// Cleanup.
foreach (var f in Directory.GetFiles(appendOutputDir)) File.Delete(f);
Directory.Delete(appendOutputDir);

// ── AutoIncrement roll mode test ──────────────────────────────────────────────
Console.WriteLine("\nAutoIncrement roll mode test:");
var aiOutputDir = Path.Combine(Path.GetTempPath(), "vtccp_ai_test");
Directory.CreateDirectory(aiOutputDir);
var aiState = new SessionState
{
    JobName          = "AutoIncrJob",
    OutputDirectory  = aiOutputDir,
    OutputFormat     = OutputFormat.Xlsx,
    SessionStarted   = new DateTime(2026, 1, 15),
    RollIncrementMode = RollIncrementMode.AutoIncrement,
    RollStartValue   = 5,
};
using var aiMgr = new SessionManager(schema);
string aiPath = aiMgr.StartSession(aiState);
string aiLabel1 = aiMgr.CurrentSession!.RollLabel;
Console.WriteLine($"  RollLabel at open (start=5): '{aiLabel1}' (expect '5')");

aiMgr.SetNewOperatorAndRoll("OP-AI");
string aiLabel2 = aiMgr.CurrentSession!.RollLabel;
Console.WriteLine($"  RollLabel after 1st increment: '{aiLabel2}' (expect '6')");

aiMgr.SetNewOperatorAndRoll("OP-AI");
string aiLabel3 = aiMgr.CurrentSession!.RollLabel;
Console.WriteLine($"  RollLabel after 2nd increment: '{aiLabel3}' (expect '7')");

aiMgr.CloseSession();
// Cleanup
foreach (var f in Directory.GetFiles(aiOutputDir)) File.Delete(f);
Directory.Delete(aiOutputDir);

bool aiStart   = aiLabel1 == "5";
bool aiInc1    = aiLabel2 == "6";
bool aiInc2    = aiLabel3 == "7";
bool aiPass    = aiStart && aiInc1 && aiInc2;
Console.WriteLine($"  AutoIncrement: start={aiStart} inc1={aiInc1} inc2={aiInc2} => {(aiPass ? "PASS" : "FAIL")}");

// ── DateTimeStamp roll mode test ──────────────────────────────────────────────
Console.WriteLine("\nDateTimeStamp roll mode test:");
var dtOutputDir = Path.Combine(Path.GetTempPath(), "vtccp_dt_test");
Directory.CreateDirectory(dtOutputDir);
var dtState = new SessionState
{
    JobName          = "DateTimeJob",
    OutputDirectory  = dtOutputDir,
    OutputFormat     = OutputFormat.Xlsx,
    SessionStarted   = new DateTime(2026, 1, 15),
    RollIncrementMode = RollIncrementMode.DateTimeStamp,
};
using var dtMgr = new SessionManager(schema);
string dtPath = dtMgr.StartSession(dtState);
string dtLabel1 = dtMgr.CurrentSession!.RollLabel;
bool dtLabel1Valid = dtLabel1.Length == 14 && dtLabel1.All(char.IsDigit);
Console.WriteLine($"  RollLabel at open: '{dtLabel1}' (length={dtLabel1.Length}, all-digits={dtLabel1Valid})");

// Small delay so the timestamp can differ.
System.Threading.Thread.Sleep(1100);
dtMgr.SetNewOperatorAndRoll("OP-DT");
string dtLabel2 = dtMgr.CurrentSession!.RollLabel;
bool dtLabel2Valid   = dtLabel2.Length == 14 && dtLabel2.All(char.IsDigit);
bool dtTimestampDiff = dtLabel1 != dtLabel2;
Console.WriteLine($"  RollLabel after roll change: '{dtLabel2}' (valid={dtLabel2Valid}, changed={dtTimestampDiff})");

dtMgr.CloseSession();
// Cleanup
foreach (var f in Directory.GetFiles(dtOutputDir)) File.Delete(f);
Directory.Delete(dtOutputDir);

bool dtPass = dtLabel1Valid && dtLabel2Valid && dtTimestampDiff;
Console.WriteLine($"  DateTimeStamp: open={dtLabel1Valid} afterRoll={dtLabel2Valid} changed={dtTimestampDiff} => {(dtPass ? "PASS" : "FAIL")}");

// ── GS1 Auto-Batch extraction test ────────────────────────────────────────────
Console.WriteLine("\nGS1 auto-batch test:");

// 1. Parser unit tests — various string formats
var gs1_simple   = "<F1>010123456789012817251231101234-LOT-A<F1>2199887766";
var gs1_fixchain = "<F1>01012345678901281725123110BATCHB";  // AI(10) directly after two fixed-length AIs
var gs1_raw      = "\u001d010123456789012817251231101234-LOT-A\u001d2199887766";
var gs1_noai10   = "<F1>010123456789012817251231";           // no AI(10) at all
var gs1_nonfmt   = "CODE128-DATA-ONLY";                     // not GS1 at all

bool p1 = ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(gs1_simple)   == "1234-LOT-A";
bool p2 = ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(gs1_fixchain) == "BATCHB";
bool p3 = ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(gs1_raw)      == "1234-LOT-A";
bool p4 = ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(gs1_noai10)   == null;
bool p5 = ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(gs1_nonfmt)   == null;
bool p6 = ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(null)         == null;

Console.WriteLine($"  simple FNC1-sep:   {(p1 ? "PASS" : $"FAIL (got '{ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(gs1_simple)}')")}");
Console.WriteLine($"  fixed-len chain:   {(p2 ? "PASS" : $"FAIL (got '{ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(gs1_fixchain)}')")}");
Console.WriteLine($"  raw GS char:       {(p3 ? "PASS" : $"FAIL (got '{ExcelEngine.Utilities.GS1Parser.ExtractBatchLot(gs1_raw)}')")}");
Console.WriteLine($"  no AI(10):         {(p4 ? "PASS" : "FAIL (expected null)")}");
Console.WriteLine($"  non-GS1 string:    {(p5 ? "PASS" : "FAIL (expected null)")}");
Console.WriteLine($"  null input:        {(p6 ? "PASS" : "FAIL (expected null)")}");

// 1b. ISO 15434 / ANSI MH10.8.2 / MIL-STD-130 parser unit tests.
Console.WriteLine("\n  ISO 15434 / MH10.8.2 / MIL-STD-130 parser:");

// Full envelope with raw control characters — 4L batch DI.
var mh_4l_raw   = "[)>\u001e06\u001dP1234567890\u001d4LMIL-LOT-42\u001dS99887766\u001e\u0004";
// Full envelope with DataMan-style text escapes — 4L batch DI.
var mh_4l_esc   = "[)><RS>06<GS>P1234567890<GS>4LBATCH-ESC<GS>S99887766<RS><EOT>";
// 10L alternate lot DI (no 4L present).
var mh_10l      = "[)>\u001e06\u001dP0987654321\u001d10LALTERNATE-LOT\u001e";
// Envelope present but no batch DI at all.
var mh_no_batch = "[)>\u001e06\u001dP1234567890\u001dS99887766\u001e";
// Not a 15434 envelope.
var mh_not_env  = "plain-text-data";

bool q1 = ExcelEngine.Utilities.ISO15434Parser.ExtractBatchLot(mh_4l_raw)   == "MIL-LOT-42";
bool q2 = ExcelEngine.Utilities.ISO15434Parser.ExtractBatchLot(mh_4l_esc)   == "BATCH-ESC";
bool q3 = ExcelEngine.Utilities.ISO15434Parser.ExtractBatchLot(mh_10l)      == "ALTERNATE-LOT";
bool q4 = ExcelEngine.Utilities.ISO15434Parser.ExtractBatchLot(mh_no_batch) == null;
bool q5 = ExcelEngine.Utilities.ISO15434Parser.ExtractBatchLot(mh_not_env)  == null;

Console.WriteLine($"  4L raw ctrl chars: {(q1 ? "PASS" : $"FAIL (got '{ExcelEngine.Utilities.ISO15434Parser.ExtractBatchLot(mh_4l_raw)}')")}");
Console.WriteLine($"  4L text escapes:   {(q2 ? "PASS" : $"FAIL (got '{ExcelEngine.Utilities.ISO15434Parser.ExtractBatchLot(mh_4l_esc)}')")}");
Console.WriteLine($"  10L alt DI:        {(q3 ? "PASS" : $"FAIL (got '{ExcelEngine.Utilities.ISO15434Parser.ExtractBatchLot(mh_10l)}')")}");
Console.WriteLine($"  no batch DI:       {(q4 ? "PASS" : "FAIL (expected null)")}");
Console.WriteLine($"  not 15434:         {(q5 ? "PASS" : "FAIL (expected null)")}");

// 2. End-to-end: SessionManager AutoFromGS1 stamps AI(10) into the Batch cell.
var gs1BatchDir  = Path.Combine(Path.GetTempPath(), "vtccp_gs1batch_test");
Directory.CreateDirectory(gs1BatchDir);

var gs1State = new ExcelEngine.Models.SessionState
{
    JobName         = "GS1AutoBatch",
    OperatorId      = "GW4",
    BatchMode       = ExcelEngine.Models.BatchMode.AutoFromGS1,
    OutputFormat    = ExcelEngine.Models.OutputFormat.Xlsx,
    DeviceSerial    = "TEST",
    DeviceName      = "DMTest",
    FirmwareVersion = "1.0",
    OutputDirectory = gs1BatchDir,
    SessionStarted  = new DateTime(2025, 1, 1),
};

// Record whose decoded data contains AI(10)=AUTO-BATCH-01; caller sets BatchNumber null.
var gs1Rec1 = new ExcelEngine.Models.VerificationRecord
{
    VerificationDateTime = new DateTime(2025, 1, 1, 10, 0, 0),
    Symbology       = "GS1 DataMatrix",
    SymbologyFamily = ExcelEngine.Models.SymbologyFamily.GS1DataMatrix,
    DecodedData     = "<F1>010123456789012817251231101234-LOT-A<F1>2199887766",
    FormalGrade     = "4.0/16/660/45Q",
    OverallGrade    = ExcelEngine.Models.GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS"),
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    // BatchNumber intentionally null — should be auto-populated from AI(10)
};

// Record with no AI(10) — should fall through to null (empty Batch cell).
var gs1Rec2 = new ExcelEngine.Models.VerificationRecord
{
    VerificationDateTime = new DateTime(2025, 1, 1, 10, 1, 0),
    Symbology       = "GS1 DataMatrix",
    SymbologyFamily = ExcelEngine.Models.SymbologyFamily.GS1DataMatrix,
    DecodedData     = "<F1>010123456789012817251231",
    FormalGrade     = "4.0/16/660/45Q",
    OverallGrade    = ExcelEngine.Models.GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS"),
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
    BatchNumber     = "MANUAL-FALLBACK",   // caller-supplied; used when no AI(10) found
};

string gs1BatchPath;
using (var gs1Mgr = new SessionManager(schema))
{
    gs1BatchPath = gs1Mgr.StartSession(gs1State);
    gs1Mgr.AddRecord(gs1Rec1);
    gs1Mgr.AddRecord(gs1Rec2);
    gs1Mgr.CloseSession();
}

// Verify Batch cells (col 6 = BatchNumber in schema).
string? batchCell1 = null, batchCell2 = null;
OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
using (var pkg = new OfficeOpenXml.ExcelPackage(new FileInfo(gs1BatchPath)))
{
    var ws = pkg.Workbook.Worksheets["Main"];
    batchCell1 = ws.Cells[3, 6].Text;   // row 3 = first data row (title+header = rows 1-2)
    batchCell2 = ws.Cells[4, 6].Text;
}

bool gs1e2ePass1 = batchCell1 == "1234-LOT-A";
bool gs1e2ePass2 = batchCell2 == "MANUAL-FALLBACK";
Console.WriteLine($"  e2e AI(10) stamp:  {(gs1e2ePass1 ? "PASS" : $"FAIL (got '{batchCell1}')")}");
Console.WriteLine($"  e2e fallback:      {(gs1e2ePass2 ? "PASS" : $"FAIL (got '{batchCell2}')")}");

// 3. End-to-end: SessionManager AutoFromGS1 with ISO 15434 / MH10.8.2 decoded data.
Console.WriteLine("\n  15434 e2e via AutoBatchExtractor:");

var mh10BatchDir = Path.Combine(Path.GetTempPath(), "vtccp_mh10_test");
Directory.CreateDirectory(mh10BatchDir);

var mh10State = new ExcelEngine.Models.SessionState
{
    JobName         = "MH10AutoBatch",
    OperatorId      = "GW4",
    BatchMode       = ExcelEngine.Models.BatchMode.AutoFromGS1,
    OutputFormat    = ExcelEngine.Models.OutputFormat.Xlsx,
    DeviceSerial    = "TEST",
    DeviceName      = "DMTest",
    FirmwareVersion = "1.0",
    OutputDirectory = mh10BatchDir,
    SessionStarted  = new DateTime(2025, 2, 1),
};

// Record with ISO 15434 decoded data containing 4L batch DI.
var mhRec1 = new ExcelEngine.Models.VerificationRecord
{
    VerificationDateTime = new DateTime(2025, 2, 1, 10, 0, 0),
    Symbology       = "DataMatrix",
    SymbologyFamily = ExcelEngine.Models.SymbologyFamily.GS1DataMatrix,
    DecodedData     = "[)>\u001e06\u001dP1234567890\u001d4LMIL-LOT-42\u001dS99887766\u001e\u0004",
    FormalGrade     = "4.0/16/660/45Q",
    OverallGrade    = ExcelEngine.Models.GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS"),
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
};

// Record with DataMan text-escape 15434 (4L DI).
var mhRec2 = new ExcelEngine.Models.VerificationRecord
{
    VerificationDateTime = new DateTime(2025, 2, 1, 10, 1, 0),
    Symbology       = "DataMatrix",
    SymbologyFamily = ExcelEngine.Models.SymbologyFamily.GS1DataMatrix,
    DecodedData     = "[)><RS>06<GS>P9876543210<GS>4LESC-LOT-99<GS>S11223344<RS><EOT>",
    FormalGrade     = "4.0/16/660/45Q",
    OverallGrade    = ExcelEngine.Models.GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS"),
    Aperture = 16, Wavelength = 660, Lighting = "45Q", Standard = "ISO 15415:2011",
};

string mh10BatchPath;
using (var mh10Mgr = new SessionManager(schema))
{
    mh10BatchPath = mh10Mgr.StartSession(mh10State);
    mh10Mgr.AddRecord(mhRec1);
    mh10Mgr.AddRecord(mhRec2);
    mh10Mgr.CloseSession();
}

string? mhBatchCell1 = null, mhBatchCell2 = null;
using (var pkg = new OfficeOpenXml.ExcelPackage(new FileInfo(mh10BatchPath)))
{
    var ws = pkg.Workbook.Worksheets["Main"];
    mhBatchCell1 = ws.Cells[3, 6].Text;
    mhBatchCell2 = ws.Cells[4, 6].Text;
}

bool mh10e2ePass1 = mhBatchCell1 == "MIL-LOT-42";
bool mh10e2ePass2 = mhBatchCell2 == "ESC-LOT-99";
Console.WriteLine($"  4L raw ctrl e2e:   {(mh10e2ePass1 ? "PASS" : $"FAIL (got '{mhBatchCell1}')")}");
Console.WriteLine($"  4L text-esc e2e:   {(mh10e2ePass2 ? "PASS" : $"FAIL (got '{mhBatchCell2}')")}");

bool gs1BatchPass = p1 && p2 && p3 && p4 && p5 && p6
                 && q1 && q2 && q3 && q4 && q5
                 && gs1e2ePass1 && gs1e2ePass2
                 && mh10e2ePass1 && mh10e2ePass2;
Console.WriteLine($"  Auto-batch (GS1 + 15434 + MIL-STD): {(gs1BatchPass ? "PASS" : "FAIL")}");

// ── Full Task 4 pass/fail ─────────────────────────────────────────────────────
bool t4Pass =
    // Sanitization
    sanSlash && sanSpace && sanBracket && sanGlob && sanAmpersand &&
    // Custom filename pattern
    patternOk &&
    // Roll semantics (manual mode)
    rollNewSession && rollIncrement && opChanged &&
    // Roll increment modes
    aiPass && dtPass &&
    // GS1 auto-batch
    gs1BatchPass &&
    // Resume append correctness
    appendCorrect &&
    // SessionManager lifecycle
    recordsWrittenBeforeClose == 6 &&
    sidecarCreated        == true &&
    sidecarStillPresent   == true &&
    sidecarDeleted        == true &&
    // Output file structure
    t4Rows                == 18 &&
    svMarker              == "VTCCP" &&
    svName                == schema.Name &&
    svVersion             == schema.Version &&
    t4Main!.Cells[3, 9].Text == "GS1 DataMatrix" &&
    t4Main.Cells[4, 9].Text  == "GS1 DataMatrix" &&
    t4Main.Cells[5, 9].Text  == "GS1 DataMatrix" &&
    t4Main.Cells[6, 9].Text  == "DataMatrix" &&
    t4Main.Cells[7, 9].Text  == "UPCA" &&
    t4Main.Cells[8, 1].Text  == "Scan 1" &&
    t4Main.Cells[12, 1].Text == "Scan 5" &&
    t4Main.Cells[13, 9].Text == "EAN13" &&
    t4Main.Cells[14, 1].Text == "Scan 1" &&
    t4Main.Cells[18, 1].Text == "Scan 5" &&
    // Sidecar full-state resume
    resumeAll;

Console.WriteLine($"\nTask 4 verification: {(t4Pass ? "PASS" : "FAIL")}");
if (!t4Pass)
{
    Console.WriteLine("  Diagnostics:");
    if (!sanSlash || !sanSpace || !sanBracket || !sanGlob)
                                         Console.WriteLine($"    Sanitization failures: slash={sanSlash} space={sanSpace} bracket={sanBracket} glob={sanGlob}");
    if (!patternOk)                      Console.WriteLine($"    Custom pattern: '{customFileName}'");
    if (recordsWrittenBeforeClose != 6)  Console.WriteLine($"    RecordsWritten = {recordsWrittenBeforeClose} (expect 6)");
    if (!sidecarCreated)                 Console.WriteLine("    Sidecar not created on StartSession");
    if (!sidecarStillPresent)            Console.WriteLine("    Sidecar missing while session open");
    if (!sidecarDeleted)                 Console.WriteLine("    Sidecar not deleted after CloseSession");
    if (t4Rows != 18)                    Console.WriteLine($"    Rows = {t4Rows} (expect 18)");
    if (svMarker  != "VTCCP")            Console.WriteLine($"    svMarker  = '{svMarker}'");
    if (svName    != schema.Name)        Console.WriteLine($"    svName    = '{svName}'");
    if (svVersion != schema.Version)     Console.WriteLine($"    svVersion = '{svVersion}'");
    if (!resumeAll)                      Console.WriteLine("    Sidecar resume: FAIL");
}

Console.WriteLine("\nTask 4 complete.");

// ═══════════════════════════════════════════════════════════════════════════════
// Phase 2: Device Integration (DMCC / DMST)
// ═══════════════════════════════════════════════════════════════════════════════
Console.WriteLine("\n─────────────────────────────────────────────────────────");
Console.WriteLine("Phase 2: Device Integration — DMCC / DMST");
Console.WriteLine("─────────────────────────────────────────────────────────");

// ── 2-A: DmccCommand.SanitizeForDmcc ─────────────────────────────────────────
Console.WriteLine("\n2-A  DmccCommand.SanitizeForDmcc:");
bool sc1 = DeviceInterface.Dmcc.DmccCommand.SanitizeForDmcc("Job & Name") == "Job _ Name";
bool sc2 = DeviceInterface.Dmcc.DmccCommand.SanitizeForDmcc("Fine-Name_01") == "Fine-Name_01";
bool sc3 = DeviceInterface.Dmcc.DmccCommand.SanitizeForDmcc("<bad>") == "_bad_";
bool sc4 = DeviceInterface.Dmcc.DmccCommand.SanitizeForDmcc(null) == "";
bool sc5 = DeviceInterface.Dmcc.DmccCommand.SanitizeForDmcc("A\r\nB") == "A__B";
Console.WriteLine($"  & → _:            {(sc1 ? "PASS" : $"FAIL (got '{DeviceInterface.Dmcc.DmccCommand.SanitizeForDmcc("Job & Name")}')")}");
Console.WriteLine($"  safe chars:       {(sc2 ? "PASS" : "FAIL")}");
Console.WriteLine($"  < > → _:          {(sc3 ? "PASS" : $"FAIL (got '{DeviceInterface.Dmcc.DmccCommand.SanitizeForDmcc("<bad>")}')")}");
Console.WriteLine($"  null → empty:     {(sc4 ? "PASS" : "FAIL")}");
Console.WriteLine($"  CRLF → _:         {(sc5 ? "PASS" : "FAIL")}");
bool p2aPass = sc1 && sc2 && sc3 && sc4 && sc5;
Console.WriteLine($"  Sanitize: {(p2aPass ? "PASS" : "FAIL")}");

// ── 2-B: DmccResponse.Parse ───────────────────────────────────────────────────
Console.WriteLine("\n2-B  DmccResponse.Parse:");
var rOk     = DeviceInterface.Dmcc.DmccResponse.Parse("\r\n0\r\n\r\nDM260Q\r\n");
var rNoBody = DeviceInterface.Dmcc.DmccResponse.Parse("\r\n0\r\n");
var rErr    = DeviceInterface.Dmcc.DmccResponse.Parse("\r\n6\r\n");
var rEmpty  = DeviceInterface.Dmcc.DmccResponse.Parse("");

bool rp1 = rOk.IsSuccess && rOk.Body == "DM260Q";
bool rp2 = rNoBody.IsSuccess && rNoBody.Body == "";
bool rp3 = rErr.StatusCode == 6 && !rErr.IsSuccess;
bool rp4 = rEmpty.StatusCode == DeviceInterface.Dmcc.DmccStatus.NoResponse;
Console.WriteLine($"  OK + body:        {(rp1 ? "PASS" : $"FAIL (status={rOk.StatusCode} body='{rOk.Body}')")}");
Console.WriteLine($"  OK no body:       {(rp2 ? "PASS" : $"FAIL (status={rNoBody.StatusCode} body='{rNoBody.Body}')")}");
Console.WriteLine($"  Status 6 (noread):{(rp3 ? "PASS" : $"FAIL (status={rErr.StatusCode})")}");
Console.WriteLine($"  Empty string:     {(rp4 ? "PASS" : $"FAIL (status={rEmpty.StatusCode})")}");
bool p2bPass = rp1 && rp2 && rp3 && rp4;
Console.WriteLine($"  DmccResponse: {(p2bPass ? "PASS" : "FAIL")}");

// ── 2-C: VerificationXmlMap.ClassifySymbology ────────────────────────────────
Console.WriteLine("\n2-C  VerificationXmlMap.ClassifySymbology:");
var xmlMap = new DeviceInterface.Dmst.VerificationXmlMap();
bool cs1 = xmlMap.ClassifySymbology("GS1 DataMatrix")      == ExcelEngine.Models.SymbologyFamily.GS1DataMatrix;
bool cs2 = xmlMap.ClassifySymbology("DataMatrix")          == ExcelEngine.Models.SymbologyFamily.DataMatrix;
bool cs3 = xmlMap.ClassifySymbology("UPC-A")               == ExcelEngine.Models.SymbologyFamily.Linear1D;
bool cs4 = xmlMap.ClassifySymbology("EAN-13")              == ExcelEngine.Models.SymbologyFamily.Linear1D;
bool cs5 = xmlMap.ClassifySymbology("QR Code")             == ExcelEngine.Models.SymbologyFamily.QRCode;
bool cs6 = xmlMap.ClassifySymbology("PDF417")              == ExcelEngine.Models.SymbologyFamily.Linear1D;
bool cs7 = xmlMap.ClassifySymbology(null)                  == ExcelEngine.Models.SymbologyFamily.Unknown;
Console.WriteLine($"  GS1 DataMatrix:   {(cs1 ? "PASS" : "FAIL")}");
Console.WriteLine($"  DataMatrix:       {(cs2 ? "PASS" : "FAIL")}");
Console.WriteLine($"  UPC-A:            {(cs3 ? "PASS" : "FAIL")}");
Console.WriteLine($"  EAN-13:           {(cs4 ? "PASS" : "FAIL")}");
Console.WriteLine($"  QR Code:          {(cs5 ? "PASS" : "FAIL")}");
Console.WriteLine($"  PDF417:           {(cs6 ? "PASS" : "FAIL")}");
Console.WriteLine($"  null → Unknown:   {(cs7 ? "PASS" : "FAIL")}");
bool p2cPass = cs1 && cs2 && cs3 && cs4 && cs5 && cs6 && cs7;
Console.WriteLine($"  ClassifySymbology: {(p2cPass ? "PASS" : "FAIL")}");

// ── 2-D: DmstResultParser — 2D GS1 DataMatrix XML ────────────────────────────
Console.WriteLine("\n2-D  DmstResultParser (2D GS1 DataMatrix):");
var ctx2D = new ExcelEngine.Models.VerificationRecord
{
    Symbology       = "Unknown",
    DeviceSerial    = "DM-TEST-001",
    DeviceName      = "VTCCP-Test-Device",
    FirmwareVersion = "5.7.4.0015",
    OperatorId      = "OP1",
    JobName         = "TestJob",
};
var rec2D = DeviceInterface.Dmst.DmstResultParser.Parse(
    DeviceInterface.Testing.MockDmccServer.SampleDm2DXml, xmlMap, ctx2D);

bool d1  = rec2D.Symbology        == "GS1 DataMatrix";
bool d2  = rec2D.SymbologyFamily  == ExcelEngine.Models.SymbologyFamily.GS1DataMatrix;
bool d3  = rec2D.DecodedData      == "<F1>010123456789012817251231101234-LOT-A<F1>2199887766";
bool d4  = rec2D.FormalGrade      == "4.0/16/660/45Q";
bool d5  = rec2D.OverallGrade?.LetterGradeString == "A";
bool d6  = rec2D.Aperture         == 16;
bool d7  = rec2D.Wavelength       == 660;
bool d8  = rec2D.Lighting         == "45Q";
bool d9  = rec2D.Standard         == "ISO 15415:2011";
bool d10 = rec2D.UEC_Percent      == 100m;
bool d11 = rec2D.SC_Percent       == 84m;
bool d12 = rec2D.ANU_Percent      == 0.2m;
bool d13 = rec2D.GNU_Percent      == 2.3m;
bool d14 = rec2D.AG_Value         == 4.0m;
bool d15 = rec2D.MatrixSize       == "22x22";
bool d16 = rec2D.TotalCodewords   == 144;
bool d17 = rec2D.ErrorsCorrected  == 0;
bool d18 = rec2D.TTR_Percent      == 95.5m;
bool d19 = rec2D.DeviceSerial     == "DM-TEST-001";   // from context
bool d20 = rec2D.OperatorId       == "OP1";            // from context

Console.WriteLine($"  Symbology:        {(d1 ? "PASS" : $"FAIL ('{rec2D.Symbology}')")}");
Console.WriteLine($"  SymbologyFamily:  {(d2 ? "PASS" : $"FAIL ({rec2D.SymbologyFamily})")}");
Console.WriteLine($"  DecodedData:      {(d3 ? "PASS" : $"FAIL ('{rec2D.DecodedData}')")}");
Console.WriteLine($"  FormalGrade:      {(d4 ? "PASS" : $"FAIL ('{rec2D.FormalGrade}')")}");
Console.WriteLine($"  OverallGrade A:   {(d5 ? "PASS" : $"FAIL ('{rec2D.OverallGrade?.LetterGradeString}')")}");
Console.WriteLine($"  Aperture 16:      {(d6 ? "PASS" : $"FAIL ({rec2D.Aperture})")}");
Console.WriteLine($"  Wavelength 660:   {(d7 ? "PASS" : $"FAIL ({rec2D.Wavelength})")}");
Console.WriteLine($"  UEC 100%:         {(d10 ? "PASS" : $"FAIL ({rec2D.UEC_Percent})")}");
Console.WriteLine($"  SC 84%:           {(d11 ? "PASS" : $"FAIL ({rec2D.SC_Percent})")}");
Console.WriteLine($"  ANU 0.2%:         {(d12 ? "PASS" : $"FAIL ({rec2D.ANU_Percent})")}");
Console.WriteLine($"  GNU 2.3%:         {(d13 ? "PASS" : $"FAIL ({rec2D.GNU_Percent})")}");
Console.WriteLine($"  AG 4.0:           {(d14 ? "PASS" : $"FAIL ({rec2D.AG_Value})")}");
Console.WriteLine($"  MatrixSize 22x22: {(d15 ? "PASS" : $"FAIL ('{rec2D.MatrixSize}')")}");
Console.WriteLine($"  TotalCodewords:   {(d16 ? "PASS" : $"FAIL ({rec2D.TotalCodewords})")}");
Console.WriteLine($"  TTRPercent 95.5:  {(d18 ? "PASS" : $"FAIL ({rec2D.TTR_Percent})")}");
Console.WriteLine($"  Context serial:   {(d19 ? "PASS" : $"FAIL ('{rec2D.DeviceSerial}')")}");
Console.WriteLine($"  Context operator: {(d20 ? "PASS" : $"FAIL ('{rec2D.OperatorId}')")}");
bool p2dPass = d1 && d2 && d3 && d4 && d5 && d6 && d7 && d8 && d9
            && d10 && d11 && d12 && d13 && d14 && d15 && d16 && d17
            && d18 && d19 && d20;
Console.WriteLine($"  2D parse: {(p2dPass ? "PASS" : "FAIL")}");

// ── 2-E: DmstResultParser — 1D UPC-A XML ─────────────────────────────────────
Console.WriteLine("\n2-E  DmstResultParser (1D UPC-A):");
var rec1D = DeviceInterface.Dmst.DmstResultParser.Parse(
    DeviceInterface.Testing.MockDmccServer.SampleUpcAXml, xmlMap, ctx2D);

bool e1 = rec1D.Symbology        == "UPCA";
bool e2 = rec1D.SymbologyFamily  == ExcelEngine.Models.SymbologyFamily.Linear1D;
bool e3 = rec1D.DecodedData      == "012345678905";
bool e4 = rec1D.Aperture         == 6;
bool e5 = rec1D.Avg_Edge         == 59m;
bool e6 = rec1D.Avg_SC           == 84m;
bool e7 = rec1D.Avg_LQZ          == 10m;
bool e8 = rec1D.Avg_RQZ          == 11m;
bool e9 = rec1D.Avg_MinQZ        == 10m;
bool e10 = rec1D.BWG_Percent     == 8m;
bool e11 = rec1D.ScanResults.Count == 2;
bool e12 = rec1D.ScanResults[0].ScanNumber == 1;
bool e13 = rec1D.ScanResults[0].Edge       == 4m;
bool e14 = rec1D.ScanResults[0].LQZ        == 4m;
Console.WriteLine($"  Symbology UPCA:   {(e1 ? "PASS" : $"FAIL ('{rec1D.Symbology}')")}");
Console.WriteLine($"  SymbologyFamily:  {(e2 ? "PASS" : $"FAIL ({rec1D.SymbologyFamily})")}");
Console.WriteLine($"  DecodedData:      {(e3 ? "PASS" : $"FAIL ('{rec1D.DecodedData}')")}");
Console.WriteLine($"  Aperture 6:       {(e4 ? "PASS" : $"FAIL ({rec1D.Aperture})")}");
Console.WriteLine($"  AvgEdge 59:       {(e5 ? "PASS" : $"FAIL ({rec1D.Avg_Edge})")}");
Console.WriteLine($"  AvgSC 84:         {(e6 ? "PASS" : $"FAIL ({rec1D.Avg_SC})")}");
Console.WriteLine($"  AvgLQZ 10:        {(e7 ? "PASS" : $"FAIL ({rec1D.Avg_LQZ})")}");
Console.WriteLine($"  AvgMinQZ 10:      {(e9 ? "PASS" : $"FAIL ({rec1D.Avg_MinQZ})")}");
Console.WriteLine($"  BWG% 8:           {(e10 ? "PASS" : $"FAIL ({rec1D.BWG_Percent})")}");
Console.WriteLine($"  ScanCount 2:      {(e11 ? "PASS" : $"FAIL ({rec1D.ScanResults.Count})")}");
Console.WriteLine($"  Scan1 Edge 4:     {(e13 ? "PASS" : $"FAIL ({rec1D.ScanResults[0].Edge})")}");
Console.WriteLine($"  Scan1 LQZ 4:      {(e14 ? "PASS" : $"FAIL ({rec1D.ScanResults[0].LQZ})")}");
bool p2ePass = e1 && e2 && e3 && e4 && e5 && e6 && e7 && e8 && e9 && e10
            && e11 && e12 && e13 && e14;
Console.WriteLine($"  1D parse: {(p2ePass ? "PASS" : "FAIL")}");

// ── 2-F: DmccClient + MockDmccServer loopback round-trip ─────────────────────
Console.WriteLine("\n2-F  DmccClient + MockDmccServer loopback:");
bool p2fPass = false;
string? connBanner   = null;
string? devTypeBody  = null;
string? firmVerBody  = null;
bool    triggerOk    = false;
bool    resultIsXml  = false;

await using (var mockServer = new DeviceInterface.Testing.MockDmccServer())
{
    var testCfg = new DeviceInterface.DeviceConfig
    {
        Host              = "127.0.0.1",
        Port              = mockServer.Port,
        ConnectTimeoutMs  = 2_000,
        ResponseTimeoutMs = 2_000,
        IdleGapMs         = 60,
    };

    await using var client = new DeviceInterface.Dmcc.DmccClient(testCfg);
    await client.ConnectAsync();
    connBanner = client.WelcomeBanner;

    var r1 = await client.SendAsync(DeviceInterface.Dmcc.DmccCommand.GetDeviceType);
    devTypeBody = r1.Body;

    var r2 = await client.SendAsync(DeviceInterface.Dmcc.DmccCommand.GetFirmwareVer);
    firmVerBody = r2.Body;

    var r3 = await client.SendAsync(DeviceInterface.Dmcc.DmccCommand.Trigger);
    triggerOk = r3.IsSuccess;

    var r4 = await client.SendAsync(DeviceInterface.Dmcc.DmccCommand.GetSymbolResult);
    resultIsXml = r4.IsSuccess && r4.IsXml;
}

bool lp1 = connBanner?.Contains("Welcome") == true;
bool lp2 = devTypeBody  == "DM260Q";
bool lp3 = firmVerBody  == "5.7.4.0015";
bool lp4 = triggerOk;
bool lp5 = resultIsXml;
Console.WriteLine($"  Welcome banner:   {(lp1 ? "PASS" : $"FAIL ('{connBanner?.Trim()}')")}");
Console.WriteLine($"  GET DEVICE.TYPE:  {(lp2 ? "PASS" : $"FAIL ('{devTypeBody}')")}");
Console.WriteLine($"  GET FIRMWARE.VER: {(lp3 ? "PASS" : $"FAIL ('{firmVerBody}')")}");
Console.WriteLine($"  TRIGGER ok:       {(lp4 ? "PASS" : "FAIL")}");
Console.WriteLine($"  GET SYMBOL.RESULT is XML: {(lp5 ? "PASS" : "FAIL")}");
p2fPass = lp1 && lp2 && lp3 && lp4 && lp5;
Console.WriteLine($"  Loopback: {(p2fPass ? "PASS" : "FAIL")}");

// ── 2-G: DeviceSession + MockDmccServer full round-trip ──────────────────────
Console.WriteLine("\n2-G  DeviceSession end-to-end (mock):");
bool p2gPass = false;
DeviceInterface.DeviceInfo? devInfo = null;
ExcelEngine.Models.VerificationRecord? sessionRec = null;

await using (var mockServer2 = new DeviceInterface.Testing.MockDmccServer())
{
    var cfg2 = new DeviceInterface.DeviceConfig
    {
        Host              = "127.0.0.1",
        Port              = mockServer2.Port,
        ConnectTimeoutMs  = 2_000,
        ResponseTimeoutMs = 2_000,
        IdleGapMs         = 60,
    };

    await using var devSession = new DeviceInterface.DeviceSession(cfg2, xmlMap);
    await devSession.ConnectAsync();
    devInfo = devSession.DeviceInfo;

    var sessionCtx = new ExcelEngine.Models.VerificationRecord
    {
        Symbology       = "Unknown",
        DeviceSerial    = devInfo.Serial,
        DeviceName      = devInfo.Name,
        FirmwareVersion = devInfo.FirmwareVersion,
        OperatorId      = "OP-LIVE",
        JobName         = "LiveSession",
    };

    sessionRec = await devSession.TriggerAndGetResultAsync(sessionCtx);
}

bool gp1 = devInfo?.Type            == "DM260Q";
bool gp2 = devInfo?.FirmwareVersion == "5.7.4.0015";
bool gp3 = devInfo?.Serial          == "DM-TEST-001";
bool gp4 = devInfo?.Name            == "VTCCP-Test-Device";
bool gp5 = sessionRec is not null;
bool gp6 = sessionRec?.Symbology     == "GS1 DataMatrix";
bool gp7 = sessionRec?.OperatorId    == "OP-LIVE";     // from context
bool gp8 = sessionRec?.DeviceSerial  == "DM-TEST-001"; // from devInfo via context
bool gp9 = sessionRec?.UEC_Percent   == 100m;
bool gp10 = sessionRec?.MatrixSize   == "22x22";
Console.WriteLine($"  DeviceInfo.Type:  {(gp1 ? "PASS" : $"FAIL ('{devInfo?.Type}')")}");
Console.WriteLine($"  DeviceInfo.FW:    {(gp2 ? "PASS" : $"FAIL ('{devInfo?.FirmwareVersion}')")}");
Console.WriteLine($"  DeviceInfo.Serial:{(gp3 ? "PASS" : $"FAIL ('{devInfo?.Serial}')")}");
Console.WriteLine($"  DeviceInfo.Name:  {(gp4 ? "PASS" : $"FAIL ('{devInfo?.Name}')")}");
Console.WriteLine($"  Record returned:  {(gp5 ? "PASS" : "FAIL (null)")}");
Console.WriteLine($"  Record Symbology: {(gp6 ? "PASS" : $"FAIL ('{sessionRec?.Symbology}')")}");
Console.WriteLine($"  Record Operator:  {(gp7 ? "PASS" : $"FAIL ('{sessionRec?.OperatorId}')")}");
Console.WriteLine($"  Record UEC 100%:  {(gp9 ? "PASS" : $"FAIL ({sessionRec?.UEC_Percent})")}");
Console.WriteLine($"  Record MatrixSize:{(gp10 ? "PASS" : $"FAIL ('{sessionRec?.MatrixSize}')")}");
p2gPass = gp1 && gp2 && gp3 && gp4 && gp5 && gp6 && gp7 && gp8 && gp9 && gp10;
Console.WriteLine($"  DeviceSession e2e: {(p2gPass ? "PASS" : "FAIL")}");

// ── Phase 2 summary ───────────────────────────────────────────────────────────
bool p2Pass = p2aPass && p2bPass && p2cPass && p2dPass && p2ePass && p2fPass && p2gPass;
Console.WriteLine($"\nPhase 2 verification: {(p2Pass ? "PASS" : "FAIL")}");
if (!p2Pass)
{
    if (!p2aPass) Console.WriteLine("  FAIL: SanitizeForDmcc");
    if (!p2bPass) Console.WriteLine("  FAIL: DmccResponse.Parse");
    if (!p2cPass) Console.WriteLine("  FAIL: ClassifySymbology");
    if (!p2dPass) Console.WriteLine("  FAIL: 2D XML parse");
    if (!p2ePass) Console.WriteLine("  FAIL: 1D XML parse");
    if (!p2fPass) Console.WriteLine("  FAIL: DmccClient loopback");
    if (!p2gPass) Console.WriteLine("  FAIL: DeviceSession e2e");
}
Console.WriteLine("\nPhase 2 complete.");

