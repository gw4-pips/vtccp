// VTCCP TestHarness -- Phase 1 Tasks 1+2
// Writes sample 2D Data Matrix verification records to both .xlsx and .xls

using ExcelEngine.Adapters;
using ExcelEngine.Models;
using ExcelEngine.Schema;
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
    Console.WriteLine($"  Columns match schema: {(headerCols == schema.Columns.Count ? "YES" : $"NO ({headerCols} vs {schema.Columns.Count})")}");
}

// ─── File naming test ──────────────────────────────────────────────────────────
Console.WriteLine("\nFile naming tests:");
Console.WriteLine($"  Xlsx name: {ExcelFileManager.GenerateFileName(session, OutputFormat.Xlsx)}");
Console.WriteLine($"  Xls name:  {ExcelFileManager.GenerateFileName(session, OutputFormat.Xls)}");
var noJob = new SessionState { OperatorId = "GW4", SessionStarted = new DateTime(2025, 8, 11) };
Console.WriteLine($"  No-job xlsx: {ExcelFileManager.GenerateFileName(noJob, OutputFormat.Xlsx)}");
var noJobNoOp = new SessionState { SessionStarted = new DateTime(2025, 8, 11) };
Console.WriteLine($"  Fallback:    {ExcelFileManager.GenerateFileName(noJobNoOp, OutputFormat.Xlsx)}");

Console.WriteLine("\nTask 2 complete.");
