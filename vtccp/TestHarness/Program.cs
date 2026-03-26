// VTCCP TestHarness — Phase 1 scaffold verification
// This stub will be expanded in Tasks 2-4 with actual Excel writing and sample data.

using ExcelEngine.Models;
using ExcelEngine.Schema;

Console.WriteLine("VTCCP Excel Engine - Phase 1 Scaffold");
Console.WriteLine("======================================");

// Verify the column schema manager initializes correctly
var manager = new ColumnSchemaManager();
var schema = manager.GetActive();

Console.WriteLine($"Active schema: {schema.Name}");
Console.WriteLine($"Total columns: {schema.Columns.Count}");

// Print column block summary
var groups = schema.Columns
    .GroupBy(c => c.Group)
    .OrderBy(g => (int)g.Key);

foreach (var group in groups)
{
    Console.WriteLine($"  {group.Key}: {group.Count()} columns");
}

// Validate the schema
var errors = ColumnSchemaManager.Validate(schema);
if (errors.Count == 0)
{
    Console.WriteLine("\nSchema validation: PASS -- no errors");
}
else
{
    Console.WriteLine($"\nSchema validation: FAIL -- {errors.Count} error(s)");
    foreach (var err in errors)
        Console.WriteLine($"  ! {err}");
}

// Verify a sample VerificationRecord constructs cleanly
var record = new VerificationRecord
{
    Symbology = "GS1 DataMatrix",
    SymbologyFamily = SymbologyFamily.GS1DataMatrix,
    DecodedData = "<F1>010123456789012810GRADE-4-A-AI-INC",
    FormalGrade = "4.0/16/660/45Q",
    MatrixSize = "22x22 (Data: 20x20)",
    OverallGrade = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS"),
    UEC_Percent = 100,
    UEC_Grade = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS", "100%"),
    SC_Percent = 84,
    SC_RlRd = "89/4",
    SC_Grade = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS", "84%"),
    ANU_Percent = 0.1m,
    ANU_Grade = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS", "0.1%"),
    GNU_Percent = 2.3m,
    GNU_Grade = GradingResult.FromLetterAndNumeric("A", 4.0m, "PASS", "2.3%"),
    HorizontalBWG = -11,
    VerticalBWG = -11,
    NominalXDim_2D = 19.8m,
};

Console.WriteLine($"\nSample record: {record.Symbology} -- {record.FormalGrade}");
Console.WriteLine($"  IsLargeMatrix: {record.IsLargeMatrix}");
Console.WriteLine($"  Is2D: {record.Is2D}");
Console.WriteLine($"  UEC: {record.UEC_Grade?.NumericGradeString} ({record.UEC_Grade?.LetterGradeString}) {record.UEC_Grade?.PassFailString}");

// EPPlus license context (NonCommercial -- flag for production licensing review)
// PRODUCTION NOTE: Replace with ExcelPackage.LicenseContext = LicenseContext.Commercial
//                  when a commercial EPPlus license is acquired.
OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
Console.WriteLine("\nEPPlus license context: NonCommercial (development only)");

Console.WriteLine("\nPhase 1 scaffold complete. Tasks 2-4 will implement Excel writing.");
