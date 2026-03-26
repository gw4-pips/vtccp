// VTCCP TestHarness -- Phase 1 scaffold verification
// Expanded in Tasks 2-4 with actual Excel writing and sample data.

using ExcelEngine.Models;
using ExcelEngine.Schema;

Console.WriteLine("VTCCP Excel Engine - Phase 1 Scaffold");
Console.WriteLine("======================================");

// --- Schema load and column count ---
var manager = new ColumnSchemaManager();
var schema = manager.GetActive();

Console.WriteLine($"Active schema: {schema.Name}");
Console.WriteLine($"Total columns: {schema.Columns.Count}");

var groups = schema.Columns
    .GroupBy(c => c.Group)
    .OrderBy(g => (int)g.Key);

foreach (var group in groups)
    Console.WriteLine($"  {group.Key}: {group.Count()} columns");

// --- Schema validation ---
var errors = ColumnSchemaManager.Validate(schema);
Console.WriteLine(errors.Count == 0
    ? "\nSchema validation: PASS -- no errors"
    : $"\nSchema validation: FAIL -- {errors.Count} error(s)");
foreach (var err in errors)
    Console.WriteLine($"  ! {err}");

// --- Schema round-trip: SaveToFile -> LoadFromFile ---
var tmpPath = Path.Combine(Path.GetTempPath(), "vtccp_schema_rt.json");
ColumnSchemaManager.SaveToFile(schema, tmpPath);
Console.WriteLine($"\nSchema saved to: {tmpPath}");

// Rename so it doesn't collide with the built-in schema name
var json = File.ReadAllText(tmpPath);
json = json.Replace($"\"Name\": \"{schema.Name}\"", "\"Name\": \"WebscanCompatible_RoundTrip\"");
var rtPath = Path.Combine(Path.GetTempPath(), "vtccp_schema_rt2.json");
File.WriteAllText(rtPath, json);

var manager2 = new ColumnSchemaManager();
manager2.LoadFromFile(rtPath);
manager2.SetActive("WebscanCompatible_RoundTrip");
var rt = manager2.GetActive();

Console.WriteLine($"Reloaded schema: {rt.Name}, columns: {rt.Columns.Count}");
Console.WriteLine(rt.Columns.Count == schema.Columns.Count
    ? "Schema round-trip: PASS"
    : $"Schema round-trip: FAIL (expected {schema.Columns.Count}, got {rt.Columns.Count})");

File.Delete(tmpPath);
File.Delete(rtPath);

// --- Sample VerificationRecord ---
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

// --- EPPlus license ---
// PRODUCTION NOTE: Replace with LicenseContext.Commercial when commercial EPPlus license is acquired.
OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
Console.WriteLine("\nEPPlus license context: NonCommercial (development only)");

Console.WriteLine("\nPhase 1 scaffold complete. Tasks 2-4 will implement Excel writing.");
