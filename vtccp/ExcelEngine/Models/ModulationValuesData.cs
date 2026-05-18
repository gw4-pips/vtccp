namespace ExcelEngine.Models;

/// <summary>
/// Raw modulation-array data extracted from q.modulationArray in the firmware push object.
/// The firmware array has (rows+2)×(cols+2) elements — it includes the 1-module quiet-zone
/// border on all four sides.  For a 16×36 symbol, GridRows=18 and GridCols=38 (confirmed
/// v1.28 device scan: modulationArray.length=684=18×38 ✓).
///
/// Written to the "Modulation Values" worksheet by ModValuesSheetWriter.
/// Attached to VerificationRecord.ModulationValues — populated by the push parser.
/// </summary>
public sealed class ModulationValuesData
{
    /// <summary>
    /// Number of rows in the firmware grid, including the 1-module QZ border on each side.
    /// SymbolRows = GridRows − 2.
    /// </summary>
    public required int GridRows { get; init; }

    /// <summary>
    /// Number of columns in the firmware grid, including the 1-module QZ border on each side.
    /// SymbolCols = GridCols − 2.
    /// </summary>
    public required int GridCols { get; init; }

    /// <summary>
    /// Raw reflectance value per module, 0–255 scale as returned by the firmware.
    /// Indexed row-major: index = row × GridCols + col (both 0-based).
    /// Length must equal GridRows × GridCols.
    /// </summary>
    public required IReadOnlyList<int> Raw { get; init; }

    /// <summary>
    /// IsBlack flag per module.  True = dark module, false = light module.
    /// Same length and indexing as <see cref="Raw"/>.
    /// </summary>
    public required IReadOnlyList<bool> IsBlack { get; init; }

    /// <summary>
    /// Grade character per module (single char, e.g. "(").
    /// Confirmed v1.28: grade="(" is the complete value (gradeLen=1, not truncation).
    /// Same length and indexing as <see cref="Raw"/>.
    /// May be shorter than Raw if the firmware omitted grade for some elements.
    /// </summary>
    public IReadOnlyList<string> Grade { get; init; } = [];

    /// <summary>
    /// Symbol rows excluding the quiet-zone border (= GridRows − 2).
    /// For a 16×36 symbol this is 16.
    /// </summary>
    public int SymbolRows => GridRows - 2;

    /// <summary>
    /// Symbol columns excluding the quiet-zone border (= GridCols − 2).
    /// For a 16×36 symbol this is 36.
    /// </summary>
    public int SymbolCols => GridCols - 2;

    /// <summary>Timestamp of the scan (from VerificationRecord.VerificationDateTime).</summary>
    public DateTime ScanDateTime { get; init; }

    /// <summary>
    /// Human-readable label for the sheet section header.
    /// E.g. "2026-05-18 11:24:40  |  Data Matrix  |  A  |  (010)..."
    /// </summary>
    public string? Label { get; init; }
}
