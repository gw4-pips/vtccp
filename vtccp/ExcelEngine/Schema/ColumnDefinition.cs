namespace ExcelEngine.Schema;

using ExcelEngine.Models;

/// <summary>
/// Defines a single column in the Excel output schema.
/// The field id links to a corresponding property or extraction path on VerificationRecord.
/// </summary>
public sealed class ColumnDefinition
{
    /// <summary>Unique field identifier — used to look up values from VerificationRecord.</summary>
    public required string FieldId { get; init; }

    /// <summary>Column header text shown in Excel row 2.</summary>
    public required string DisplayName { get; init; }

    /// <summary>Excel column width in character units (approximate).</summary>
    public double Width { get; init; } = 12;

    /// <summary>Which symbology group this column belongs to.</summary>
    public SymbologyGroup Group { get; init; } = SymbologyGroup.Universal;

    /// <summary>
    /// Which symbology families this column applies to.
    /// Empty list = applies to all symbologies.
    /// </summary>
    public IReadOnlyList<SymbologyFamily> ApplicableFamilies { get; init; } = [];

    /// <summary>Excel number format string. Null = General.</summary>
    public string? NumberFormat { get; init; }

    /// <summary>If true, column is always written even if value is null.</summary>
    public bool AlwaysInclude { get; init; } = true;
}
