namespace ExcelEngine.Schema;

/// <summary>
/// A named, ordered collection of column definitions.
/// The column order determines the Excel column positions (left to right).
/// </summary>
public sealed class ColumnSchema
{
    public required string Name { get; init; }
    public string Version { get; init; } = "1.0";
    public string? Description { get; init; }
    public required IReadOnlyList<ColumnDefinition> Columns { get; init; }

    /// <summary>Returns the 1-based Excel column index for a given field id, or null if not present.</summary>
    public int? GetColumnIndex(string fieldId)
    {
        for (int i = 0; i < Columns.Count; i++)
        {
            if (Columns[i].FieldId == fieldId)
                return i + 1;
        }
        return null;
    }

    public ColumnDefinition? GetColumn(string fieldId) =>
        Columns.FirstOrDefault(c => c.FieldId == fieldId);
}
