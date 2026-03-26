namespace ExcelEngine.Schema;

using ExcelEngine.Adapters;

/// <summary>
/// Writes a compact VTCCP schema-version metadata block into the Main worksheet.
///
/// Location: Row 1 (title row), columns immediately after the last schema data column.
/// This keeps the identifier out of all data rows and column ranges while staying
/// on the same row as the existing job title text.
///
/// Format (3 consecutive cells):
///   Col N+1: "VTCCP"
///   Col N+2: schema.Name   (e.g. "WebscanCompatible")
///   Col N+3: schema.Version (e.g. "1.0")
///
/// Purpose: allows downstream tooling to distinguish VTCCP-generated files from
/// legacy Webscan-generated files and from other Excel files.
/// </summary>
public static class SchemaVersionWriter
{
    public static void Write(IExcelAdapter adapter, ColumnSchema schema, int titleRow = 1)
    {
        int startCol = schema.Columns.Count + 2; // one gap after last data column
        adapter.WriteString(titleRow, startCol,     "VTCCP");
        adapter.WriteString(titleRow, startCol + 1, schema.Name);
        adapter.WriteString(titleRow, startCol + 2, schema.Version);
    }
}
