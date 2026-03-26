namespace ExcelEngine.Models;

/// <summary>
/// Element width data from 1D barcode verification.
/// Maps to the "Element Widths" second worksheet in the Webscan format.
///
/// UPC/EAN layout (from Webscan reference PDFs):
///   Columns: CHAR | SPACE/BAR/SPACE (left guard) | SPACE/BAR/SPACE (center) | SPACE/BAR/SPACE (right guard)
///   Rows: LGB (left guard bar), digit rows 0-5, CGB (center guard bar),
///         digit rows 6-last, RGB (right guard bar)
///
/// The same row/column structure is used for both Element Sizes and Element Deviations tables.
/// A null cell value means "not applicable" for that element/column combination.
/// </summary>
public sealed class ElementWidthData
{
    /// <summary>
    /// Column header labels for the element table (e.g. "CHAR", "SPACE", "BAR", "SPACE", ...).
    /// Webscan UPC-A/EAN-13: typically 7 columns: CHAR, SP/BAR/SP (left), SP/BAR/SP (center), SP/BAR/SP (right)
    /// </summary>
    public IReadOnlyList<string> ColumnHeaders { get; init; } = [];

    /// <summary>Rows for the Element Sizes table (measured widths in mils or modules).</summary>
    public IReadOnlyList<ElementWidthRow> ElementSizes { get; init; } = [];

    /// <summary>Rows for the Element Deviations table (deviations from nominal).</summary>
    public IReadOnlyList<ElementWidthRow> ElementDeviations { get; init; } = [];

    /// <summary>Record identifier for the header line on the Element Widths sheet (e.g. "UPC-A | 2025-08-11 | Scan 1").</summary>
    public string? RecordLabel { get; init; }
}

/// <summary>
/// A single element row in the Element Sizes or Element Deviations table.
/// ElementName is the row label (e.g. "LGB", "0", "CGB", "RGB").
/// Values correspond to the ElementWidthData.ColumnHeaders columns.
/// </summary>
public sealed class ElementWidthRow
{
    public required string ElementName { get; init; }

    /// <summary>
    /// Cell values in column order matching ElementWidthData.ColumnHeaders.
    /// A null entry means "not applicable / blank" for that column.
    /// </summary>
    public IReadOnlyList<decimal?> Values { get; init; } = [];
}
