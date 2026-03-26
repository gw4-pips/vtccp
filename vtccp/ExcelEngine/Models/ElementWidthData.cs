namespace ExcelEngine.Models;

/// <summary>
/// Element width data from 1D barcode verification.
/// Maps to the "Element Widths" second worksheet in Webscan format.
/// Contains element sizes and deviations for each bar/space in the symbol.
/// </summary>
public sealed class ElementWidthData
{
    public IReadOnlyList<ElementSizeRow> ElementSizes { get; init; } = [];
    public IReadOnlyList<ElementDeviationRow> ElementDeviations { get; init; } = [];
}

public sealed class ElementSizeRow
{
    public required string ElementName { get; init; }
    public decimal? NominalMils { get; init; }
    public decimal? MeasuredMils { get; init; }
}

public sealed class ElementDeviationRow
{
    public required string ElementName { get; init; }
    public decimal? DeviationMils { get; init; }
    public decimal? DeviationPercent { get; init; }
}
