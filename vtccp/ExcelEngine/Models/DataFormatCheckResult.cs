namespace ExcelEngine.Models;

/// <summary>
/// Result of GS1 Application Identifier data format check or UID/MIL-STD-130 format check.
/// Corresponds to the GS1 Data Format Check (DFC) section of DMV TruCheck verification reports.
/// </summary>
public sealed class DataFormatCheckResult
{
    /// <summary>Overall pass/fail for the data format check.</summary>
    public OverallPassFail Overall { get; init; } = OverallPassFail.NotApplicable;

    /// <summary>Format standard applied, e.g. "GS1 Application Data Format" or "MIL-STD-130"</summary>
    public string? Standard { get; init; }

    /// <summary>Parsed AI/field table rows. Each row: Name, Data, Check (PASS/FAIL).</summary>
    public IReadOnlyList<DataFormatCheckRow> Rows { get; init; } = [];
}

public sealed class DataFormatCheckRow
{
    public required string Name { get; init; }
    public required string Data { get; init; }
    public required string Check { get; init; }
}
