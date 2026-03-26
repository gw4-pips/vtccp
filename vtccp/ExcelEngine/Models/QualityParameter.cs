namespace ExcelEngine.Models;

/// <summary>
/// A single named quality parameter with its grading outcome.
/// Used for both 1D (ISO 15416) and 2D (ISO 15415) parameters.
/// </summary>
public sealed class QualityParameter
{
    /// <summary>Full parameter name, e.g. "Symbol Contrast"</summary>
    public required string Name { get; init; }

    /// <summary>Short abbreviation used in column headers, e.g. "SC"</summary>
    public required string Abbreviation { get; init; }

    /// <summary>The grading outcome for this parameter.</summary>
    public required GradingResult Result { get; init; }

    /// <summary>
    /// For SC: Rl/Rd values as a string, e.g. "89/4". Null for most params.
    /// </summary>
    public string? AuxiliaryValue { get; init; }
}
