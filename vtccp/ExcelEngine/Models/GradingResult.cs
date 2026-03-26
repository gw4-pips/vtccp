namespace ExcelEngine.Models;

/// <summary>
/// Represents the grading outcome for a single quality parameter.
/// Numeric grade is to one decimal place per ISO 15415:2024 (41-band scale).
/// Legacy integer grades (4.0, 3.0, 2.0, 1.0, 0.0) are fully supported.
/// </summary>
public sealed class GradingResult
{
    /// <summary>Numeric grade 0.0–4.0 to one decimal place. Null if not measured.</summary>
    public decimal? NumericGrade { get; init; }

    /// <summary>Letter grade A/B/C/D/F. Null if not measured.</summary>
    public GradeLetterValue? LetterGrade { get; init; }

    /// <summary>PASS / FAIL / N/A</summary>
    public OverallPassFail PassFail { get; init; } = OverallPassFail.NotApplicable;

    /// <summary>Raw measured value (e.g. "84%" for SC, "0.2%" for ANU). Null if not applicable.</summary>
    public string? MeasuredValue { get; init; }

    /// <summary>
    /// Returns letter grade string: "A", "B", "C", "D", "F", or empty.
    /// </summary>
    public string LetterGradeString => LetterGrade switch
    {
        GradeLetterValue.A => "A",
        GradeLetterValue.B => "B",
        GradeLetterValue.C => "C",
        GradeLetterValue.D => "D",
        GradeLetterValue.F => "F",
        _ => string.Empty,
    };

    /// <summary>
    /// Returns numeric grade formatted to one decimal place, or empty string.
    /// </summary>
    public string NumericGradeString => NumericGrade.HasValue
        ? NumericGrade.Value.ToString("0.0")
        : string.Empty;

    /// <summary>
    /// Returns "PASS", "FAIL", or empty string.
    /// </summary>
    public string PassFailString => PassFail switch
    {
        OverallPassFail.Pass => "PASS",
        OverallPassFail.Fail => "FAIL",
        _ => string.Empty,
    };

    public static GradingResult NotMeasured => new() { PassFail = OverallPassFail.NotApplicable };

    public static GradingResult FromLetterAndNumeric(string letter, decimal numeric, string passFail, string? value = null)
    {
        var letterGrade = letter.Trim().ToUpper() switch
        {
            "A" => GradeLetterValue.A,
            "B" => GradeLetterValue.B,
            "C" => GradeLetterValue.C,
            "D" => GradeLetterValue.D,
            "F" => GradeLetterValue.F,
            _ => GradeLetterValue.NotApplicable,
        };

        var pf = passFail.Trim().ToUpper() switch
        {
            "PASS" => OverallPassFail.Pass,
            "FAIL" => OverallPassFail.Fail,
            _ => OverallPassFail.NotApplicable,
        };

        return new GradingResult
        {
            NumericGrade = numeric,
            LetterGrade = letterGrade,
            PassFail = pf,
            MeasuredValue = value,
        };
    }
}
