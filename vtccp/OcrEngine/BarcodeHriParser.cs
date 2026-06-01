namespace OcrEngine;

using System.Text;

/// <summary>
/// Pattern-aware HRI (Human Readable Interpretation) parser for UPC and EAN symbologies.
///
/// Physical layout of HRI digits under/beside a barcode:
///
///   UPC-A  (12 digits)  — pattern 1-5-5-1
///     • 1 digit  : number-system digit, printed OUTSIDE the left descender bar
///     • 5 digits : left half-bar group
///     • 5 digits : right half-bar group
///     • 1 digit  : check digit, printed OUTSIDE the right descender bar
///
///   EAN-13 (13 digits)  — pattern 1-6-6 or 1-6-6->
///     • 1 digit  : number-system digit, printed OUTSIDE / LEFT of the left guard
///     • 6 digits : left half-bar group
///     • 6 digits : right half-bar group
///     • '>'      : optional quiet-zone protector to the right of the right guard
///
///   EAN-8  (8 digits)   — pattern 4-4
///     • 4 digits : left half
///     • 4 digits : right half
///
/// OCR engines frequently insert spaces between groups (reflecting the physical inter-
/// group gaps) and may capture the EAN right-margin '>'.  This parser normalises those
/// artefacts and validates the GS1 modulo-10 check digit.
/// </summary>
public static class BarcodeHriParser
{
    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>
    /// Strips whitespace, '>' and all non-digit characters from <paramref name="raw"/>.
    /// Returns the digit-only string, or <see cref="string.Empty"/> when nothing remains.
    /// </summary>
    public static string ExtractDigits(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        var sb = new StringBuilder(raw.Length);
        foreach (char c in raw)
            if (char.IsAsciiDigit(c)) sb.Append(c);
        return sb.ToString();
    }

    /// <summary>
    /// Attempts to parse an OCR output string as a UPC/EAN HRI value.
    ///
    /// Recognises UPC-A (12 digits), EAN-13 (13 digits), and EAN-8 (8 digits).
    /// UPC-E expansion is not performed in v1; 6-digit compressed form returns
    /// <see cref="HriParseResult.PatternMatch"/> = false.
    ///
    /// Returns <c>null</c> when:
    ///   • <paramref name="hint"/> is <see cref="HriSymbologyHint.None"/>, or
    ///   • fewer than 6 digit characters are present in the OCR output.
    /// </summary>
    public static HriParseResult? TryParse(string? rawOcr, HriSymbologyHint hint)
    {
        if (hint == HriSymbologyHint.None) return null;

        string digits = ExtractDigits(rawOcr);

        return digits.Length switch
        {
            12 => new HriParseResult(digits, ValidateGs1CheckDigit(digits), PatternMatch: true),
            13 => new HriParseResult(digits, ValidateGs1CheckDigit(digits), PatternMatch: true),
            8  => new HriParseResult(digits, ValidateGs1CheckDigit(digits), PatternMatch: true),
            >= 6 => new HriParseResult(digits, CheckDigitValid: false,       PatternMatch: false),
            _    => null
        };
    }

    /// <summary>
    /// Validates the GS1 modulo-10 check digit for UPC-A (12 digits),
    /// EAN-13 (13 digits), and EAN-8 (8 digits).
    ///
    /// Algorithm: sum each digit multiplied by an alternating weight (3 or 1),
    /// where the starting weight is determined by the total digit count so that
    /// both UPC-A and EAN-13 use the correct GS1 phase.  The check digit (last)
    /// must bring the total to a multiple of 10.
    ///
    ///   even length (UPC-A=12, EAN-8=8) : i=0 weight=3, i=1 weight=1, …
    ///   odd  length (EAN-13=13)          : i=0 weight=1, i=1 weight=3, …
    /// </summary>
    public static bool ValidateGs1CheckDigit(string digits)
    {
        if (digits.Length < 2) return false;

        int n   = digits.Length;
        int sum = 0;

        for (int i = 0; i < n - 1; i++)
        {
            if (!char.IsAsciiDigit(digits[i])) return false;
            int d          = digits[i] - '0';
            int multiplier = ((n + i) % 2 == 0) ? 3 : 1;
            sum           += d * multiplier;
        }

        if (!char.IsAsciiDigit(digits[n - 1])) return false;
        int expected = (10 - (sum % 10)) % 10;
        return digits[n - 1] - '0' == expected;
    }
}

// ── Supporting types ──────────────────────────────────────────────────────────

/// <summary>
/// Result of a <see cref="BarcodeHriParser.TryParse"/> call.
/// </summary>
/// <param name="Digits">
/// Digit-only string extracted from the raw OCR output (spaces, '>' and other
/// non-digit characters removed).
/// </param>
/// <param name="CheckDigitValid">
/// <c>true</c> when the GS1 modulo-10 check digit passes.
/// <c>false</c> when the digit count is unrecognised or the check fails.
/// </param>
/// <param name="PatternMatch">
/// <c>true</c> when the digit count matches a known UPC/EAN variant (8, 12, or 13).
/// <c>false</c> when digits were recovered but in an unexpected count.
/// </param>
public sealed record HriParseResult(
    string Digits,
    bool   CheckDigitValid,
    bool   PatternMatch);

/// <summary>
/// Indicates which symbology family is being scanned so <see cref="BarcodeHriParser"/>
/// can apply the correct HRI pattern.
/// </summary>
public enum HriSymbologyHint
{
    /// <summary>
    /// Not a UPC or EAN barcode; HRI pattern parsing does not apply.
    /// The OCR engines run normally but no structured digit extraction is attempted.
    /// </summary>
    None,

    /// <summary>
    /// UPC-A, UPC-E, EAN-8, or EAN-13.
    /// The parser auto-selects the variant based on recovered digit count.
    /// </summary>
    UpcEan,
}
