namespace DeviceInterface.Dmst;

using System.Globalization;
using ExcelEngine.Models;

/// <summary>
/// Cross-validates a push XML <see cref="VerificationRecord"/> against the
/// corresponding DMST HTML <see cref="DmstHtmlReport"/> for the same scan.
///
/// Two distinct operations per scan:
///
///   1. Supplemental merge — fields absent from push XML but present in the HTML
///      (QR_ECLevel, QR_MaskPattern, QR_ECI, ImagePolarity, DataCodewords,
///      ErrorCorrectionBudget, EncodedCharacters) are merged into the record and
///      catalogued in DataSourceExceptions.
///
///   2. Cross-validation — every field producible from BOTH sources is compared.
///      Mismatches → ValidationDiscrepancies + Debug log.
///      - Agreement: data-integrity credibility.
///      - Discrepancy: parser bug or firmware anomaly — Cognex bug-report candidate.
///
/// Confirmed from 2026-05-25 live HTML sample (QR GUID, fw 6.1.16_sr4):
///   ECLevel="M", DataMaskPattern="2", ECI="000003", ImagePolarity="Black on white"
///   DataCodewords=44, ErrorCorrectionBudget=26, EncodedCharacters=36(vs push 39).
///
/// EncodedCharacters: push XML eaLen is wrong (39 vs HTML authoritative 36).
/// HTML value is always taken as authoritative; discrepancy is always flagged.
/// </summary>
public static class DmstReportValidator
{
    /// <summary>
    /// Absolute tolerance for decimal field comparisons.
    /// Absorbs rounding differences between push XML and HTML rendering.
    /// </summary>
    public const decimal NumericTolerance = 0.05m;

    // ── Public entry point ────────────────────────────────────────────────────

    /// <summary>
    /// Merges supplemental fields from <paramref name="html"/> into
    /// <paramref name="record"/> and runs full cross-validation.
    /// Returns a new record instance with all merged and validated fields set.
    /// </summary>
    public static VerificationRecord MergeAndValidate(
        VerificationRecord record,
        DmstHtmlReport     html)
    {
        var exceptions    = new List<string>();
        var discrepancies = new List<string>();

        // Carry forward any existing DataSourceExceptions (e.g. from SYMBOL.RESULT FULL).
        if (!string.IsNullOrEmpty(record.DataSourceExceptions))
            exceptions.Add(record.DataSourceExceptions);

        // ── 1a. Merge: four permanently unresolvable from push XML ────────────

        string?       qrECLevel     = record.QR_ECLevel;
        string?       qrMaskPattern = record.QR_MaskPattern;
        string?       qrECI         = record.QR_ECI;
        ImagePolarity imagePolarity  = record.ImagePolarity;

        if (string.IsNullOrEmpty(qrECLevel) && !string.IsNullOrEmpty(html.ECLevel))
        {
            qrECLevel = html.ECLevel;
            exceptions.Add("QR_ECLevel:HtmlReport");
        }
        if (string.IsNullOrEmpty(qrMaskPattern) && !string.IsNullOrEmpty(html.DataMaskPattern))
        {
            qrMaskPattern = html.DataMaskPattern;
            exceptions.Add("QR_MaskPattern:HtmlReport");
        }
        if (string.IsNullOrEmpty(qrECI) && !string.IsNullOrEmpty(html.ECI))
        {
            qrECI = html.ECI;
            exceptions.Add("QR_ECI:HtmlReport");
        }
        if (imagePolarity == ImagePolarity.Unknown && !string.IsNullOrEmpty(html.ImagePolarity))
        {
            var parsed = ParseImagePolarity(html.ImagePolarity);
            if (parsed != ImagePolarity.Unknown)
            {
                imagePolarity = parsed;
                exceptions.Add("ImagePolarity:HtmlReport");
            }
        }

        // ── 1b. Merge: bonus fields (empty in push XML, present in HTML) ──────

        int? dataCodewords         = record.DataCodewords;
        int? errorCorrectionBudget = record.ErrorCorrectionBudget;
        int? encodedCharacters     = record.EncodedCharacters;

        if ((dataCodewords is null or 0) && html.DataCodewords.HasValue)
        {
            dataCodewords = html.DataCodewords;
            exceptions.Add("DataCodewords:HtmlReport");
        }
        if ((errorCorrectionBudget is null or 0) && html.ErrorCorrectionBudget.HasValue)
        {
            errorCorrectionBudget = html.ErrorCorrectionBudget;
            exceptions.Add("ErrorCorrectionBudget:HtmlReport");
        }
        // EncodedCharacters: push XML eaLen fallback is WRONG. HTML is authoritative.
        // Always take HTML value; always flag discrepancy if push differed.
        if (html.EncodedCharacters.HasValue)
        {
            if (encodedCharacters.HasValue && encodedCharacters != html.EncodedCharacters)
                discrepancies.Add(
                    $"EncodedCharacters:Push={encodedCharacters},Html={html.EncodedCharacters}");
            encodedCharacters = html.EncodedCharacters;
            exceptions.Add("EncodedCharacters:HtmlReport");
        }

        // ── 2. Cross-validation ───────────────────────────────────────────────

        CompareString ("OverallGrade",   record.OverallGrade?.LetterGradeString,
                                         html.OverallGrade,                      discrepancies);
        CompareString ("MatrixSize",     NormaliseMatrixSize(record.MatrixSize),
                                         NormaliseMatrixSize(html.MatrixSize),   discrepancies);
        // NominalXDim_2D (decimal, e.g. 12.6) vs HTML "12.6 mil" (parse first token).
        CompareDecimal("NomXDim_2D",     record.NominalXDim_2D,
                                         ParseNominalXDimHtml(html.NominalXDim), discrepancies);
        CompareDecimal("ANU%",           record.ANU_Percent,
                                         html.ANUPercent,                        discrepancies);
        CompareDecimal("GNU%",           record.GNU_Percent,
                                         html.GNUPercent,                        discrepancies);
        CompareDecimal("UEC%",           record.UEC_Percent,
                                         html.UECPercent,                        discrepancies);
        CompareDecimal("SC%",            record.SC_Percent,
                                         html.SCPercent,                         discrepancies);
        CompareDecimal("HBwg%",          record.HorizontalBWG,
                                         html.HorizontalBWG,                     discrepancies);
        CompareDecimal("VBwg%",          record.VerticalBWG,
                                         html.VerticalBWG,                       discrepancies);
        CompareInt    ("TotalCodewords", record.TotalCodewords,
                                         html.TotalCodewords,                    discrepancies);
        CompareInt    ("ErrorsCorrected",record.ErrorsCorrected,
                                         html.ErrorsCorrected,                   discrepancies);
        CompareInt    ("ErrCapUsed",     record.ErrorCapacityUsed,
                                         html.ErrorCapacityUsed,                 discrepancies);
        // ImagePolarity: compare push enum → canonical string vs HTML literal.
        CompareString ("ImagePolarity",  ImagePolarityToHtmlString(record.ImagePolarity),
                                         html.ImagePolarity,                     discrepancies);

        // Log.
        if (discrepancies.Count > 0)
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-VALID] {discrepancies.Count} discrepancy(s) — " +
                $"{record.VerificationDateTime:HH:mm:ss}:\n  " +
                string.Join("\n  ", discrepancies));
        else
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-VALID] All fields match — {record.VerificationDateTime:HH:mm:ss}.");

        // ── 3. Produce merged record ──────────────────────────────────────────
        return record with
        {
            QR_ECLevel            = qrECLevel,
            QR_MaskPattern        = qrMaskPattern,
            QR_ECI                = qrECI,
            ImagePolarity         = imagePolarity,
            DataCodewords         = dataCodewords,
            ErrorCorrectionBudget = errorCorrectionBudget,
            EncodedCharacters     = encodedCharacters,

            DataSourceExceptions    = exceptions.Count > 0
                                      ? string.Join(";", exceptions)
                                      : null,
            ValidationDiscrepancies = discrepancies.Count > 0
                                      ? string.Join(";", discrepancies)
                                      : null,
        };
    }

    // ── Comparison helpers ────────────────────────────────────────────────────

    private static void CompareString(
        string field, string? push, string? html, List<string> d)
    {
        if (push is null && html is null) return;
        if (string.Equals(push?.Trim(), html?.Trim(), StringComparison.OrdinalIgnoreCase)) return;
        d.Add($"{field}:Push={push ?? "null"},Html={html ?? "null"}");
    }

    private static void CompareDecimal(
        string field, decimal? push, decimal? html, List<string> d)
    {
        if (push is null && html is null) return;
        if (push is null || html is null)
        { d.Add($"{field}:Push={push?.ToString() ?? "null"},Html={html?.ToString() ?? "null"}"); return; }
        if (Math.Abs(push.Value - html.Value) <= NumericTolerance) return;
        d.Add($"{field}:Push={push},Html={html}");
    }

    private static void CompareInt(
        string field, int? push, int? html, List<string> d)
    {
        if (push is null && html is null) return;
        if (push == html) return;
        d.Add($"{field}:Push={push?.ToString() ?? "null"},Html={html?.ToString() ?? "null"}");
    }

    // ── Conversion helpers ────────────────────────────────────────────────────

    private static ImagePolarity ParseImagePolarity(string htmlValue)
        => htmlValue.Trim().ToLowerInvariant() switch
        {
            "black on white" => ImagePolarity.BlackOnWhite,
            "white on black" => ImagePolarity.WhiteOnBlack,
            _                => ImagePolarity.Unknown,
        };

    private static string? ImagePolarityToHtmlString(ImagePolarity p)
        => p switch
        {
            ImagePolarity.BlackOnWhite => "Black on white",
            ImagePolarity.WhiteOnBlack => "White on black",
            _                          => null,
        };

    private static string? NormaliseMatrixSize(string? raw)
        => raw?.Replace(" x ", "x", StringComparison.OrdinalIgnoreCase).Trim().ToLowerInvariant();

    /// <summary>
    /// Parses the numeric part of an HTML NominalXDim string such as "12.6 mil".
    /// Splits on the first space and parses the leading token as a decimal.
    /// Returns null if the string is absent or unparseable.
    /// </summary>
    private static decimal? ParseNominalXDimHtml(string? htmlValue)
    {
        if (string.IsNullOrEmpty(htmlValue)) return null;
        var token = htmlValue.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        return decimal.TryParse(token, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)
            ? v : null;
    }
}
