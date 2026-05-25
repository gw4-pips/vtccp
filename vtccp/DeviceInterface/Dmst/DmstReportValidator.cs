namespace DeviceInterface.Dmst;

using ExcelEngine.Models;

/// <summary>
/// Cross-validates a push XML <see cref="VerificationRecord"/> against the
/// corresponding DMST HTML <see cref="DmstHtmlReport"/> for the same scan.
///
/// Two distinct outputs per scan:
///
///   1. <b>Supplemental merge</b> — fields absent from push XML but present in
///      the HTML report (QR_ECLevel, QR_MaskPattern, QR_ECI, ImagePolarity) are
///      written into the record and catalogued in DataSourceExceptions.
///
///   2. <b>Cross-validation</b> — every field producible from BOTH sources is
///      compared. Mismatches are recorded in ValidationDiscrepancies and logged.
///      Recurring patterns across multiple scans indicate parser bugs or firmware
///      anomalies — both categories are candidates for Cognex bug reports.
///
/// Numeric comparisons use a configurable tolerance (default: 0.05) to absorb
/// rounding differences between push XML decimal formatting and HTML rendering.
///
/// "Set before scan, clear after" model:
///   The caller (DeviceSession scan loop) is responsible for arming
///   DmstHtmlScraper before each trigger and disarming it after the HTML is
///   scraped. This validator is stateless and operates on already-parsed data.
/// </summary>
public static class DmstReportValidator
{
    /// <summary>
    /// Absolute tolerance for decimal field comparisons (push vs HTML).
    /// Covers rounding differences in how the firmware formats numbers in
    /// the JS push output vs. the HTML report renderer.
    /// </summary>
    public const decimal NumericTolerance = 0.05m;

    // ── Public entry point ────────────────────────────────────────────────────

    /// <summary>
    /// Merges supplemental fields from <paramref name="html"/> into
    /// <paramref name="record"/> and runs cross-validation.
    ///
    /// The returned record is a new instance with:
    ///   - QR_ECLevel, QR_MaskPattern, QR_ECI, ImagePolarity populated from HTML
    ///     (if absent from push XML and present in HTML)
    ///   - DataSourceExceptions updated with source tags for merged fields
    ///   - ValidationDiscrepancies populated if any overlapping field disagrees
    ///
    /// Cross-validation runs unconditionally — even if SYMBOL.RESULT FULL
    /// eventually provides some fields, we still compare all overlapping push XML
    /// values against the HTML report for every scan. Discrepancies surface parser
    /// bugs or firmware anomalies; agreement builds data-integrity credibility.
    /// </summary>
    public static VerificationRecord MergeAndValidate(
        VerificationRecord record,
        DmstHtmlReport     html)
    {
        var exceptions    = new List<string>();
        var discrepancies = new List<string>();

        // ── 1. Supplemental merge ─────────────────────────────────────────────

        // Carry forward any existing DataSourceExceptions (e.g. from SYMBOL.RESULT FULL).
        if (!string.IsNullOrEmpty(record.DataSourceExceptions))
            exceptions.Add(record.DataSourceExceptions);

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
            imagePolarity = ParseImagePolarity(html.ImagePolarity);
            if (imagePolarity != ImagePolarity.Unknown)
                exceptions.Add("ImagePolarity:HtmlReport");
        }

        // ── 2. Cross-validation ───────────────────────────────────────────────
        // Compare every field whose value comes from BOTH push XML and HTML report.
        // Null-vs-null is silently skipped (no data from either side = no comparison).

        CompareString ("OverallGrade",   record.OverallGrade?.Letter.ToString(),
                                         html.OverallGrade,                        discrepancies);
        CompareString ("DecodedData",    record.DecodedData,
                                         html.DecodedData,                         discrepancies);
        CompareString ("MatrixSize",     NormaliseMatrixSize(record.MatrixSize),
                                         NormaliseMatrixSize(html.MatrixSize),     discrepancies);
        CompareString ("ContrastUnif",   record.ContrastUniformity,
                                         html.ContrastUniformity,                  discrepancies);
        CompareString ("MRD",            record.MRD,
                                         html.MRD,                                 discrepancies);
        CompareString ("ErrCorrType",    record.ErrorCorrectionType,
                                         html.ErrorCorrectionType,                 discrepancies);
        CompareDecimal("ANU%",           record.ANU_Percent,
                                         html.ANUPercent,                          discrepancies);
        CompareDecimal("GNU%",           record.GNU_Percent,
                                         html.GNUPercent,                          discrepancies);
        CompareDecimal("UEC%",           record.UEC_Percent,
                                         html.UECPercent,                          discrepancies);
        CompareDecimal("SC%",            record.SC_Percent,
                                         html.SCPercent,                           discrepancies);
        CompareDecimal("FPD",            record.FPD_Value,
                                         html.FPDValue,                            discrepancies);
        CompareDecimal("HBwg%",          record.HorizontalBWG,
                                         html.HorizontalBWG,                       discrepancies);
        CompareDecimal("VBwg%",          record.VerticalBWG,
                                         html.VerticalBWG,                         discrepancies);
        // ImagePolarity: compare push enum as canonical string vs HTML literal string
        CompareString ("ImagePolarity",  ImagePolarityToHtmlString(record.ImagePolarity),
                                         html.ImagePolarity,                       discrepancies);

        // Log result.
        if (discrepancies.Count > 0)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-VALID] {discrepancies.Count} discrepancy(s) — scan " +
                $"{record.VerificationDateTime:HH:mm:ss}:\n  " +
                string.Join("\n  ", discrepancies));
        }
        else
        {
            System.Diagnostics.Debug.WriteLine(
                $"[VTCCP-VALID] All overlapping fields match — scan " +
                $"{record.VerificationDateTime:HH:mm:ss}.");
        }

        // ── 3. Produce merged record ──────────────────────────────────────────
        return record with
        {
            QR_ECLevel      = qrECLevel,
            QR_MaskPattern  = qrMaskPattern,
            QR_ECI          = qrECI,
            ImagePolarity   = imagePolarity,

            DataSourceExceptions  = exceptions.Count > 0
                                    ? string.Join(";", exceptions)
                                    : null,
            ValidationDiscrepancies = discrepancies.Count > 0
                                    ? string.Join(";", discrepancies)
                                    : null,
        };
    }

    // ── Comparison helpers ────────────────────────────────────────────────────

    private static void CompareString(
        string       field,
        string?      push,
        string?      html,
        List<string> discrepancies)
    {
        // Both absent → no data from either side; skip silently.
        if (push is null && html is null) return;
        if (string.Equals(push?.Trim(), html?.Trim(), StringComparison.OrdinalIgnoreCase)) return;
        discrepancies.Add($"{field}:Push={push ?? "null"},Html={html ?? "null"}");
    }

    private static void CompareDecimal(
        string       field,
        decimal?     push,
        decimal?     html,
        List<string> discrepancies)
    {
        if (push is null && html is null) return;
        if (push is null || html is null)
        {
            discrepancies.Add($"{field}:Push={push?.ToString() ?? "null"},Html={html?.ToString() ?? "null"}");
            return;
        }
        if (Math.Abs(push.Value - html.Value) <= NumericTolerance) return;
        discrepancies.Add($"{field}:Push={push},Html={html}");
    }

    // ── Conversion helpers ────────────────────────────────────────────────────

    /// <summary>
    /// Parses the HTML report's image polarity string to the VerificationRecord enum.
    /// "Normal" → BlackOnWhite (dark marks on light background)
    /// "Inverted" → WhiteOnBlack (light marks on dark background)
    /// Anything else → Unknown
    /// </summary>
    private static ImagePolarity ParseImagePolarity(string htmlValue)
        => htmlValue.Trim().ToLowerInvariant() switch
        {
            "normal"   => ImagePolarity.BlackOnWhite,
            "inverted" => ImagePolarity.WhiteOnBlack,
            _          => ImagePolarity.Unknown,
        };

    /// <summary>
    /// Converts the VerificationRecord ImagePolarity enum back to the HTML report
    /// string for cross-validation comparison.
    /// </summary>
    private static string? ImagePolarityToHtmlString(ImagePolarity polarity)
        => polarity switch
        {
            ImagePolarity.BlackOnWhite => "Normal",
            ImagePolarity.WhiteOnBlack => "Inverted",
            _                          => null,   // Unknown → skip comparison
        };

    /// <summary>
    /// Normalises matrix size strings for comparison.
    /// "16 x 36" → "16x36"; "22x22 (Data: 20x20)" → "22x22 (data: 20x20)".
    /// </summary>
    private static string? NormaliseMatrixSize(string? raw)
        => raw?.Replace(" x ", "x", StringComparison.OrdinalIgnoreCase)
               .Trim()
               .ToLowerInvariant();
}
