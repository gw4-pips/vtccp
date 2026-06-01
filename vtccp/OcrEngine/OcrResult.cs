namespace OcrEngine;

/// <summary>
/// Result of a dual-engine OCR pass on a single verification scan image.
///
/// Populated by <see cref="DualEngineOcrRunner.RunAsync"/>.
/// Stored as a nullable property on <c>VerificationRecord</c> and written
/// to the Excel row when present.
/// </summary>
public sealed record OcrResult
{
    /// <summary>
    /// Agreed text to report in the Excel row and on the D1 verification report.
    ///
    /// Derivation:
    ///   UPC/EAN + parseable output → canonical digit string (spaces and '>' stripped)
    ///   HIGH / MEDIUM             → Windows engine text (primary)
    ///   LOW                       → Windows engine text (flagged by tier)
    ///   SINGLE                    → text from whichever engine succeeded
    ///   UNREADABLE                → null
    /// </summary>
    public string? AgreedText { get; init; }

    /// <summary>
    /// Cross-validation confidence tier.
    /// </summary>
    public OcrConfidenceTier Tier { get; init; }

    /// <summary>
    /// Raw text returned by Windows.Media.Ocr.  Null if the engine failed or was not run.
    /// </summary>
    public string? WindowsEngineText { get; init; }

    /// <summary>
    /// Per-word confidence scores from Windows.Media.Ocr (0–100 per word, averaged).
    /// Null if Windows engine did not produce a result.
    /// </summary>
    public double? WindowsEngineConfidence { get; init; }

    /// <summary>
    /// Raw text returned by Tesseract.  Null if the engine failed or was not run.
    /// </summary>
    public string? TesseractText { get; init; }

    /// <summary>
    /// Tesseract mean word confidence (0–100).  Null if Tesseract did not produce a result.
    /// </summary>
    public double? TesseractConfidence { get; init; }

    /// <summary>
    /// Character-level edit distance between the two engine primary outputs.
    /// For UPC/EAN this is computed on the digit-extracted strings, not raw text.
    /// Null when fewer than two engines produced a result.
    /// </summary>
    public int? EditDistance { get; init; }

    /// <summary>
    /// Which image level the OCR was applied to.
    /// </summary>
    public OcrImageSource ImageSource { get; init; }

    /// <summary>
    /// Digit-only string extracted from the OCR output by <see cref="BarcodeHriParser"/>.
    /// Populated only for UPC-A (12 digits), EAN-13 (13 digits), and EAN-8 (8 digits) scans.
    /// Null for all other symbologies or when the OCR output contained fewer than 6 digits.
    /// </summary>
    public string? ParsedDigits { get; init; }

    /// <summary>
    /// True when the GS1 modulo-10 check digit of <see cref="ParsedDigits"/> is valid.
    /// Null when <see cref="ParsedDigits"/> is null or the digit count is unrecognised.
    /// </summary>
    public bool? CheckDigitValid { get; init; }
}

/// <summary>
/// Identifies which image in the three-level stack was fed to the OCR pipeline.
/// </summary>
public enum OcrImageSource
{
    /// <summary>
    /// The barcode crop carried in push XML <c>r.trucheck.jpegImage</c>.
    /// This is the same image shown in the DMST verification panel.
    /// Approximately 200–600 px; tight crop around the decoded symbol.
    /// </summary>
    BarcodeCrop,

    /// <summary>
    /// The operator-configured ROI image retrieved via IMAGE.SEND after a live scan.
    /// Wider than the barcode crop; includes surrounding label area.
    /// Exact dimensions depend on the ROI setting in DMST.
    /// </summary>
    RoiFrame,

    /// <summary>
    /// The full sensor frame: 2448×2048 (DM475V / DM395V) or 2048×1536 (DM390 / DM394).
    /// Retrieved via <c>DataManSystem.GetLastReadImage()</c>.
    /// Highest resolution; includes the full field of view.
    /// </summary>
    FullSensorFrame,
}
