namespace OcrEngine;

/// <summary>
/// Runs a JPEG image through two independent OCR engines (<see cref="WindowsOcrEngine"/>
/// and <see cref="TesseractOcrEngine"/>) in parallel, cross-validates the results,
/// and assigns a <see cref="OcrConfidenceTier"/>.
///
/// Cross-validation model:
///
///   HIGH       — both engines read; primary outputs agree exactly after normalisation
///   MEDIUM     — both engines read; edit distance ≤ 2 (single transposition / typo)
///   LOW        — both engines read; edit distance > 2 (materially different readings)
///   SINGLE     — only one engine produced output
///   UNREADABLE — neither engine produced output
///
/// For UPC/EAN symbologies (<see cref="HriSymbologyHint.UpcEan"/>):
///   • <see cref="BarcodeHriParser"/> extracts canonical digit strings from both
///     engine outputs, stripping inter-group spaces and the EAN right-margin '>'.
///   • Tier cross-validation and <see cref="OcrResult.AgreedText"/> are based on
///     the parsed digits, not the raw OCR text.
///   • <see cref="OcrResult.ParsedDigits"/> and <see cref="OcrResult.CheckDigitValid"/>
///     are populated; both are null for non-UPC/EAN symbologies.
///   • <see cref="OcrResult.WindowsEngineText"/> and <see cref="OcrResult.TesseractText"/>
///     still hold the original raw engine outputs for diagnostic purposes.
///
/// The agreed text for Excel output:
///   UPC/EAN with parsed digits → canonical digit string (Windows engine preferred)
///   HIGH / MEDIUM              → Windows engine text
///   LOW                        → Windows engine text (flagged by tier)
///   SINGLE                     → whichever engine succeeded
///   UNREADABLE                 → null
///
/// Intended usage from Command Pilot:
///   var runner = new DualEngineOcrRunner();
///   OcrResult result = await runner.RunAsync(
///       record.JpegImageBytes, OcrImageSource.BarcodeCrop, hint);
/// </summary>
public sealed class DualEngineOcrRunner
{
    private readonly IOcrEngine _windowsEngine;
    private readonly IOcrEngine _tesseractEngine;

    public DualEngineOcrRunner(
        IOcrEngine? windowsEngine    = null,
        IOcrEngine? tesseractEngine  = null)
    {
        _windowsEngine    = windowsEngine    ?? new WindowsOcrEngine();
        _tesseractEngine  = tesseractEngine  ?? new TesseractOcrEngine();
    }

    /// <summary>
    /// Runs both engines on <paramref name="jpegBytes"/> in parallel.
    ///
    /// Pass <paramref name="hint"/> = <see cref="HriSymbologyHint.UpcEan"/> for UPC-A,
    /// UPC-E, EAN-8, or EAN-13 scans to enable pattern-aware digit extraction and
    /// GS1 check-digit validation.
    ///
    /// Never throws — all exceptions are absorbed within the engine implementations.
    /// </summary>
    public async Task<OcrResult> RunAsync(
        byte[]            jpegBytes,
        OcrImageSource    imageSource = OcrImageSource.BarcodeCrop,
        HriSymbologyHint  hint        = HriSymbologyHint.None,
        CancellationToken ct          = default)
    {
        var winTask  = _windowsEngine.RecognizeAsync(jpegBytes, ct);
        var tessTask = _tesseractEngine.RecognizeAsync(jpegBytes, ct);

        await Task.WhenAll(winTask, tessTask).ConfigureAwait(false);

        var (winText,  winConf)  = winTask.Result;
        var (tessText, tessConf) = tessTask.Result;

        winText  = Normalise(winText);
        tessText = Normalise(tessText);

        // ── UPC/EAN pattern extraction ────────────────────────────────────────
        // For UPC and EAN symbologies the raw OCR output contains inter-group
        // spaces and possibly a right-margin '>' character.  Extract the canonical
        // digit string from each engine's output so that cross-validation operates
        // on the actual digit content rather than layout artefacts.

        HriParseResult? winParsed  = BarcodeHriParser.TryParse(winText,  hint);
        HriParseResult? tessParsed = BarcodeHriParser.TryParse(tessText, hint);

        // Primary comparison token: parsed digits when available, raw text otherwise.
        string? winPrimary  = winParsed?.Digits  ?? winText;
        string? tessPrimary = tessParsed?.Digits ?? tessText;

        bool winOk  = !string.IsNullOrEmpty(winPrimary);
        bool tessOk = !string.IsNullOrEmpty(tessPrimary);

        // Best parsed result — Windows engine is the preferred source.
        HriParseResult? bestParsed = winParsed ?? tessParsed;

        // ── Tier classification ───────────────────────────────────────────────

        if (!winOk && !tessOk)
        {
            return new OcrResult
            {
                Tier        = OcrConfidenceTier.Unreadable,
                ImageSource = imageSource,
            };
        }

        if (winOk && !tessOk)
        {
            return new OcrResult
            {
                AgreedText              = winParsed?.Digits ?? winText,
                Tier                    = OcrConfidenceTier.Single,
                WindowsEngineText       = winText,
                WindowsEngineConfidence = winConf,
                ImageSource             = imageSource,
                ParsedDigits            = winParsed?.Digits,
                CheckDigitValid         = winParsed?.CheckDigitValid,
            };
        }

        if (!winOk && tessOk)
        {
            return new OcrResult
            {
                AgreedText          = tessParsed?.Digits ?? tessText,
                Tier                = OcrConfidenceTier.Single,
                TesseractText       = tessText,
                TesseractConfidence = tessConf,
                ImageSource         = imageSource,
                ParsedDigits        = tessParsed?.Digits,
                CheckDigitValid     = tessParsed?.CheckDigitValid,
            };
        }

        int dist = LevenshteinDistance(winPrimary!, tessPrimary!);

        OcrConfidenceTier tier = dist == 0 ? OcrConfidenceTier.High
                               : dist <= 2 ? OcrConfidenceTier.Medium
                               :             OcrConfidenceTier.Low;

        return new OcrResult
        {
            AgreedText              = bestParsed?.Digits ?? winText,
            Tier                    = tier,
            WindowsEngineText       = winText,
            WindowsEngineConfidence = winConf,
            TesseractText           = tessText,
            TesseractConfidence     = tessConf,
            EditDistance            = dist,
            ImageSource             = imageSource,
            ParsedDigits            = bestParsed?.Digits,
            CheckDigitValid         = bestParsed?.CheckDigitValid,
        };
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string? Normalise(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        return System.Text.RegularExpressions.Regex
            .Replace(text.Trim(), @"\s+", " ")
            .ToUpperInvariant();
    }

    private static int LevenshteinDistance(string a, string b)
    {
        if (a == b)        return 0;
        if (a.Length == 0) return b.Length;
        if (b.Length == 0) return a.Length;

        int[] prev = Enumerable.Range(0, b.Length + 1).ToArray();
        int[] curr = new int[b.Length + 1];

        for (int i = 1; i <= a.Length; i++)
        {
            curr[0] = i;
            for (int j = 1; j <= b.Length; j++)
            {
                int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                curr[j] = Math.Min(
                    Math.Min(curr[j - 1] + 1, prev[j] + 1),
                    prev[j - 1] + cost);
            }
            Array.Copy(curr, prev, curr.Length);
        }
        return prev[b.Length];
    }
}
