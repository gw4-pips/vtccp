namespace OcrEngine;

/// <summary>
/// Runs a JPEG image through two independent OCR engines (<see cref="WindowsOcrEngine"/>
/// and <see cref="TesseractOcrEngine"/>) in parallel, cross-validates the results,
/// and assigns a <see cref="OcrConfidenceTier"/>.
///
/// Cross-validation model:
///
///   HIGH       — both engines read; outputs agree exactly after normalisation
///   MEDIUM     — both engines read; edit distance ≤ 2 (single transposition / typo)
///   LOW        — both engines read; edit distance > 2 (materially different readings)
///   SINGLE     — only one engine produced output
///   UNREADABLE — neither engine produced output
///
/// The agreed text for EXCEL output:
///   HIGH / MEDIUM  → Windows engine text (primary; higher baseline accuracy on label stock)
///   LOW            → Windows engine text (flagged)
///   SINGLE         → whichever engine succeeded
///   UNREADABLE     → null
///
/// Intended usage from Command Pilot:
///   var runner = new DualEngineOcrRunner();
///   OcrResult result = await runner.RunAsync(record.JpegImageBytes, OcrImageSource.BarcodeCrop);
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
    /// Never throws — all exceptions are absorbed within the engine implementations.
    /// </summary>
    public async Task<OcrResult> RunAsync(
        byte[]            jpegBytes,
        OcrImageSource    imageSource = OcrImageSource.BarcodeCrop,
        CancellationToken ct         = default)
    {
        var (winText, winConf) = (string?)null, (double?)null;
        var (tessText, tessConf) = (string?)null, (double?)null;

        var winTask  = _windowsEngine.RecognizeAsync(jpegBytes, ct);
        var tessTask = _tesseractEngine.RecognizeAsync(jpegBytes, ct);

        await Task.WhenAll(winTask, tessTask).ConfigureAwait(false);

        (winText,  winConf)  = winTask.Result;
        (tessText, tessConf) = tessTask.Result;

        winText  = Normalise(winText);
        tessText = Normalise(tessText);

        bool winOk  = !string.IsNullOrEmpty(winText);
        bool tessOk = !string.IsNullOrEmpty(tessText);

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
                AgreedText           = winText,
                Tier                 = OcrConfidenceTier.Single,
                WindowsEngineText    = winText,
                WindowsEngineConfidence = winConf,
                ImageSource          = imageSource,
            };
        }

        if (!winOk && tessOk)
        {
            return new OcrResult
            {
                AgreedText         = tessText,
                Tier               = OcrConfidenceTier.Single,
                TesseractText      = tessText,
                TesseractConfidence = tessConf,
                ImageSource        = imageSource,
            };
        }

        int dist = LevenshteinDistance(winText!, tessText!);

        OcrConfidenceTier tier = dist == 0  ? OcrConfidenceTier.High
                               : dist <= 2  ? OcrConfidenceTier.Medium
                               :              OcrConfidenceTier.Low;

        return new OcrResult
        {
            AgreedText              = winText,
            Tier                    = tier,
            WindowsEngineText       = winText,
            WindowsEngineConfidence = winConf,
            TesseractText           = tessText,
            TesseractConfidence     = tessConf,
            EditDistance            = dist,
            ImageSource             = imageSource,
        };
    }

    private static string? Normalise(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        return System.Text.RegularExpressions.Regex
            .Replace(text.Trim(), @"\s+", " ")
            .ToUpperInvariant();
    }

    private static int LevenshteinDistance(string a, string b)
    {
        if (a == b)   return 0;
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
