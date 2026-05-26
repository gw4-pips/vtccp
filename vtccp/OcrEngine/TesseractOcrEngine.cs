using Tesseract;

namespace OcrEngine;

/// <summary>
/// OCR engine backed by Tesseract via the <c>Tesseract</c> NuGet package.
/// Open-source; different internal algorithm to <see cref="WindowsOcrEngine"/>,
/// which is exactly what cross-validation requires.
///
/// Accuracy characteristics:
///   • Better than Windows.Media.Ocr on degraded or non-standard fonts when
///     a trained model is available; roughly equivalent on clean label stock
///   • Mean word confidence (0–100) is available from the iterator
///
/// Runtime requirement:
///   A <c>tessdata/</c> directory must be present next to the executable containing
///   at least <c>eng.traineddata</c>.
///
///   Download (development):
///     https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata
///   Place at:  {executable directory}/tessdata/eng.traineddata
///
///   The path can be overridden at construction via <paramref name="tessDataPath"/>.
///   In the deployed Command Pilot the tessdata folder is bundled into the installer.
/// </summary>
public sealed class TesseractOcrEngine : IOcrEngine
{
    private readonly string _tessDataPath;
    private readonly string _language;

    public string EngineName => "Tesseract";

    /// <param name="tessDataPath">
    /// Absolute or relative path to the tessdata directory.
    /// Defaults to <c>tessdata</c> relative to the calling assembly location.
    /// </param>
    /// <param name="language">
    /// Tesseract language code.  Default: "eng".
    /// </param>
    public TesseractOcrEngine(
        string? tessDataPath = null,
        string  language     = "eng")
    {
        _language     = language;
        _tessDataPath = tessDataPath
            ?? Path.Combine(
                AppContext.BaseDirectory,
                "tessdata");
    }

    public Task<(string? Text, double? Confidence)> RecognizeAsync(
        byte[]            jpegBytes,
        CancellationToken ct = default)
    {
        return Task.Run(() => RecognizeSync(jpegBytes), ct);
    }

    private (string? Text, double? Confidence) RecognizeSync(byte[] jpegBytes)
    {
        try
        {
            using var engine = new TesseractEngine(_tessDataPath, _language, EngineMode.Default);
            using var pix    = Pix.LoadFromMemory(jpegBytes);
            using var page   = engine.Process(pix);

            string text = page.GetText()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
                return (null, null);

            double confidence = page.GetMeanConfidence() * 100.0;
            return (text, confidence);
        }
        catch
        {
            return (null, null);
        }
    }
}
