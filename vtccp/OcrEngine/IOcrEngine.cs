namespace OcrEngine;

/// <summary>
/// Abstraction over a single OCR engine.  Implementations:
///   <see cref="WindowsOcrEngine"/>  — Windows.Media.Ocr (Windows 10 / 11 built-in)
///   <see cref="TesseractOcrEngine"/> — Tesseract via NuGet (open-source, locally-run)
/// </summary>
public interface IOcrEngine
{
    /// <summary>
    /// Name used in logging and in <see cref="OcrResult"/> diagnostic fields.
    /// </summary>
    string EngineName { get; }

    /// <summary>
    /// Runs OCR on a raw JPEG byte array.
    ///
    /// Returns (text, averageConfidence) on success, or (null, null) on complete failure.
    /// Never throws — all exceptions are caught internally and result in a null return.
    ///
    /// <paramref name="jpegBytes"/> must be a valid JPEG image.
    /// </summary>
    Task<(string? Text, double? Confidence)> RecognizeAsync(
        byte[]            jpegBytes,
        CancellationToken ct = default);
}
