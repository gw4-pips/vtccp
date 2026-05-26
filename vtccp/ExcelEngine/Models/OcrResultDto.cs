namespace ExcelEngine.Models;

/// <summary>
/// Flat data-transfer object for an OCR result stored on <see cref="VerificationRecord"/>.
///
/// This type lives in ExcelEngine so that ExcelEngine does not need a project reference
/// to OcrEngine.  Command Pilot maps the rich <c>OcrEngine.OcrResult</c> record into
/// this DTO before storing it on the VerificationRecord.
///
/// All fields are nullable strings / numbers so the DTO can be trivially serialised
/// to JSON for the raw push XML sidecar archive (D4 scope).
/// </summary>
public sealed record OcrResultDto
{
    /// <summary>
    /// Text to report in the Excel row.
    /// Derived from the higher-confidence engine; null when both engines fail.
    /// </summary>
    public string? AgreedText { get; init; }

    /// <summary>
    /// Confidence tier as a display string: "High", "Medium", "Low", "Single", "Unreadable".
    /// </summary>
    public string? Tier { get; init; }

    /// <summary>
    /// Raw text from the Windows.Media.Ocr engine.  Null if engine was not run or failed.
    /// </summary>
    public string? WindowsText { get; init; }

    /// <summary>
    /// Average word confidence from Windows.Media.Ocr, 0–100.
    /// </summary>
    public double? WindowsConfidence { get; init; }

    /// <summary>
    /// Raw text from the Tesseract engine.  Null if engine was not run or failed.
    /// </summary>
    public string? TesseractText { get; init; }

    /// <summary>
    /// Mean confidence from Tesseract, 0–100.
    /// </summary>
    public double? TesseractConfidence { get; init; }

    /// <summary>
    /// Levenshtein edit distance between the two engine outputs.
    /// Null when fewer than two engines succeeded.
    /// </summary>
    public int? EditDistance { get; init; }

    /// <summary>
    /// Image level the OCR was applied to: "BarcodeCrop", "RoiFrame", or "FullSensorFrame".
    /// </summary>
    public string? ImageSource { get; init; }
}
