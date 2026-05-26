namespace OcrEngine;

/// <summary>
/// Confidence tier assigned by <see cref="DualEngineOcrRunner"/> after cross-validating
/// the Windows.Media.Ocr and Tesseract engine results.
///
/// Tier assignment rules:
///   HIGH       — both engines read successfully and their outputs agree exactly (after normalisation)
///   MEDIUM     — both engines read, outputs differ by ≤2 characters (edit distance ≤ 2)
///   LOW        — both engines read, outputs differ materially (edit distance > 2)
///   SINGLE     — exactly one engine produced a result; the other failed entirely
///   UNREADABLE — both engines failed to produce any output
/// </summary>
public enum OcrConfidenceTier
{
    /// <summary>Both engines agree exactly.</summary>
    High,

    /// <summary>Both engines read; outputs differ by ≤2 characters.</summary>
    Medium,

    /// <summary>Both engines read; outputs differ materially (edit distance > 2).</summary>
    Low,

    /// <summary>One engine succeeded; the other produced no output.</summary>
    Single,

    /// <summary>Both engines produced no output.</summary>
    Unreadable,
}
