namespace ExcelEngine.Utilities;

/// <summary>
/// Unified batch/lot extractor for <see cref="ExcelEngine.Models.BatchMode.AutoFromGS1"/>.
///
/// Tries each supported format parser in priority order and returns the first
/// non-null result.  If no parser finds a batch/lot value the method returns null,
/// which causes <see cref="ExcelEngine.Session.SessionManager"/> to fall back to
/// the caller-supplied <see cref="ExcelEngine.Models.VerificationRecord.BatchNumber"/>.
///
/// Supported formats (in order tried):
///   1. GS1 (FNC1-delimited symbols, Application Identifier 10)  — GS1Parser
///   2. ISO 15434 / ANSI MH10.8.2  ([)> envelope, DI 4L / 10L / 1L / L)
///      Covers MIL-STD-129 (PDF417 shipment labels) and MIL-STD-130 (Data Matrix UID)
///      — ISO15434Parser
/// </summary>
public static class AutoBatchExtractor
{
    /// <summary>
    /// Extracts the batch/lot value from a decoded barcode string using all
    /// known format parsers.  Returns null if none finds a batch value.
    /// </summary>
    public static string? ExtractBatchLot(string? decoded) =>
        GS1Parser.ExtractBatchLot(decoded)
        ?? ISO15434Parser.ExtractBatchLot(decoded);
}
