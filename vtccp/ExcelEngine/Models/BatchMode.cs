namespace ExcelEngine.Models;

/// <summary>
/// Controls how the Batch/Lot field is populated on each VerificationRecord
/// written by SessionManager.
/// </summary>
public enum BatchMode
{
    /// <summary>
    /// Batch/Lot is supplied by the caller on each VerificationRecord.BatchNumber.
    /// SessionManager never overrides it.  Default.
    /// </summary>
    Manual,

    /// <summary>
    /// SessionManager automatically extracts GS1 AI(10) Batch/Lot from the
    /// record's decoded data string and stamps it on the Excel row.
    /// Falls back to VerificationRecord.BatchNumber if AI(10) is absent
    /// (e.g. non-GS1 symbols in a mixed session).
    /// </summary>
    AutoFromGS1,
}
