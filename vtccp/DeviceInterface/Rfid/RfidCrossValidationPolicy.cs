using ExcelEngine.Models;

namespace DeviceInterface.Rfid;

/// <summary>
/// Determines when an RFID scan adds value beyond the TruCheck result already
/// available for the barcode.
/// </summary>
public static class RfidCrossValidationPolicy
{
    /// <summary>
    /// Returns false when a GS1 DataMatrix has already passed TruCheck validation.
    /// In that case RFID cross-validation is redundant and must not be invoked or
    /// reported.
    /// </summary>
    public static bool ShouldRun(VerificationRecord record)
    {
        ArgumentNullException.ThrowIfNull(record);
        return !HasPassedTruCheckGs1Validation(record);
    }

    /// <summary>
    /// True when the GS1 DataMatrix record has a passing TruCheck application,
    /// grade, or correlated Data Format Check result.
    /// </summary>
    public static bool HasPassedTruCheckGs1Validation(VerificationRecord record)
    {
        ArgumentNullException.ThrowIfNull(record);

        bool isGs1DataMatrix =
            record.SymbologyFamily == SymbologyFamily.GS1DataMatrix ||
            string.Equals(record.Symbology, "GS1 DataMatrix", StringComparison.OrdinalIgnoreCase);
        if (!isGs1DataMatrix)
            return false;

        return string.Equals(record.ApplicationPass?.Trim(), "Pass", StringComparison.OrdinalIgnoreCase) ||
               record.OverallGrade?.PassFail == OverallPassFail.Pass ||
               record.HtmlDataFormatCheck?.Overall == OverallPassFail.Pass;
    }
}