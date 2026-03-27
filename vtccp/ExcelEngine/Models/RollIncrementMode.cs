namespace ExcelEngine.Models;

/// <summary>
/// Controls how the roll identifier is determined and advanced within a session.
/// </summary>
public enum RollIncrementMode
{
    /// <summary>
    /// Roll value is set and changed only by explicit operator action
    /// (matches Webscan TruCheck "Set New Operator/Roll" behaviour).
    /// The caller supplies the roll number; VTCCP never changes it automatically.
    /// </summary>
    Manual,

    /// <summary>
    /// Roll starts at <see cref="SessionState.RollStartValue"/> and increments by 1
    /// each time <c>SetNewOperatorAndRoll()</c> is called.
    /// Decimal counting, step size = 1.
    /// </summary>
    AutoIncrement,

    /// <summary>
    /// Roll identifier is a date/time stamp generated when the session opens or when
    /// <c>SetNewOperatorAndRoll()</c> is called.
    /// Format: yyyyMMddHHmmss  (14 digits, UTC-local, sortable).
    /// </summary>
    DateTimeStamp,
}
