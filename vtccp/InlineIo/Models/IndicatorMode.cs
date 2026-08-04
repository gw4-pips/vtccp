namespace InlineIo.Models;

/// <summary>
/// Display mode for an active indicator colour.
/// Flash period is controlled by <see cref="IndicatorPoleController.FlashPeriodMs"/>.
/// </summary>
public enum IndicatorMode
{
    Steady,
    Flash,
}
