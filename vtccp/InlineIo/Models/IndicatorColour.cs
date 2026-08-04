namespace InlineIo.Models;

/// <summary>
/// Indicator pole lamp colours supported by CP Inline.
/// BLUE is reserved (TBD Engineering); Off is used by <see cref="IndicatorPoleController.ClearAsync"/>.
/// </summary>
public enum IndicatorColour
{
    Off,
    Red,
    Amber,
    Green,
    Blue,   // Reserved — channel assignment TBD Engineering
}
