namespace InlineIo.Models;

/// <summary>
/// Maps logical CP Inline signals to physical relay-board channel numbers.
/// Channel indices are 1-based to match relay-board labelling.
/// Use -1 to mark a channel as "not wired" — the controllers skip it safely.
/// </summary>
/// <remarks>
/// ⚠ ENGINEERING TODO: confirm all channel assignments against the relay board
/// wiring diagram before first hardware test.  Default values are placeholders.
/// See also: ConveyorInterruptController — confirm whether ConveyorRestart is a
/// separate channel or implicit (de-energising ConveyorStop resumes the belt).
/// </remarks>
public record RelayChannelMap
{
    // ── Indicator pole ───────────────────────────────────────────────────────
    /// <summary>RED lamp relay channel.  Grade &lt; 1.8 (steady) or no-decode (flash).</summary>
    public int Red   { get; init; } = 1;

    /// <summary>AMBER lamp relay channel.  Grades 1.8–2.8.</summary>
    public int Amber { get; init; } = 2;

    /// <summary>GREEN lamp relay channel.  Grades 2.9–4.0.</summary>
    public int Green { get; init; } = 3;

    /// <summary>BLUE lamp relay channel.  Assignment TBD Engineering; -1 = not wired.</summary>
    public int Blue  { get; init; } = -1;

    // ── Conveyor interrupt ───────────────────────────────────────────────────
    /// <summary>
    /// Relay channel that opens/interrupts the conveyor drive circuit.
    /// Energise = stop line.
    /// </summary>
    public int ConveyorStop    { get; init; } = 4;

    /// <summary>
    /// Relay channel that sends an explicit restart pulse to the conveyor PLC.
    /// Set to -1 if de-energising ConveyorStop is sufficient to resume (confirm with engineering).
    /// </summary>
    public int ConveyorRestart { get; init; } = -1;

    // ── Auxiliary ────────────────────────────────────────────────────────────
    /// <summary>Audible buzzer relay channel.  -1 = not wired.</summary>
    public int Buzzer { get; init; } = -1;

    /// <summary>Default map — all channels are placeholder values.  Override before production wiring.</summary>
    public static RelayChannelMap Default => new();
}
