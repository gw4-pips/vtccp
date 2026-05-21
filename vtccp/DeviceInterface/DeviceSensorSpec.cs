namespace DeviceInterface;

/// <summary>
/// Physical sensor specification for a known Cognex DataMan reader model.
/// All values are manufacturing constants — they do not change for a given model
/// and are not queryable via DMCC.  Resolution data confirmed from official
/// Cognex reference manuals (DM475V ref 25.4.1.1; DM390 ref 25.4.1.2).
/// </summary>
public sealed record SensorSpec(
    /// <summary>Native sensor width in pixels, e.g. 2448.</summary>
    int    WidthPx,
    /// <summary>Native sensor height in pixels, e.g. 2048.</summary>
    int    HeightPx,
    /// <summary>Square pixel pitch in micrometres, e.g. 3.45.</summary>
    double PixelPitchUm,
    /// <summary>Physical sensor diagonal or H×V description, e.g. "2/3\"".</summary>
    string SensorSize);

/// <summary>
/// Per-model sensor lookup table.
/// Call <see cref="TryGet"/> with the string returned by DEVICE.TYPE at connect time.
///
/// Models confirmed from primary Cognex documentation:
///
///   DM475V / DM475   — 2448×2048, 3.45 µm, 2/3" CMOS, 8.8×6.6 mm (Verifier)
///                       8.5×7.1 mm (DPM/HD variant) — same pixel count
///   DM395V / DM395   — 2448×2048, 3.45 µm (DM390 reference manual: DM395 = 5 MP)
///                       DM395V is the next-gen successor to the DM475V series
///   DM394            — 2048×1536, 3.45 µm (DM390 reference manual: DM394 = 3 MP)
///   DM390            — 2048×1536, 3.45 µm (base DM390-series model)
///   DM380            — Coglink-capable tunnel reader; sensor spec TBD pending
///                       DM380 hardware reference manual acquisition
///
/// Add entries as new models are validated against primary documentation.
/// Unknown model strings return null; the calling code treats null as "not in table".
/// </summary>
public static class DeviceSensorSpecs
{
    private static readonly Dictionary<string, SensorSpec> Table =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // ── DM475V series (current production verifier) ────────────────────
            ["DM475V"] = new(2448, 2048, 3.45, "2/3\""),
            ["DM475"]  = new(2448, 2048, 3.45, "2/3\""),

            // ── DM395V series (next-gen successor to DM475V) ──────────────────
            // Source: DM390 reference manual 25.4.1.2 — DM395 = 2448×2048 (5 MP)
            // DM395V hardware reference manual not yet in library; entry ready now.
            ["DM395V"] = new(2448, 2048, 3.45, "2/3\""),
            ["DM395"]  = new(2448, 2048, 3.45, "2/3\""),

            // ── DM390 / DM394 series (3 MP) ───────────────────────────────────
            // Source: DM390 reference manual 25.4.1.2 — DM394 = 2048×1536 (3 MP)
            ["DM394"]  = new(2048, 1536, 3.45, "~8.99 mm diag"),
            ["DM390"]  = new(2048, 1536, 3.45, "~8.99 mm diag"),

            // ── DM380 (Coglink tunnel reader) ──────────────────────────────────
            // Sensor spec pending DM380 hardware reference manual.
            // Entry intentionally omitted until confirmed from primary source.
        };

    /// <summary>
    /// Returns the <see cref="SensorSpec"/> for <paramref name="model"/>,
    /// or <c>null</c> if the model is unknown or <paramref name="model"/> is null/empty.
    /// Match is case-insensitive.
    /// </summary>
    public static SensorSpec? TryGet(string? model) =>
        !string.IsNullOrWhiteSpace(model) && Table.TryGetValue(model, out var spec)
            ? spec
            : null;
}
