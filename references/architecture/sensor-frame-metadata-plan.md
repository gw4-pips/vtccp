# Sensor & Frame Metadata — Planning Document

**Status**: PLANNED — not yet implemented  
**Created**: 2026-05-21  
**Scope**: Additive fields to `VerificationRecord`, `DeviceInfo`, and the Excel schema.
Zero breaking changes to existing schema (all new columns appended to Universal block).

---

## Motivation

Sensor resolution and transmitted image dimensions are deep-technical metadata that:
1. Cost nothing to capture once a device-model lookup table and a JPEG header parser exist
2. Are essential for correlating physical symbol size (X-dimension) with pixel measurements
3. Disambiguate DM390 (2MP) from DM395 (5MP) when `DEVICE.TYPE` returns ambiguous strings
4. Provide the full provenance chain needed for image-load re-verification (D4)

---

## Three distinct concepts — must not be conflated

| Concept | What it is | How determined |
|---|---|---|
| **Sensor full frame** | Native hardware resolution — always fixed per model | Static lookup by `DeviceModel` |
| **IMAGE.SIZE setting** | DMCC downscale factor for `IMAGE.SEND` output | `GET IMAGE.SIZE` at connect |
| **Transmitted verification frame** | Firmware ROI crop around the symbol in push XML JPEG | JPEG SOF0 header parse per scan |

These are three separate things. The push XML JPEG is **not** the full sensor frame and
**not** controlled by `IMAGE.SIZE` — it is a firmware-processed crop whose dimensions vary
with symbol position and size within the field of view.

---

## Sensor resolution — known values (from reference manuals)

| DeviceModel string | Resolution | Pixel pitch | Sensor size | Source |
|---|---|---|---|---|
| `DM475V` (Verifier) | **2448 × 2048** | 3.45 µm | 8.8 × 6.6 mm, 2/3" CMOS | DM475V ref manual |
| `DM475` (DPM/HD) | **2448 × 2048** | 3.45 µm | 8.5 × 7.1 mm, 2/3" CMOS | DM475V ref manual |
| `DM395` / `DM395V` | **2448 × 2048** | 3.45 µm | TBD — DM390 manual confirms 5MP | DM390 ref manual |
| `DM394` | **2048 × 1536** | 3.45 µm | ~8.99 mm diagonal | DM390 ref manual |
| `DM390` | **2048 × 1536** | 3.45 µm | ~8.99 mm diagonal | DM390 ref manual |

Note: The DM390 manual explicitly lists:
- DM394: Image Resolution 2048 × 1536
- DM395: Image Resolution 2448 × 2048

This means DM395V (next-gen verifier successor to DM475V) has the same sensor as the
DM475V — same pixel pitch, same frame size. The lookup table entry is ready now.

---

## DMCC: IMAGE.SIZE command

**Command**: `IMAGE.SIZE`  
**Actions**: SET/GET  
**Range**: enum 0–3 (0 = Full, 1 = 1/4, 2 = 1/16, 3 = 1/64)  
**Platforms**: ALL  
**Version**: 4.4.0+

`IMAGE.SIZE` controls the output resolution of `IMAGE.SEND` (the DMCC raw-image
retrieval command). It does **not** affect the push XML JPEG crop.  
Querying it at session start tells us the device's current output downscale preference,
which is useful for D4 image archival planning.

**Companion command**: `IMAGE.SEND` — retrieves the last acquired image at the
configured IMAGE.SIZE / IMAGE.FORMAT / IMAGE.QUALITY. Platforms: ALL. Version: 5.5.0+.

---

## Planned fields

### Group 1 — Per-session (captured at ConnectAsync, static for session lifetime)

| VerificationRecord field | DeviceInfo field | Schema column | Header | Source |
|---|---|---|---|---|
| `SensorWidthPx` | `SensorWidthPx` | "Sensor W (px)" | 8 | Static lookup by DeviceModel |
| `SensorHeightPx` | `SensorHeightPx` | "Sensor H (px)" | 8 | Static lookup by DeviceModel |
| `SensorPixelPitchUm` | `SensorPixelPitchUm` | "Pixel (µm)" | 7 | Static lookup by DeviceModel |
| `ImageSizeSetting` | `ImageSizeSetting` | "Image Size" | 7 | `GET IMAGE.SIZE` → "Full" / "1/4" / "1/16" / "1/64" |

### Group 2 — Per-scan (dynamic, D4 scope)

| VerificationRecord field | Schema column | Header | Source |
|---|---|---|---|
| `VerifFrameWidthPx` | "Frame W (px)" | 8 | JPEG SOF0 header from push XML base64 |
| `VerifFrameHeightPx` | "Frame H (px)" | 8 | JPEG SOF0 header from push XML base64 |
| `VerifImageBytes` | "Image (bytes)" | 9 | base64-decoded byte length of push JPEG |

---

## Implementation plan

### Step 1 — `DeviceSensorSpec.cs` (new file, `DeviceInterface` project)

```csharp
public sealed record SensorSpec(
    int    WidthPx,
    int    HeightPx,
    double PixelPitchUm,
    string SensorSize);

public static class DeviceSensorSpecs
{
    private static readonly Dictionary<string, SensorSpec> Table =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["DM475V"] = new(2448, 2048, 3.45, "2/3\""),
            ["DM475"]  = new(2448, 2048, 3.45, "2/3\""),
            ["DM395V"] = new(2448, 2048, 3.45, "2/3\""),   // DM390 manual confirmed
            ["DM395"]  = new(2448, 2048, 3.45, "2/3\""),
            ["DM394"]  = new(2048, 1536, 3.45, "TBD"),
            ["DM390"]  = new(2048, 1536, 3.45, "TBD"),
        };

    public static SensorSpec? TryGet(string? model) =>
        model is not null && Table.TryGetValue(model, out var s) ? s : null;
}
```

### Step 2 — `DmccCommand.cs` — add `GetImageSize`

```csharp
public const string GetImageSize = "GET IMAGE.SIZE";
```

### Step 3 — `DeviceInfo` — add 4 new properties

```csharp
public int?    SensorWidthPx      { get; init; }
public int?    SensorHeightPx     { get; init; }
public double? SensorPixelPitchUm { get; init; }
public string? ImageSizeSetting   { get; init; }  // "Full", "1/4", "1/16", "1/64"
```

### Step 4 — `DeviceSession.ConnectAsync` — populate from lookup + DMCC

```csharp
var spec = DeviceSensorSpecs.TryGet(devType);
var imageSizeRaw = (await _client.SendAsync(DmccCommand.GetImageSize, ct)).Body;
DeviceInfo = new DeviceInfo
{
    // existing fields ...
    SensorWidthPx      = spec?.WidthPx,
    SensorHeightPx     = spec?.HeightPx,
    SensorPixelPitchUm = spec?.PixelPitchUm,
    ImageSizeSetting   = imageSizeRaw switch {
        "0" => "Full", "1" => "1/4", "2" => "1/16", "3" => "1/64",
        _   => imageSizeRaw,
    },
};
```

### Step 5 — `VerificationRecord` — add 7 new fields (Groups 1 + 2)

Group 1 fields added alongside existing device fields; Group 2 alongside future D4 image fields.

### Step 6 — `ContextFromDeviceInfo()` — wire Group 1 fields

### Step 7 — Schema + mappers — 7 new Universal columns

Insert Group 1 columns after `ConnectionMedium` (before `CalibrationDate`).
Insert Group 2 columns alongside D4 image columns (TBD position).

### Step 8 (D4 scope) — JPEG SOF0 header parser

```csharp
public static (int Width, int Height) ReadJpegDimensions(ReadOnlySpan<byte> jpeg)
{
    // Scan for SOF0 (0xFFC0) or SOF2 (0xFFC2) marker
    for (int i = 0; i < jpeg.Length - 8; i++)
    {
        if (jpeg[i] == 0xFF && (jpeg[i+1] == 0xC0 || jpeg[i+1] == 0xC2))
        {
            int height = (jpeg[i+5] << 8) | jpeg[i+6];
            int width  = (jpeg[i+7] << 8) | jpeg[i+8];
            return (width, height);
        }
    }
    return (0, 0);
}
```

---

## Dependencies

| Step | Blocked by |
|---|---|
| Steps 1–7 | Nothing — can be done now |
| Step 8 (JPEG parse) | D4 implementation (base64 decode of push XML JPEG payload) |
| Group 2 schema columns | Step 8 |

Steps 1–7 (session-level sensor metadata) are fully unblocked and can be
implemented in the same session as the next batch of parser work.

---

## Open questions

1. Does `GET IMAGE.SIZE` return the device's currently stored setting, or the
   setting last applied to an `IMAGE.SEND` call? (Assumption: stored device setting.)
   Probe: include in v1.31 diag output if needed.

2. DM475V: what does `GET IMAGE.SIZE` return on a factory-fresh unit? Expected: `0` (Full).
   If it returns `0` the column will read "Full" on every DM475V row — informative.

3. DM395V SensorSize string (8.99 mm diagonal vs the DM475V 8.8×6.6 mm / 8.5×7.1 mm).
   DM390 manual gives "8.99 mm diagonal" for the DM390-series sensor. DM395V may differ.
   Fill in when DM395V hardware manual is available.

4. Push XML: does `<JpegImageBase64>` always carry a JPEG, or can it be absent on no-read?
   Answer expected from first D4 raw-push capture.
