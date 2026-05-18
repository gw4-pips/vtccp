# OpticsSource Model

**Date**: 2026-05-18
**Status**: Confirmed — discriminant logic verified across two loaded-image captures and multiple live scans.

---

## 1. The OpticsSource enum

```
OpticsSource = "LiveScan"    // symbol scanned under device optics
OpticsSource = "LoadedImage" // digital image file pushed to device via DMCC
```

The `<OpticsSource>` element is computed **in the push script** (v1.25), not by the parser.
The parser reads it as a pre-computed string and stores it in `VerificationRecord.OpticsSource`.

---

## 2. Discriminant logic (confirmed v1.24, 2026-05-18)

### Reliable discriminant (BOTH conditions required simultaneously)

```
ContrastUniformity === -1  AND  MRD === -1
    → OpticsSource = "LoadedImage"
otherwise
    → OpticsSource = "LiveScan"
```

**Both fields must be −1 simultaneously.** Either alone is insufficient — a degraded live scan
could theoretically produce an extreme CU or MRD reading.

### Evidence

| Scan | ContrastUniformity | MRD | SymbolAngle | Verdict |
|---|---|---|---|---|
| Live DM (GS1 Format 06) | 72 at module(10,5) | 71% (77%−6%) | 1° | LiveScan ✓ |
| Loaded QR (URL image) | −1 | −1 | 360° | LoadedImage ✓ |
| Loaded QR (email image) | −1 | −1 | 0° | LoadedImage ✓ |

### SymbolAngle is NOT a reliable discriminant

The URL QR (loaded) had SymbolAngle=360° and the email QR (loaded) had SymbolAngle=0°.
Both are loaded images. SymbolAngle reports the angle the symbol appeared at in the file —
an upright image gives 0°, some images are stored rotated giving non-zero values. Do not use
SymbolAngle for OpticsSource detection.

### v1.25 script implementation

```javascript
var _cuNum  = parseFloat(gnProp("contrastUniformity"));
var _mrdNum = parseFloat(gnProp("MRD") || gnProp("mrd"));
o += elem("OpticsSource",
    (!isNaN(_cuNum) && _cuNum === -1 && !isNaN(_mrdNum) && _mrdNum === -1)
    ? "LoadedImage" : "LiveScan");
```

---

## 3. OpticsCompliant flag

A record is **OpticsCompliant** when the measurement was performed under controlled optics that
meet ISO/IEC 15426-2 §5 (illumination geometry, wavelength, aperture tolerances).

```
OpticsCompliant = (OpticsSource == "LiveScan")
               && (FieldCalibrated == true || FactoryCalibrated == true)
```

A loaded-image scan is **not** optics-compliant regardless of calibration state because:
- The device's illumination system is not involved.
- The image may have been captured at non-conforming wavelength, aperture, or geometry.
- ISO/IEC 15426-2 §8.3 defines primary reference test symbols as physical targets scanned
  under the reference illuminant, not digital files.

`OpticsCompliant` is a **derived property** — it is not emitted by the push script and is not
stored in `VerificationRecord`. It is computed at report-generation time:

```csharp
bool OpticsCompliant(VerificationRecord r) =>
    r.OpticsSource == "LiveScan"
    && (r.FieldCalibrated == true || r.FactoryCalibrated == true);
```

---

## 4. Per-field suppression rules under LoadedImage

When `OpticsSource == "LoadedImage"`, the following fields are **not applicable** (device emits
sentinel values and should be reported as N/A, not as measurements):

### Suppressed on LoadedImage (sentinel values, do not display as measurements)

| Field | Sentinel on LoadedImage | Reason |
|---|---|---|
| `ContrastUniformity` | −1 | Not computed — discriminant field |
| `MRD` | −1 | Not computed — discriminant field |
| `LLS_Grade`, `BLS_Grade` | "X" (N/A) | Finder-pattern border not assessed for loaded images |
| `LQZ_Grade`, `BQZ_Grade`, `TQZ_Grade`, `RQZ_Grade` | "X" (N/A) | Quiet zone assessment requires optics |
| `SC_Percent`, `SC_Grade` | May be NaN/"" | Symbol contrast under no-optics condition is undefined |

### Retained on LoadedImage (still meaningful)

| Field | Retained reason |
|---|---|
| `OverallGrade` | Device computes ISO 15415 grades from the digital image; grade reflects image quality, not print quality |
| `MatrixSize` | Structural property of the symbol |
| `DecodedData` | Content decode is independent of optics |
| `EncodedCharacters`, `ErrorsCorrected` | Structural / ECC properties |
| `ANU_Percent`, `GNU_Percent` | Computed from the image data |
| `ModuleSizePx` | Pixels per module in the loaded image |
| `QR_*` params | Structure-based grades are computable from digital image |
| `JpegImageBase64` | The source image itself |

### OpticsCompliant report banner

All reports must display a visible banner when `OpticsCompliant == false`:

> **⚠ Loaded Image — Not Optics Compliant**
> This record was produced from a digital image file, not from a live scan under device optics.
> Grade values reflect the quality of the digital image, not the physical print under ISO/IEC 15426-2
> reference conditions. Not valid for compliance to ISO/IEC 15415.

When `OpticsSource == "LiveScan"` but calibration is unknown (`FieldCalibrated == null && FactoryCalibrated == null`):

> **⚠ Calibration status unknown**
> Field and factory calibration status were not reported. Compliance cannot be confirmed.

---

## 5. Future variant: SBG-non-V (not yet encountered)

The Cognex DataMan product line includes non-verifier readers that can produce trucheck-like output
in some configurations. If a future push result has:

- `ContrastUniformity` not equal to −1 (real value present)
- BUT `FieldCalibrated == false` AND `FactoryCalibrated == false`
- AND firmware variant string indicates a non-V model

This would be an `OpticsSource = "LiveScan"` but `OpticsCompliant = false` because the device is
uncalibrated. The current discriminant logic handles this correctly — `OpticsSource` would still
be `"LiveScan"` (optics were used), but `OpticsCompliant` would be `false` (not calibrated).

No firmware variant string parsing is implemented yet. This note is reserved for when non-V device
support is added.

---

## 6. D4 implementation requirements

When `D4` (WPF image-load full implementation) is built:

1. WPF file-open dialog must filter to JPEG/JPG only.
   Device confirmed: JPEG accepted, PRN rejected.
2. If user selects PNG, BMP, TIFF, or other non-JPEG: VTCCP converts to JPEG before DMCC push.
3. After push, parser will see `OpticsSource = "LoadedImage"` in the response.
4. Record stored with `OpticsSource = "LoadedImage"`.
5. Report generation suppresses the fields in §4 above and shows the banner.
