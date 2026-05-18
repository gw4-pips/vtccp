# Cognex DataMan Communications and Programming Guide — Digest

**Source**: `comms-and-programming-guide-25.4.1.1.pdf` (2.8 MB)
**Version**: 2025.4.1.1
**Digested by**: explorer subagent, 2026-05-18

**Provenance note**: This digest was produced by a subagent reading the source
PDF. It is comprehensive at the top level but **not page-exhaustive** —
section 2 below shows the documented commands grouped by family but is not
the complete flat catalog. For any specific command syntax used in
production code, verify the exact wording against the source PDF.

---

## TOP-LINE FINDINGS (read this section first)

1. **DMCC framing is `||><command><args>\r\n`** with optional checksum + command-id header. **This is different from what's documented in our codebase right now** (`vtccp/DeviceInterface/Dmcc/DmccResponse.cs` parses `\r\n<status>\r\n[\r\n<body>\r\n]`). Both forms exist depending on response mode:
   - **Silent mode** (`COM.DMCC-RESPONSE 0`, default): raw ASCII / empty — what our DmccClient handles today.
   - **Extended mode** (`COM.DMCC-RESPONSE 1`): framed `||checksum:command-id[status]<payload>\r\n` — what the guide shows.
   - **Implication**: our DmccClient is correct for default device config; if anyone toggles to extended mode, parser breaks.

2. **DMCC status codes (canonical)**: 0=Success, 100=Unidentified error, 101=Command invalid, 102=Parameter invalid, 103=Checksum incorrect, 104=Parameter rejected (state), 105=Reader unavailable.
   - **The codes 6/8/-1/-2/-3 documented in `vtccp/README.md` are TRIGGER-specific or VTCCP-internal** (-1/-2/-3 = parse fail / timeout / no bytes; 6 = no-read; 8 = busy). They live in a separate namespace from the canonical 100-series. Worth a README clarification.

3. **`r.image` is metadata only.** Properties: `index`, `FoV`, `RoI`, `exposureTime`, `gain`, `illEnabled`, `id`. **The actual JPEG bytes live on `r.trucheck.jpegImage` as a Base64-encoded string.** This is the answer to the long-standing "how do we get at the image?" probe question. **Huge for v1.24** — we can emit `<JpegImageBase64>...</JpegImageBase64>` directly in the push payload.

4. **DMCC `GET IMAGE`** returns raw binary of the last captured image. Combined with **TCP-IMAGE push protocol** `[4 bytes Size][4 bytes Type][128 bytes Filename][Image Data]` (Type: 0=BMP, 1=PNG, 2=JPG, 9=SVG). This is the pull path for images.

5. **Image-load (the loaded-image flow we've been planning) appears to use DMCC `WBU` (Write Buffer)** to ingest external images for processing. Confirmed in the DMST digest (sibling file); referenced here as the DMCC mechanism. Not deeply documented in the Comms guide — needs additional source mining.

6. **Naming-case mismatch in our empirical probe**: the guide spells one metric `axialNonUniformity` (capital U); our v1.23 probe enumerated `axialNonuniformity` (lowercase u). At least one of these is wrong; v1.24 must probe both spellings to determine which actually returns data.

7. **Trigger-mode codes**: 0=Manual/Software, 2=Continuous, 3=External, 4=Presentation, 5=Burst. Software `TRIGGER ON` is **ignored in Continuous or External modes**.

8. **Output channel targeting in scripts** — instead of `output.content = ...`, can use `output.NetworkClient = "<xml>...</xml>"` to target specific clients. Our v1.23 sets `output.content`; this works but isn't channel-specific.

---

## 1. DMCC Protocol Overview

ASCII-based, runs over TCP (port 23 default), RS232, or USB.

### Framing
```
Header:  ||checksum:command-id>
Command: ASCII string + space-separated args
Footer:  \r\n
```

- `checksum`: 0 (none, default) or 1 (XOR checksum before footer)
- `command-id`: integer returned in response, for tracking
- Commands are case-insensitive
- String arguments must escape: `\"`, `\\`, `\|`, `\t`, `\r`, `\n`

### Response Modes — `SET COM.DMCC-RESPONSE <0|1>`

| Mode | Value | Format |
|---|---|---|
| Silent | 0 (default) | Raw ASCII or empty |
| Extended | 1 | `||checksum:command-id[status]<payload>\r\n` |

### Status Codes (canonical DMCC)

| Code | Meaning |
|---|---|
| 0 | Success |
| 100 | Unidentified error |
| 101 | Command invalid |
| 102 | Parameter invalid |
| 103 | Checksum incorrect |
| 104 | Parameter rejected (reader state) |
| 105 | Reader unavailable (offline) |

**See top-line finding #2** for separation from TRIGGER-specific codes
and VTCCP-internal codes.

---

## 2. DMCC Command Namespace (by family)

Subagent did not provide an exhaustive flat catalog — these are the
families and notable commands surfaced from the guide. **For complete
syntax of any specific command, verify against source PDF.**

| Family | Examples |
|---|---|
| **TRIGGER** | `TRIGGER ON`, `TRIGGER OFF`, `SET TRIGGER.TYPE <0..5>` |
| **DEVICE** | `GET DEVICE.TYPE`, `GET DEVICE.NAME`, `GET DEVICE.MAC-ADDRESS`, `GET DEVICE.ID` |
| **SYMBOL** | `GET SYMBOL.DATAMATRIX <ON\|OFF>`, `SYMBOL.QR`, `SYMBOL.CODE128`, also per-symbology toggles |
| **FORMAT** | `SET FORMAT.MODE <0\|1>` (0=Basic, 1=Script), `SCRIPT.LOAD`, `SCRIPT.SEND` |
| **COM** | `COM.DMCC-RESPONSE`, `COM.DMCC-TARGET` (base-station routing) |
| **FTP / TCP** | `FTP-IMAGE.IP-ADDRESS`, `SET TCP-IMAGE.ENABLE ON`, push-target configs |
| **IMAGE** | `GET IMAGE`, `LIVEIMG.MODE` |
| **SYSTEM** | `REBOOT`, `CONFIG.SAVE`, `STATISTICS.RESET`, `SET TIME.NOW` |
| **CALIBRATION** | (TruCheck cal commands — see §6) |

---

## 3. Image-Related Commands (CRITICAL for VTCCP)

### Pull (get image from device)

- **`GET IMAGE`** — returns raw binary of last captured image
- **SDK equivalent**: `DataManSystem.GetLastReadImage()`

### Live image polling

- **`SET LIVEIMG.MODE 2`** — enables live mode
- **SDK helper**: `GetLiveImage(format, size, quality)`

### Push (device sends image to host)

- **`SET TCP-IMAGE.ENABLE ON`**
- Reader can act as TCP server or client
- **TCP framing**: `[4 bytes Size][4 bytes Type][128 bytes Filename][Variable Image Data]`

### Image type codes (for TCP push)

| Type | Format |
|---|---|
| 0 | BMP |
| 1 | PNG |
| 2 | JPG |
| 9 | SVG |

### Load image into device (for VTCCP loaded-image flow)

- **DMCC `WBU` (Write Buffer)** — ingest external image for processing
  pipeline. Documented more thoroughly in the DMST digest; the Comms
  guide references the mechanism but doesn't fully spec it.
- **Implication**: VTCCP's loaded-image flow path is `WBU` → trigger →
  push payload arrives via existing Network Client path. **Reuses
  everything we already have.**

---

## 4. Push-Script JavaScript API — `r.*` formal reference

The script entrypoint is `onResult(decodeResults, readerProperties, output)`
where `decodeResults` is an array of `DecodeResult` objects (we call this
`r`).

### `r.*` siblings (DecodeResult)

| Property | Type | Description |
|---|---|---|
| `r.decoded` | bool | True if decode succeeded |
| `r.content` | string | Raw decoded data |
| `r.decodeTime` | int | ms |
| `r.triggerTime` | int | ms |
| `r.symbology` | Object | See below |
| `r.image` | Object | **Capture metadata only** — not the bytes |
| `r.metrics` | Object | Standard ISO/AIM quality metrics |
| `r.trucheck` | Object | Verifier-specific extended metrics + jpegImage |
| `r.validation` | Object | GS1 / DoD UID validation |
| `r.source` | string | Device name |

### `r.symbology` properties

| Property | Type | Description |
|---|---|---|
| `name` | string | e.g. "Data Matrix" |
| `id` | string | ISO 15424 symbology identifier |
| `quality` | int | 0-100 |
| `moduleSize` | float | pixels per module |
| `corners` | Point[] | 4 points {x,y}, order UL/UR/LR/LL |
| `center` | Point | {x,y} |
| `angle` | int | degrees |
| `size` | Point | columns × rows |
| `failureCode` | (per probe) |

### `r.image` properties (METADATA — see top-line #3)

| Property | Description |
|---|---|
| `index` | Image index |
| `FoV` | Field of view rectangle {x, y, w, h} |
| `RoI` | Region of interest rectangle {x, y, w, h} |
| `exposureTime` | Capture exposure |
| `gain` | Capture gain |
| `illEnabled` | Illumination state |
| `id` | Unique image ID |

### `r.trucheck` properties (verifier-specific)

| Property | Description |
|---|---|
| `cellContrast`, `cellModulation`, `gridNonUniformity`, `unusedErrorCorrection` | Verification metrics |
| `overall` | Comprehensive formal grade |
| **`jpegImage`** | **Base64-encoded JPEG of the verified symbol** |
| `calibrationDate` | Calibration timestamp (string) |
| `axialNonUniformity` | (guide spelling — note case vs our probe) |

### `r.metrics` (QualityMetrics)

Each property is `{raw: number, grade: string}`. Common keys (intersection
with our v1.23 enumeration):

`symbolContrast`, `printGrowth`, `axialNonUniformity` (guide case),
`unusedErrorCorrection`, `modulation`, `gridNonUniformity`,
`fixedPatternDamage`.

### Gaps vs our v1.23 empirical probe

**Documented but not seen in probe:**
- `r.image.FoV`, `r.image.RoI` (we know `r.image` is a sibling but didn't probe properties)
- `r.metrics.singleScanInt`

**In our probe but not in guide:**
- `r.ledIntensity` — likely firmware extension
- `r.metrics.contrastUniformity` — guide maps this to `TruCheckMetricGeneral.contrastUniformity`

**Naming-case mismatch** (top-line #6):
- Guide: `axialNonUniformity` (capital U)
- Our v1.23: `axialNonuniformity` (lowercase u)
- v1.24 should probe both to determine which actually returns data.

---

## 5. File Transfer & Network Push

### FTP push
- Triggered automatically on result if `FTP-IMAGE.ENABLE` is ON.

### TCP push (alternative to FTP)
- Reader can be **Local TCP Server** (host connects to reader) or **TCP Client** (reader pushes to host).
- Local TCP Server: **max 1 connection, no buffering** — important constraint.

### Scripting upload
- `dmccCommand("SCRIPT.LOAD", length)` allows pushing JavaScript logic to the device via TCP — alternative to DMST UI for script deployment.

---

## 6. Calibration & Verification (TruCheck)

### Properties accessible from script
- `r.trucheck.calibrationDate` (string)
- `readerProperties.status3D.fieldCalibrated` (bool)
- `readerProperties.status3D.factoryCalibrated` (bool)

### Verifier metrics on `r.trucheck`
- `cellContrast`, `cellModulation`, `gridNonUniformity`, `unusedErrorCorrection`
- `overall` — comprehensive formal grade
- `jpegImage` — base64 JPEG

**VTCCP implication**: We currently put calibration date in `<CustomNote>` as
a manual workaround. We can promote to first-class `<CalibrationDate>` field
in v1.24, sourced directly from `r.trucheck.calibrationDate`. Likewise
`<Calibrated>`/`<FactoryCalibrated>` booleans for the OpticsCompliant
computation.

---

## 7. Trigger Modes

| Code | Mode | Software TRIGGER ON honored? |
|---|---|---|
| 0 | Manual / Software | Yes |
| 2 | Continuous | No (ignored) |
| 3 | External | No (ignored) |
| 4 | Presentation | Maybe (firmware-specific) |
| 5 | Burst | Maybe (firmware-specific) |

**Aligns with our field-tested README knowledge** that Single mode is
required for the ⚡ Trigger Scan button. Other modes "ignore DMCC TRIGGER
commands entirely" per README.

---

## 8. SDK vs Raw DMCC

| Concern | Raw DMCC | DataMan SDK |
|---|---|---|
| Platform | Cross-platform (any TCP) | Windows .NET (DLLs) |
| Overhead | Low | Higher (event wrappers, ResultCollector, ImageArrived) |
| Configuration | Manual `SET`/`GET` | High-level wrappers |
| VTCCP use | `DmccClient.cs` (control + trigger) | `DataManSdkClient.cs` (primary push session) |

VTCCP uses both — SDK for the long-lived push session, raw DMCC for
config tweaks and triggers from the WPF UI.

---

## 9. Quirks & Gotchas

- **Case sensitivity**: Command names + public parameter names are
  case-insensitive. String arguments require escaping.
- **Telnet port 23 sharing**: Allowed but high-frequency polling from
  DMST can collide with custom clients (matches our README observation).
- **DTR requirement**: For serial/USB-COM, `DtrEnable=true` mandatory.
- **Local TCP Server limit**: 1 concurrent connection, no buffering —
  if VTCCP disconnects, results in flight are lost.

---

## 10. What this changes for VTCCP

### Immediate v1.24 wire-ups unlocked

1. **`r.trucheck.jpegImage`** → new `<JpegImageBase64>` push field.
   Enables: report image inclusion, audit trail, reverse-report
   reproduction with image. **Big.**
2. **`r.trucheck.calibrationDate`** → promote from CustomNote to
   first-class `<CalibrationDate>` column.
3. **`readerProperties.status3D.fieldCalibrated`** → first-class
   `<FieldCalibrated>` boolean. Becomes input to OpticsCompliant.
4. **Probe both `axialNonUniformity` and `axialNonuniformity` spellings**
   — empirical resolution of the case-mismatch.
5. **`r.image.FoV` / `r.image.RoI` / `exposureTime` / `gain` / `id`** —
   probe these accessors; if populated, surface to schema.

### Loaded-image flow path (D4 in session plan)

`WBU` (Write Buffer) → push image bytes → trigger → existing Network
Client push path returns results. **Reuses 100% of existing infrastructure
downstream of the device.** Architecture is simpler than we'd assumed.

### Documentation fixes

- README §"DMCC TRIGGER response codes": clarify that codes 6/8 are
  TRIGGER-specific (no-read / busy), -1/-2/-3 are VTCCP-internal, and
  the canonical 100-series DMCC errors live in a different namespace.
- README should also note that the `DmccResponse.cs` parser assumes
  silent mode and would need extending if extended mode is ever enabled.
