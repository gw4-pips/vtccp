# DMCC Command Reference Digest — fw 6.1.16_sr4
> Source: DataMan DMCC Reference, Revision 26.1.0.27 (2026-04-21), 794 HTML files
> Extracted 2026-05-20. All commands confirmed against the authoritative reference
> for firmware 6.1.16_sr4 running on DM475V @ 10.10.10.7.

---

## Critical corrections vs. prior VTCCP code

| Old (wrong) | Correct | Impact |
|---|---|---|
| `GET FIRMWARE.VER` | `GET DEVICE.FIRMWARE-VER` | `DmccCommand.cs` — fixed |
| `GET DEVICE.ID` | `GET DEVICE.SERIAL-NUMBER` | `DmccCommand.cs` + `DeviceSession.cs` — fixed |
| `GET CODE.UPCEAN-SUPPLEMENT-DIGIT` | `GET UPC-EAN.SUPPLEMENT` | `DmccCommand.cs` — fixed |
| UPC supplement values 0–5 | 0–4 only (no value 5 on DM475V) | `SessionViewModel.cs` — values revised |
| `CALIBRATION.DATE` | Not in reference — verify live | Unknown; may be firmware-specific |

---

## System / Device Identity

| Command | R/W | Platforms | Description |
|---|---|---|---|
| `GET DEVICE.TYPE` | R | ALL | Model string, e.g. `DM475V` |
| `GET DEVICE.FIRMWARE-VER` | R | ALL | Firmware version, e.g. `6.1.16.0015` |
| `GET DEVICE.NAME` | R/W | ALL | User-configurable reader label |
| `GET DEVICE.SERIAL-NUMBER` | R | ALL | Serial number |
| `GET DEVICE.DESCRIPTION` | R | ALL | Device description |
| `GET DEVICE.FEATURE-KEYS` | R | ALL | Comma-delimited installed feature keys (e.g. OCR key `DM-KEY-OCR`) |
| `GET DATA.RESULT-TYPE` | R/W | ALL | Bit field: 0=None, 1=Result, 512=MS test response. Default: 513 |
| `DMCC.SAVE` | — | ALL | Persists connection-layer settings (`COM.DMCC-RESPONSE`, `DATA.RESULT-TYPE`, etc.) |
| `DMCC.RESET` | — | ALL | Resets connection-layer settings to defaults |

**Note**: `DEVICE.FEATURE-KEYS` is the correct way to detect whether OCR (`DM-KEY-OCR`) is
installed. Returns empty string if no feature keys are active.

---

## UPC/EAN Supplemental

| Command | R/W | Platforms | Values |
|---|---|---|---|
| `GET/SET UPC-EAN.SUPPLEMENT` | R/W | ALL | 0=Ignore, 1=Required, 2=Required 2-digit, 3=Required 5-digit, 4=Not Required |
| `GET/SET UPC-EAN.EAN8SUPL` | R/W | ALL | ON/OFF — allow EAN-8 supplementals |

**`UPC-EAN.SUPPLEMENT` value semantics** (confirmed, v3.0.0):
- `0` Ignore — supplemental digits not decoded at all
- `1` Required — any supplemental length required
- `2` Required 2-digit — exactly 2-digit add-on required
- `3` Required 5-digit — exactly 5-digit add-on required
- `4` Not Required — supplemental decoded if present, not required for pass

---

## TruCheck Commands — DM475V + DM8072V
> All commands: Version 6.1.10 unless noted. Platform: DM475V, DM8072V.

### Application Standard & Grading

| Command | R/W | Values / Range |
|---|---|---|
| `GET/SET TRUCHECK.APPLICATION-STANDARD` | R/W | 0=GS1, 1=HIBCC, 2=UDI (HIBCC+GS1), 3=UID, 4=Auto, 5=Custom, 6=Cryptocode |
| `GET/SET TRUCHECK.GRADING-STANDARD` | R/W | 0=ISO/IEC 15415/6, 1=ISO/IEC 29158:2020 |
| `GET/SET TRUCHECK.APPLICATION-GS1-TABLE` | R/W | [0–11]; 0=Auto, 1=Table 1, 2=Table 2, … |
| `GET/SET TRUCHECK.APPLICATION-CUSTOM-DATA-PARSING-STANDARD` | R/W | 0=None, 1=GS1, 2=HIBCC, 3=UID |
| `GET/SET TRUCHECK.APPLICATION-CUSTOM-PASS-GRADE` | R/W | [0–40] — grade threshold without decimal (e.g. 15 = 1.5) |
| `GET/SET TRUCHECK.APPLICATION-CUSTOM-MINIMUM-X-DIM` | R/W | [1–1000] thousandths of an inch |
| `GET/SET TRUCHECK.APPLICATION-CUSTOM-MAXIMUM-X-DIM` | R/W | [1–1000] thousandths of an inch |

### Aperture

| Command | R/W | Values / Range |
|---|---|---|
| `GET/SET TRUCHECK.APERTURE` | R/W | 0=User Set, 1=Auto 80%/50%, 2=Auto aperture |
| `GET/SET TRUCHECK.APERTURE-SIZE` | R/W | [1–300] ten-thousandths of an inch (only when Aperture=User Set) |

### Report Header Fields — SessionManager Mapping

| DMCC Command | R/W | SessionManager field |
|---|---|---|
| `GET/SET TRUCHECK.OPERATOR-NAME` | R/W | `OperatorName` |
| `GET/SET TRUCHECK.COMPANY-NAME` | R/W | `CompanyName` |
| `GET/SET TRUCHECK.CUSTOM-NOTE` | R/W | `CustomNote` |
| `GET/SET TRUCHECK.BATCH-NUMBER` | R/W | `BatchNumber` (when AUTO-BATCH=OFF) |
| `GET/SET TRUCHECK.AUTO-BATCH` | R/W | ON/OFF — firmware auto-increments batch number in headers |

These four SET commands should be issued at session start to sync device report headers with
the VTCCP session state. They persist to flash — survives power cycle.

### Calibration

| Command | Type | Arguments |
|---|---|---|
| `TRUCHECK.CALIBRATE-ON RMax RMin` | Action | RMax [0.1–100.0], RMin [0.1–100.0]. Uses conformance standard test card. |
| `TRUCHECK.CALIBRATE-CUSTOM-ON RMax RMin XDim` | Action | As above + XDimension [0.1–100.0]. Uses any symbol. |
| `TRUCHECK.CALIBRATE-OFF` | Action | None. Must be sent after reading calibration target. |

### Miscellaneous

| Command | R/W | Values |
|---|---|---|
| `GET/SET TRUCHECK.DOT-PEEN` | R/W | ON/OFF — stick algorithm for peened dots (ISO 29158) |
| `GET/SET TRUCHECK.METRIC-UNITS` | R/W | OFF=Standard (mils/inches), ON=Metric (µm/mm) |
| `TRUCHECK.REPORT-SECTION "section" ON\|OFF` | Action | Sections: `CODE_IMAGE`, `GENERAL-CHARACTERISTICS-TABLE`, `QUALITY-DETAIL-TABLE`, `MODULATION-TABLE`, `ENCODATION-DETAIL-TABLE`, `ASCII_TABLE`, `APPLICATION-DATA-TABLE`, `CODEWORD-TABLE` |

**`TRUCHECK.REPORT-SECTION`** platform note: DM280, DM370, DM470, DM475V, DM8072V. v5.7.10 SR1.

### NOT available on DM475V (newer platforms only)

| Command | Available on |
|---|---|
| `TRUCHECK.ISO15415-VERSION` (0=2011, 1=2024) | DM280/290/370/380/470 — NOT DM475V |
| `TRUCHECK.ISO15416-VERSION` (0=2016, 1=2025) | DM280/290/370/380/470 — NOT DM475V |
| `TRUCHECK.ISO29158-VERSION` (0=2011, 1=2020, 2=2025) | DM280/290/370/380/470 — NOT DM475V |
| `VERIFICATION.ENABLE` | DM8072V, DM370, DM390, DM470 — NOT DM475V |

---

## QR Code

| Command | R/W | Platforms | Values |
|---|---|---|---|
| `GET/SET QR.QUALITY-METRICS` | R/W | ALL | 0=None, 1=ISO/IEC 15415, 2=AIM-DPM / ISO/IEC TR 29158 |
| `GET/SET QR.LOW-RES-2D` | R/W | ALL | ON/OFF — read QR down to 1.6 PPM in untrained mode (impacts speed) |
| `GET/SET QR.LEARNING-POLARITY` | R/W | Fixed-mount | 0=Dark on Light, 1=Light on Dark, 2=Either |
| `GET/SET QR.LEARNING-GRID-SIZE` | R/W | Fixed-mount | [11–177] — grid size for trained QR |

---

## Camera

| Command | R/W | Platforms | Notes |
|---|---|---|---|
| `GET/SET CAMERA.EXPOSURE` | R/W | DM475V (via list) | Exposure time |
| `GET/SET CAMERA.EXPOSURE-US` | R/W | DM370/390/470/580/8700 | Exposure in µs |
| `GET/SET CAMERA.GAIN` | R/W | DM370/390/470 | Analog/digital gain |
| `GET/SET CAMERA.AUTO-REGULATION` | R/W | DM370/390/470 | Auto exposure regulation |
| `GET/SET CAMERA.HDR-MODE` | R/W | DM370/390/470 | High Dynamic Range mode |
| `IMAGE.SEND` | Action | ALL | Send current image buffer |

**IMAGE.LOAD / IMAGE.REPLAY**: Not present in this DMCC reference. These are SDK-level
operations (via `IDataManSystem.SendImage()` / result replay) rather than raw DMCC text commands.

---

## Communication / Connection Layer

| Command | R/W | Notes |
|---|---|---|
| `SET COM.DMCC-RESPONSE 1` | W | Enable extended responses (default: 0=silent) |
| `GET/SET DATA.RESULT-TYPE` | R/W | Bit field: 1=Result, 512=MS test; default 513 |
| `DMCC.SAVE` | Action | Persist COM.DMCC-RESPONSE + DATA.RESULT-TYPE etc. across sessions |
| `DMCC.RESET` | Action | Reset connection layer to defaults |
| `DMCC.OUTPUT ON\|OFF` | R/W | Suppress automatic content output on all connections (v7.2.0, DM8700 etc.) |

---

## Actions (device-level)

| Command | Platforms | Notes |
|---|---|---|
| `REBOOT` | ALL | Hard reboot — use with caution |
| `CONFIG.SAVE` | ALL | Save settings to flash |
| `CONFIG.RESTORE` | ALL | Restore last-saved settings |
| `CONFIG.DEFAULT` | ALL | Factory reset |
| `DEVICE.STATS` | ALL | Returns device statistics |
| `DEVICE.BACKUP` | ALL | Backup device configuration |
| `DEVICE.RESTORE` | ALL | Restore device configuration |
| `BACKUP.EXPORT` | ALL | Export configuration backup |

---

## OCR

OCR is a software key feature (`DM-KEY-OCR`) — not enabled by default.
Detectable via `GET DEVICE.FEATURE-KEYS`. Once installed:
- OCR operates in **single trigger mode only**
- Configured via DMST OCR setup (font upload, ROI, string training)
- No DMCC OCR read/result command found in this reference — result delivered via push XML

---

## VTCCP D1/D4 Implementation Notes

**Session start sequence** (add to `DeviceSession.ConnectAsync`):
```
GET DEVICE.FEATURE-KEYS          → detect OCR key, log to session sidecar
SET TRUCHECK.OPERATOR-NAME {n}   → sync operator to device
SET TRUCHECK.COMPANY-NAME {n}    → sync company to device
SET TRUCHECK.CUSTOM-NOTE {n}     → sync custom note to device
SET TRUCHECK.BATCH-NUMBER {n}    → sync batch (when AUTO-BATCH=OFF)
GET TRUCHECK.APPLICATION-STANDARD → log current app standard to sidecar
GET TRUCHECK.GRADING-STANDARD    → log current grading standard to sidecar
GET TRUCHECK.APERTURE            → log aperture mode to sidecar
GET TRUCHECK.APERTURE-SIZE       → log aperture size to sidecar
```

**UPC-EAN supplemental** (already implemented in SessionViewModel):
- Correct key is `UPC-EAN.SUPPLEMENT` — was `CODE.UPCEAN-SUPPLEMENT-DIGIT` (wrong, now fixed)
- Values 0–4 (not 0–5 as originally documented)
- `SessionViewModel` mode enum must be updated to match 0–4 range

**CalibrationWarning**:
- `FieldCalibrated=false` observed on all live scans → flag all archived records
- `TRUCHECK.CALIBRATE-ON` / `TRUCHECK.CALIBRATE-OFF` provide the DMCC-level calibration
  workflow for when the customer wants to calibrate from VTCCP rather than DMST
