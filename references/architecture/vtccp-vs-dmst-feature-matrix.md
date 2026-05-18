# VTCCP vs. DMST Feature Matrix
## `architecture/vtccp-vs-dmst-feature-matrix.md`

> **Status**: Current as of 2026-05-18  
> **Inputs**: DMV-8072V reference manual digest, DMCC comms+programming guide digest,
> DMST 2025 manual digest, cross-manual synthesis, reader/device metadata inventory,
> v1.26–v1.28 live scan catalogs.

---

## 1. Purpose

VCCS DMV TruCheck Command Pilot (VTCCP) is a logging and export utility that
wraps the Cognex DataMan DM475V barcode verifier.  The DataMan Setup Tool (DMST)
is Cognex's own GUI for the same device.  This matrix documents which DMST
features are reproduced in VTCCP, which are intentionally omitted, and which
are VCCS-distinctive extensions that DMST cannot produce.

---

## 2. Data collection layer

| Capability | DMST | VTCCP | Notes |
|---|---|---|---|
| Live optical scan via USB/Ethernet | ✓ DMCC TCP:23 | ✓ DMCC TCP:23 | VTCCP uses DataManSdkClient |
| Push-mode streaming (scan-triggered) | ✓ | ✓ Network Client TCP:9004 | VTCCP DmstListener |
| Loaded-image scan (IMAGE.LOAD) | ✓ DMST UI | Partial (D4) | VTCCP architecture defined; not yet wired end-to-end |
| DMCC command replay (IMAGE.REPLAY) | ✓ | Architecture only | Unblocked post-A1 |
| Batch mode (roll/job) | N/A | ✓ | VTCCP-distinctive; DMST has no job/roll concept |
| Operator ID capture | N/A | ✓ | VTCCP-distinctive |
| Push script grammar version | N/A | v1.28 (device-confirmed) | VCCS-authored; DMST has no equivalent |

---

## 3. Symbology support

| Symbology | DMST report | VTCCP parser | Notes |
|---|---|---|---|
| Data Matrix ECC200 (square, all sizes 10×10–144×144) | ✓ | ✓ | Primary target, all grades wired v1.15+ |
| Data Matrix ECC200 (rectangular: 8×18, 8×32, 12×26, 12×36, 16×36, 16×48) | ✓ | ✓ | 16×36 device-confirmed v1.28 |
| QR Code Model 2 | ✓ | ✓ 8 grade params | Live optical scan pending; loaded-image confirmed v1.26 |
| GS1 DataMatrix | ✓ | ✓ via DFC | DataFormatCheckResult GS1 rows |
| Code 128 / GS1-128 | ✓ | ✓ ISO 15416 | ISO15416Mapper + per-scan table |
| Code 39 | ✓ | ✓ | Ratio field included |
| UPC/EAN variants | ✓ | ✓ | Per-scan table |
| PDF417 | ✓ DMST display | Pending | Not yet scoped |
| Aztec / MaxiCode / other 2D | ✓ DMST display | Not scoped | Out of VCCS scope |
| DPM (Direct Part Mark) mode | ✓ different metrics | Parser-aware | CC/CMOD branch in DataMatrix2DMapper |

---

## 4. ISO 15415 parameter coverage (2D Data Matrix)

Abbreviations: **W** = wired in push script and parsed into VerificationRecord;
**E** = emitted empty (field exists, firmware does not expose via push channel);
**N/A** = not applicable to this firmware/symbol combination.

| Parameter | Push XML field | VTCCP | DMST report | Source path |
|---|---|---|---|---|
| Overall Grade (letter) | `OverallGrade` | W | ✓ | q.overall.gradeLetter |
| Overall Grade (numeric) | `OverallGradeNumeric` | W | ✓ | q.overall.gradeValue |
| Formal Grade string | `FormalGrade` | W | ✓ | Derived: numeric+"/"+letter |
| Grading Standard | `GradingStandard` | W | ✓ | q.overall.gradingStandard |
| Application Standard | `ApplicationStandard` | W | ✓ | q.overall.applicationStandardName |
| Application Pass | `ApplicationPass` | W | ✓ | q.overall.applicationStandardPass |
| UEC % | `UECPercent` | W | ✓ | q.unusedErrorCorrection.raw |
| UEC Grade | `UECGrade` | W | ✓ | q.unusedErrorCorrection.grade |
| Symbol Contrast % | `SCPercent` | W | ✓ | q.symbolContrast.raw |
| SC Rl/Rd | `SCRlRd` | W | ✓ | q.reflectanceLight + q.reflectanceDark |
| SC Grade | `SCGrade` | W | ✓ | q.symbolContrast.grade |
| Modulation Grade | `MODGrade` | W | ✓ | q.modulation.grade |
| Reflectance Margin Grade | `RMGrade` | W | ✓ | q.reflectanceMargin.grade |
| ANU % | `ANUPercent` | W | ✓ | q.axialNonuniformity.raw |
| ANU Grade | `ANUGrade` | W | ✓ | q.axialNonuniformity.grade |
| GNU % | `GNUPercent` | W | ✓ | q.gridNonuniformity.raw |
| GNU Grade | `GNUGrade` | W | ✓ | q.gridNonuniformity.grade |
| FPD Value | `FPDValue` | W | ✓ | q.fixedPatternDamage.raw |
| FPD Grade | `FPDGrade` | W | ✓ | q.fixedPatternDamage.grade |
| Decode Grade | `DecodeGrade` | W | ✓ | q.decode.grade |
| Print Growth (AG) Value | `AGValue` | W | ✓ | q.printGrowth.raw |
| Print Growth Grade | `AGGrade` | W | ✓ | q.printGrowth.grade |
| Distributed Damage Grade | `DDGrade` | W | ✓ | q.distributedDamageGrade |
| Average Grade | `AverageGrade` | W | ✓ | q.averageGrade |
| Average Grade (numeric) | `AverageGradeNumeric` | W | ✓ | q.averageGrade.raw |
| Left L-Side Grade | `LLSGrade` | W | ✓ | q.leftLSide.grade |
| Bottom L-Side Grade | `BLSGrade` | W | ✓ | q.bottomLSide.grade |
| Left Quiet Zone Grade | `LQZGrade` | W | ✓ | q.leftQuietZone.grade |
| Bottom QZ Grade | `BQZGrade` | W | ✓ | q.bottomQuietZone.grade |
| Top QZ Grade | `TQZGrade` | W | ✓ | q.topQuietZone.grade |
| Right QZ Grade | `RQZGrade` | W | ✓ | q.rightQuietZone.grade |
| TTR % | `TTRPercent` | W | ✓ | q.topTransitionRatio.raw |
| TTR Grade | `TTRGrade` | W | ✓ | q.topTransitionRatio.grade |
| RTR % | `RTRPercent` | W | ✓ | q.rightTransitionRatio.raw |
| RTR Grade | `RTRGrade` | W | ✓ | q.rightTransitionRatio.grade |
| Top Clock Track Grade | `TCTGrade` | W | ✓ | q.topClockTrack.grade |
| Right Clock Track Grade | `RCTGrade` | W | ✓ | q.rightClockTrack.grade |
| Min Reflectance | `MinReflectance` | W | ✓ | q.reflectanceDark (primitive number) |
| H/V BWG % | `HorizontalBWG`, `VerticalBWG` | W | ✓ | q.general.horizontalBWG / verticalBWG |
| Nominal X-Dim | `NominalXDim` | W | ✓ | q.general.xDimension |
| Contrast Uniformity | `ContrastUniformity` | W | ✓ | q.general.contrastUniformity |
| MRD | `MRD` | W | ✓ | q.general.MRD |
| Image Polarity | `ImagePolarity` | W | ✓ | q.general.polarity |
| Matrix Size | `MatrixSize` | W | ✓ | Derived from modulationArray.length |
| Encoded Characters | `EncodedCharacters` | W | ✓ | encodationAnalysisArray.length |
| Total Codewords | `TotalCodewords` | W | ✓ | codewordArray.length |
| Data Codewords | `DataCodewords` | E→W (v1.29) | ✓ | q.symbols[0].dataCodewords (probe v1.29) |
| ECC Budget | `ErrorCorrectionBudget` | E→W (v1.29) | ✓ | q.symbols[0].ecCodewords (probe v1.29) |
| Errors Corrected | `ErrorsCorrected` | W | ✓ | Counted from codewordArray.isCorrected |
| Error Capacity Used | `ErrorCapacityUsed` | W | ✓ | ErrorsCorrected × 2 (confirmed v1.28) |
| EC Type | `ErrorCorrectionType` | W | ✓ | "ECC200" for DM; QR = probe v1.29 |
| ULQZ/URQZ/RUQZ/RLQZ (32×32+) | — | **E (firmware limit)** | ✓ PDF-only | Not in push channel on fw 6.1.16_sr4 |
| Per-quadrant TTR/RTR/TCT/RCT | — | **E (firmware limit)** | ✓ PDF-only | Not in push channel on fw 6.1.16_sr4 |

---

## 5. ISO 15415 parameter coverage (QR Code)

| Parameter | VTCCP | DMST | Notes |
|---|---|---|---|
| Upper-Left Finder Pattern Grade | W | ✓ | q.upperLeftPattern — confirmed v1.27 |
| Upper-Right Finder Pattern Grade | W | ✓ | q.upperRightPattern |
| Lower-Left Finder Pattern Grade | W | ✓ | q.lowerLeftPattern |
| Horizontal Clock Track Grade | W | ✓ | q.horizontalClockTrack |
| Vertical Clock Track Grade | W | ✓ | q.verticalClockTrack |
| Alignment Pattern Grade | W | ✓ | q.alignmentPatterns |
| Version Information Block Grade | W (`"-"`) | ✓ | `VIBGrade="-"` firmware literal — confirmed v1.26 loaded-image |
| Format Information Block Grade | W | ✓ | q.formatInformationBlock |
| Error Correction Level | E→probe (v1.29) | ✓ | q.symbols[0].errorCorrectionLevel — DebugSymbols probe v1.29 |
| QR Version | W (parser) | ✓ | Derived from MatrixSize (QR uses 8-module QZ offset) |
| Mask Pattern | Pending | ✓ | Not yet exposed via push channel |

*Note: All 8 QR grade params confirmed as top-level q keys, NOT in q.symbols[0]
(confirmed v1.26 DebugTrucheckKeys scan). Loaded-image QR returns sentinel grades;
live optical QR scan is still needed to observe non-sentinel measurement values.*

---

## 6. ISO 15416 parameter coverage (1D)

| Parameter | VTCCP | DMST | Notes |
|---|---|---|---|
| Symbol ANSI Grade | ✓ | ✓ | r.metrics.symbolAnsiGrade |
| Per-scan: Edge Contrast min | ✓ | ✓ | Up to 10 scan results per record |
| Per-scan: Modulation | ✓ | ✓ | |
| Per-scan: Defect | ✓ | ✓ | |
| Per-scan: Decodability | ✓ | ✓ | |
| Per-scan: DEC | ✓ | ✓ | |
| Avg LQZ / RQZ / HQZ | ✓ | ✓ | |
| BWG % | ✓ | ✓ | r.metrics.printGrowth.raw |
| Magnification | ✓ | ✓ | |
| Ratio (Code 39) | ✓ | ✓ | |
| Element Widths sub-tab | ✓ | ✓ PDF | VTCCP writes "Element Widths" sheet |

---

## 7. DMST features intentionally outside VTCCP scope

| DMST feature | In VTCCP? | Rationale |
|---|---|---|
| Live optical viewer / crosshair | No | Hardware UI; VTCCP is headless logging |
| Focus/trigger calibration wizard | No | DMST-only hardware interaction |
| Reader configuration (I/O, comms settings) | No | DMST job — write once at setup |
| System log / event viewer | No | Out of scope |
| Upgrade firmware | No | DMST-only |
| Report Settings PDF customization | No | DMST generates PDFs; VTCCP replaces that function |
| "Add to Favourites" / presets | No | Not applicable |
| Ethernet/serial port config | No | Hardware config, DMST job |

---

## 8. VCCS-distinctive VTCCP features (not in DMST)

| Feature | Status | Notes |
|---|---|---|
| VCCS-branded Excel report with VCCS navy header | Shipped (B1–B6) | 163-column TruCheckCompatible schema |
| Job / Roll / Operator / Batch tracking | Shipped | SessionState; SessionSidecar |
| Modulation Values sheet (B7) | In progress | ModValuesSheetWriter — variable-size grid |
| Codeword Values sheet (B7) | In progress | CwValuesSheetWriter — data/ECC boundary marker |
| OpticsSource flag (LiveScan / LoadedImage) | Shipped v1.25 | CU==−1 AND MRD==−1 discriminator |
| CalibrationWarning flag (FieldCalibrated=false) | Architecture (D1) | All observed scans: FieldCalibrated=false |
| Reverse-report from Excel (Excel→VTCCP round-trip) | Pending D2 | Blocked on D1 |
| GS1/MIL-STD application-syntax validation | Shipped | GS1 syntax engine v1.4.0 at vtccp/lib/gs1-syntax-engine/ |
| VCCS-distinctive HTML/PDF report | Pending D1 | 4 layout variants designed (dist/report-samples/) |
| Push-script authoring and version management | Ongoing | v1.28 device-confirmed; v1.29 in progress |
| IMAGE.LOAD scan pipeline (D4) | Architecture defined | Push parser OpticsSource + sidecar ready |

---

## 9. Report format comparison

| Aspect | DMST PDF | VTCCP Excel | VTCCP HTML (D1) |
|---|---|---|---|
| Logo / branding | Cognex | VCCS | VCCS |
| Grade color coding | DMST palette | VCCS palette | VCCS palette |
| Column scope | ~40 visible | 163 columns (all) | TBD (key fields) |
| Multi-scan history | One report per session | All scans one file | Per-job |
| Modulation grid image | ✓ rendered | ✓ Excel tab (B7) | Pending |
| Codeword table | ✗ | ✓ Excel tab (B7) | Pending |
| Reverse round-trip | ✗ | Pending D2 | N/A |
| Grade color scheme | DMST default | A=#d4edda, B=#cce5ff, C=#fff3cd, D=#ffe0b2, F=#f8d7da, X=#e2e3e5 | Same |

---

## 10. Key firmware limitations on DM475V fw 6.1.16_sr4

These are architectural limits confirmed by live scan probing — not bugs or gaps
in VTCCP.  No push-script change can expose these fields on this firmware.

| Limitation | Evidence | Impact |
|---|---|---|
| ULQZ/URQZ/RUQZ/RLQZ not in push | v1.27 DebugQZObjs: `{grade:"X", numericGrade:8.8}` — firmware-internal scale | 4 quadrant QZ grade columns always empty for this firmware |
| TTR/RTR/TCT/RCT per-quadrant not in push | v1.27 DebugTTRCTObjs: `{raw:-1, grade:X, numericGrade:8.8}` | 16 per-quadrant clock/transition columns always empty |
| q.general has exactly 7 keys | v1.27 DebugGeneral confirmed; no ECLevel, DataCW, ECBudget | DataCodewords/ECBudget probed via symbols[0] in v1.29 |
| DataCodewords/ECBudget not in q.general | v1.27 confirmed | v1.29 probes q.symbols[0] as alternate |
| QR ECLevel not in q.general | v1.27 confirmed | v1.29 probes symbols[0] + validation.gs1 + AIM modifier |
| q.minimumReflectance always sentinel | v1.27 raw=0, grade=F, numericGrade=0 on ALL scans | Firmware dead field; MinReflectance wired to q.reflectanceDark |
| FieldCalibrated always false | All v1.26–v1.28 scans | Report must flag as CalibrationWarning (D1 implementation) |
| DMV-8072V ≥32×32 quadrant grades | Not a DM475V limit — RUQZ/RLQZ available on DMV-8072V ≥32×32 | VTCCP schema ready; device upgrade path identified |

---

## 11. DPM mode parameter substitutions

When `q.overall.gradingStandard` does not start with "ISO 15415", the parser
routes to DPM-specific metrics.

| ISO 15415 param | DPM equivalent | VTCCP field |
|---|---|---|
| Symbol Contrast | Cell Contrast | `SC_Percent`, `SC_Grade` (sourced from q.cellContrast) |
| Modulation | Cell Modulation | `MOD_Grade` (sourced from q.cellModulation) |
| Fixed Pattern Damage | Present (same name) | `FPD_Grade`, `FPD_Value` |
| ANU, GNU | Present (same names) | Unchanged |

---

*Generated 2026-05-18. Next update trigger: v1.29 device-confirmed scan or D1 implementation.*
