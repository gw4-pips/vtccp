# Firmware Confirmed Facts — DM475V / fw 6.1.16_sr4

**Device**: DM475-63530E-PIPS-Verif-Lab
**Firmware**: 6.1.16_sr4
**Last updated**: 2026-05-24
**Status**: Living document — append new confirmed facts as probe results arrive.
**Authority**: All entries are device-confirmed from actual push XML output or DMCC responses,
not inferred from documentation.

---

## 1. Firmware object model — what is and is not present

### q.general key inventory (confirmed v1.29, 2026-05-19)

`q.general` contains exactly **7 keys** on fw 6.1.16_sr4:

| Key | Type | Notes |
|---|---|---|
| `encodedCharacters` | integer | Use this directly (not encodationAnalysisArray.length) |
| (5 additional unnamed keys) | — | Enumerated via DebugGeneralKeys probe; names not recorded |

**Dead paths in q.general** (confirmed empty on DM and QR):
`polarity`, `imagePolarity`, `image`, `colorMode`, `imageColor`,
`eci`, `ECI`, `eciMode`, `encodedChars` — ALL return empty string on fw 6.1.16_sr4.
Do not probe these again. `ImagePolarity` is definitively unresolvable from push XML.

### q.symbols (confirmed v1.29, 2026-05-19)

`q.symbols` is **null** on both DM and QR scans on fw 6.1.16_sr4.
`DebugSymbols` probe confirmed: the property exists in the object model but is not populated
by the firmware at verification time.

**Consequence**: `DataCodewords`, `ErrorCorrectionBudget`, and `ErrorCapacityUsed` cannot be
obtained from `q.symbols` on this firmware. Use C# table lookup instead (ECC200 table for DM,
QR size table for QR).

### r.symbology key inventory (QR scan, v1.32, 2026-05-24)

`r.symbology` on a QR Grade A scan contains exactly **9 keys**:

| Key | Type | Example value |
|---|---|---|
| `name` | string | `"QR"` |
| `id` | string | `"]Q1"` |
| `quality` | integer | `100` |
| `contrast` | decimal | `0.698` |
| `moduleSize` | decimal | `9.748` |
| `corners` | array (4 elements) | corner coordinate objects |
| `center` | object | center coordinate |
| `angle` | integer | `359` |
| `size` | object | dimensions object |
| `failureCode` | integer | `0` |

**No ECLevel, dataMask, or version key present.** `ECLevel` is definitively unresolvable
from push XML on fw 6.1.16_sr4 via any path. See §3 (dead paths) for complete list.

### r.image key inventory (confirmed v1.32, 2026-05-24)

`r.image` contains camera **hardware metadata only** — no symbol quality data.

| Key | Description |
|---|---|
| `id` | Device ID string |
| `FoV` | Field of View object |
| `RoI` | Region of Interest object |
| `exposureTime` | Camera exposure setting |
| `gain` | Camera gain setting |
| `autoExposure` | Auto-exposure flag |
| `illEnabled` | Illumination enabled flag |
| `focusLength` | Focus distance |
| `mirrorAngle` | Mirror position (2-axis DPM illumination) |
| `intIllWavelength` | Internal illumination wavelength |
| `multicolorWavelength` | Multicolor illumination wavelength (if equipped) |
| `multicolorRatio` | Multicolor channel mix ratio |
| `multicolorName` | Multicolor preset name |

`r.image` contains **NO** `ImagePolarity`, `ECI`, `ECLevel`, or `DataMaskPattern`.
These 4 fields are definitively unresolvable from push XML on fw 6.1.16_sr4.

---

## 2. AIM ID modifier semantics (confirmed v1.29, 2026-05-19)

| AIM ID | Observed modifier | Modifier encodes | NOT |
|---|---|---|---|
| `]d1` | `1` | ASCII encoding type | ECLevel |
| `]Q1` | `1` | ECI **presence** (000003 = Latin-1) | ECLevel |

**Prior incorrect assumption**: modifier bits[1:0] = ECLevel.
**Confirmed behavior**: For QR, modifier `1` means ECI data is present; the ECI value itself
(e.g. `000003`) is shown in DMST but is NOT present anywhere in push XML on fw 6.1.16_sr4.

The `DebugValidationGS1` probe confirmed `ecLvl=absent` on both DM and QR.

---

## 3. Dead paths — definitively unresolvable fields on fw 6.1.16_sr4

These fields cannot be obtained from push XML on this firmware version.
Do not add new probes for them. Do not allocate space in report output.

| Field | Paths tried | Why dead |
|---|---|---|
| `ECLevel` (QR) | `q.symbols[0].ecLevel`, `r.validation.ecLvl`, AIM modifier bits, `r.symbology` | `q.symbols=null`, validation path absent, AIM modifier=ECI presence not ECLevel, `r.symbology` has no ecLevel key |
| `DataMaskPattern` (QR) | `r.symbology`, `q.general` | Not present in any of the 9 `r.symbology` keys |
| `QR_Version` | `r.symbology.size`, `q.general` | `r.symbology.size` is a dimensions object; version not exposed |
| `ImagePolarity` | `gnProp("polarity"/"imagePolarity"/"image")`, `r.image` | All q.general paths empty; r.image = hardware metadata only |
| `ECI` (value) | `gnProp("eci"/"ECI"/"eciMode")`, `r.symbology` | All paths empty; ECI value visible in DMST is not surfaced in push |

---

## 4. ANUPercent behavior (Bug #9 — resolved v1.32, 2026-05-24)

**The problem (v1.30 bug list #9)**: Prior data showed 73.7→0.7% (implied ÷100 formula)
but a grade-D scan returned 3.9→3.9% (1:1 ratio) — two different sub-keys being read
depending on scan context.

**v1.32 fix**: Push script explicitly reads `.percent` sub-key of `q.metrics.axialNonUniformity`.

**Device-confirmed results**:

| Scan | Raw push value | DMST value | Ratio | Script |
|---|---|---|---|---|
| QR GUID (scan #2, v1.29) | 73.7 | 0.7% | ÷100 | v1.29 (reading wrong sub-key) |
| QR grade-D (scan #10, v1.29) | 3.9 | 3.9% | 1:1 | v1.29 (reading wrong sub-key, coincidental) |
| QR GUID v1.32 (scan #11) | 0.8392218351364136 | 0.8392% | 1:1 | v1.32 (correct `.percent` sub-key) ✓ |

**Confirmed behavior**: The `.percent` sub-key returns the percentage value directly (0–100 scale
where 0.84% is a very good grade-A result). Do NOT apply ÷100 to the `.percent` sub-key.

---

## 5. SymbologyName in push XML vs DMST

| Symbology | DMST label | Push XML `SymbologyName` | ClassifySymbology key needed |
|---|---|---|---|
| Data Matrix | "Data Matrix" | `"DataMatrix"` | `"DataMatrix"` |
| QR Code | "QR Code" | `"QR"` | `"QR"` (added v1.32) |
| EAN-13 | "EAN-13" | `"EAN13"` | `"EAN13"` |
| Code 128 | "Code 128" | `"Code128"` | `"Code128"` |

The firmware push script emits `SymbologyName="QR"` (not `"QR Code"`). `VerificationXmlMap`
requires a `"QR"` entry in `ClassifySymbology` to resolve to `SymbologyFamily.QRCode`.
This was added in v1.32 parser update (2026-05-24).

---

## 6. QR Code grade behavior on QR scan (confirmed scan #11, v1.32, 2026-05-24)

### QR pattern grades (ISO 15415 QR, params 7–14)

All 8 QR pattern grade elements are **present and populated** when `SymbologyName="QR"`:

| Element | Description | Grade A scan value |
|---|---|---|
| `ULPGrade` | Upper-Left Finder Pattern | A |
| `URPGrade` | Upper-Right Finder Pattern | A |
| `LLPGrade` | Lower-Left Finder Pattern | A |
| `HCTGrade` | Horizontal Clock Track | A |
| `VCTGrade` | Vertical Clock Track | A |
| `ALPGrade` | Alignment Pattern (v2+) | A |
| `VIBGrade` | Version Information Blocks | `"-"` (v3 QR has no VIB — firmware literal) |
| `FIBGrade` | Format Information Blocks | A |

### VIBGrade special values

| Condition | Push value | Meaning |
|---|---|---|
| QR v1 or v3 (no version info blocks) | `"-"` | Firmware literal dash — not a letter grade |
| QR v2+ verification failure | `"F"` | Standard grade letter |
| QR v2+ passing | Letter grade A–D | Standard grade letter |

### DM-only grades on QR scan

When `SymbologyName="QR"`, these DM-specific elements are emitted as `"X"` (not null/empty):

`LLSGrade`, `BLSGrade`, `LQZGrade`, `BQZGrade`, `TQZGrade`, `RQZGrade`,
`TTRGrade`, `RTRGrade`, `TCTGrade`, `RCTGrade`

Parser receives `GradingResult{letter="X"}` — not null. Do not treat `"X"` as a parse failure.

---

## 7. Firmware-emitted sentinel values

| Value | Field(s) | Meaning |
|---|---|---|
| `"X"` | DM-only grades on QR scan | Grade not applicable to this symbology |
| `"-"` | `VIBGrade` on v1/v3 QR | No version information blocks in this QR version |
| `"F"` | `VIBGrade` on total-fail QR scan | Firmware standard fail grade |
| `8.8` | Any numeric grade field on fail scan | Internal Cognex fail marker (not a grade) |
| `""` (empty) | `MatrixSize`, `EncodedCharacters`, `TotalCodewords` on fail | DMST shows −1 for same fields |
| `−1` | `ContrastUniformity`, `MRD` on loaded-image scan | OpticsSource discriminant |
| `"X"` | `LLS_Grade`, `BLS_Grade`, `LQZ_Grade`, etc. on LoadedImage | Not assessed for loaded images |

---

## 8. NominalXDim format in push XML (confirmed v1.32, 2026-05-24)

Push XML emits `NominalXDim` as a **string with units**: `"12.6 mil"`.
The CharData path (fw6.x General Characteristics section) also emits `"13.1 mil"`.
Both require unit stripping via split-on-space, taking the first token.

Fixed in `DmstResultParser.cs` v1.32 update: fallback now covers both paths.

Confirmed values:
- `"12.6 mil"` → `12.6m` (DM475V, QR GUID Grade A, 2026-05-24)
- `"13.1 mil"` → `13.1m` (DM475V, DM live scan, 2026-05-19)

---

## 9. FieldCalibrated on all observed scans

`FieldCalibrated = false` on every observed scan to date (11 scans across all symbologies
and grades). `CalibrationDate` is present (`"5/20/2026 1:14:58 AM"` on scan #11).

The device has been factory-calibrated (CalibrationDate present) but not field-calibrated.
This requires a **CalibrationWarning** in the D1 report output.

---

## 10. OpticsSource on QR live scans (confirmed 2026-05-19)

**Counterintuitive**: Live QR scans on fw 6.1.16_sr4 return `ContrastUniformity=−1` and
`MRD=−1`, causing the OpticsSource script to classify them as `"LoadedImage"`.

This is **firmware behavior**, not an indicator that the symbol was actually loaded from a file.

**Implication for D4**: The `ContrastUniformity=−1 && MRD=−1` discriminant CANNOT distinguish
a live QR scan from an IMAGE.LOAD QR scan on this firmware. A different discriminant is needed
for D4. Candidate: require operator to tag `LoadedImageTruCheck` explicitly in the UI.

See `optics-source-model.md` §2 and §4 for the full discriminant model.

---

## 11. ErrorCapacityUsed derivation

`ErrorCapacityUsed = ErrorsCorrected × 2` confirmed on DM scan (7×2=14) and QR scan (0×2=0).
The firmware does not expose this value directly; derive it in C# when needed.

---

## 12. modulation and codeword array dimensions (DM, confirmed v1.28–v1.29)

| Symbol | Array | Length | Formula |
|---|---|---|---|
| DM 16×36 GS1 | `modulationArray` | 684 | (16+2)×(36+2) = 18×38 = 684 (symbol + 1-module QZ border each side) |
| DM 16×36 GS1 | `codewordArray` | 56 | 32 data + 24 ECC |
| DM 16×36 GS1 | `encodationAnalysisArray` | 33 | = EncodedCharacters ✓ |
| QR 29×29 (v3) | `modulationArray` | 1369 | (29+8)×(29+8) = 37×37 = 1369 (4-module QZ each side) |

`modulationArray[i].grade` = `"("` — single character, not truncated. The grade letter
is the character in the grade string, not an index.

---

## 13. Live scan catalog — 11 scans (updated 2026-05-24)

| # | Date | Script | Symbology | Grade | Key finding |
|---|---|---|---|---|---|
| 1 | 2026-05-19 | v1.29 | DM plain | A | `]d1`, `qrBranch=0` baseline |
| 2 | 2026-05-19 | v1.29 | QR GUID | A | `]Q1`, all 8 QR grades confirmed |
| 3 | 2026-05-19 | v1.29 | QR URL | F | Total fail; `VIBGrade=F` on fail |
| 4 | 2026-05-19 | v1.29 | QR GUID2 defocus | F | `LLPGrade=F` confirmed |
| 5 | 2026-05-19 | v1.29 | QR GUID2 in-focus | A | Refocus pass |
| 6 | 2026-05-19 | v1.29 | DM GS1 22×22 no-parser | D | ANU grade fail |
| 7 | 2026-05-19 | v1.29 | DM GS1 22×22 with-parser | D | AI-level GS1 parsing confirmed |
| 8 | 2026-05-19 | v1.29 | EAN-13 | A | ISO 15416:2016, `]E0` |
| 9 | 2026-05-19 | v1.29 | Code 128 | A | `]C0`; GS1 parser on; non-GS1 data → format fail; ISO grade unaffected |
| 10 | 2026-05-19 | v1.29 | QR grade D | D | `]Q1`; MOD=D, RM=D, DecodeGrade=F; ANUPercent=3.9→3.9% (broke v1.29 formula) |
| 11 | 2026-05-24 | v1.32 | QR GUID | A | v1.32 confirmed; ANUPercent=0.8392; `DebugRImage` closes 4 missing fields |

---

## 14. Push script version history — device-confirmed installs

| Version | Confirmed | Symbology | Grade | Key changes / probes answered |
|---|---|---|---|---|
| v1.24 | 2026-05-18 | DM GS1 | — | Baseline; 10 probes confirmed |
| v1.28 | 2026-05-18 | DM GS1 | — | Post-B4 parser wiring confirmed |
| v1.29 | 2026-05-19 | DM + QR | A/F/D | 11 scans; q.symbols=null; AIM modifier=ECI presence; QR grades confirmed |
| v1.30 | — | — | — | VOIDED — written 2026-05-20, never installed; superseded by v1.32 |
| v1.31 | — | — | — | VOIDED — written 2026-05-20, never installed; superseded by v1.32 |
| v1.32 | 2026-05-24 | QR GUID | A | Bug #9 (ANUPercent) confirmed fixed; `DebugRImage` answered; 4 fields closed |

---

## 15. Probe results reference

### DebugRImage (v1.32, scan #11, 2026-05-24) — CLOSED

`r.image` = camera hardware metadata object. See §1 for key inventory.
**Finding**: No symbol quality data in `r.image`. ImagePolarity, ECI, ECLevel, DataMaskPattern
are all definitively absent. This probe is complete; do not re-run.

### DebugRSymbology (v1.29/v1.32, QR scan, 2026-05-19/05-24) — CLOSED

`r.symbology` = 9 keys. See §1 for key inventory. No ECLevel or DataMaskPattern.
This probe is complete; do not re-run.

### DebugImagePolarity (v1.32, QR scan, 2026-05-24) — CLOSED

`q.general` has 7 keys. Probed: `polarity`, `imagePolarity`, `image`, `colorMode`,
`imageColor`, `eci`, `ECI`, `eciMode`, `encodedChars` — all empty.
This probe is complete; do not re-run.

### DebugSymbols (v1.29, DM + QR, 2026-05-19) — CLOSED

`q.symbols = null` on both symbologies. This probe is complete; do not re-run.

### DebugValidationGS1 (v1.29, DM + QR, 2026-05-19) — CLOSED

`ecLvl = absent` on both DM and QR. This probe is complete; do not re-run.

### DebugANUSubkeys (v1.32 carried forward) — OPEN

The sub-key probe is in v1.32 push script but the sub-key name was not separately
surfaced in the v1.32 output. The correct fix (reading `.percent`) is confirmed working
from ANUPercent=0.8392 result. Sub-key name documentation is informational only.

### DebugValidationKeys (v1.32 carried forward) — OPEN

Result not yet received. Probe is live in v1.32.

---

## 16. Fields present in push XML but NOT in DMST PDF

These fields are emitted by the push script and captured in `VerificationRecord` but do not
appear in the standard DMST export PDF. They are VTCCP-exclusive data:

| Field | Notes |
|---|---|
| `OpticsSource` | Derived by push script from CU/MRD discriminant |
| `JpegImageBase64` | Full verification image; DMST shows thumbnail only |
| `FieldCalibrated` / `FactoryCalibrated` | DMST shows CalibrationDate only |
| `ModuleSizePx` | From `r.symbology.moduleSize`; DMST shows NominalXDim only |
| `SymbolQuality` | Decoder confidence (0–100); not in DMST report |
| `SymbologyId` | AIM ID string (e.g. `]d1`); not in DMST report |
| `PushScriptDiag` | Version tag; internal only |
| `DeviceName` (from `r.source`) | Device name at time of scan |
| All `Debug*` elements | Probe output; discarded after answering |

---

## Appendix A — Fields that require C# table lookup (firmware does not provide)

| Field | Reason | Lookup table |
|---|---|---|
| `DataCodewords` | `q.symbols = null` | ECC200 table (DM) + QR codeword table by version |
| `ErrorCorrectionBudget` | `q.symbols = null` | Same table |
| `ErrorCapacityUsed` | Not exposed | `ErrorsCorrected × 2` (confirmed derivation) |
| `QR_Version` | Not in `r.symbology` | Derive from `MatrixSize` (e.g. 29×29 → v3) |
| `QR_ECLevel` | Definitively absent from all push paths | Cannot be derived — omit from report |
| `QR_MaskPattern` | Definitively absent from all push paths | Cannot be derived — omit from report |

---

## Appendix B — Tech debt and open items

| Item | Status | Notes |
|---|---|---|
| QR pattern grades in Excel sheet | PENDING | Parsed into VerificationRecord (v1.32) but ExcelWriter sheet mapper not yet updated |
| IMAGE.LOAD QR scan | PENDING (D4) | Needed to confirm D4 stored-image discriminant (CU/MRD cannot distinguish live QR from loaded QR) |
| ISO 29158 / DPM schema | PENDING (C1) | Requires external document |
| Sensor/frame metadata UI | UNBLOCKED | Plan in `references/architecture/sensor-frame-metadata-plan.md` (check if exists) |
| Report HTML/PDF (D1) | BLOCKED on C3 | Webscan TruCheck docs required |
| CalibrationWarning in D1 | PENDING | `FieldCalibrated=false` on all 11 scans to date — definite risk |
