# Firmware Confirmed Facts — DM475V / fw 6.1.16_sr4

**Device**: DM475-63530E-PIPS-Verif-Lab
**Firmware**: 6.1.16_sr4
**Last updated**: 2026-05-25 (Wireshark protocol analysis complete)
**Status**: Living document — append new confirmed facts as probe results arrive.
**Authority**: All entries are device-confirmed from actual push XML output or DMCC responses,
not inferred from documentation.

---

## 1. Firmware object model — what is and is not present

### q.general key inventory (EXACT, confirmed scan #12, 2026-05-24)

`q.general` contains exactly **7 keys** on fw 6.1.16_sr4. The DebugImagePolarity `allKeys`
field from scan #12 (DM live) confirms the complete exact key set:

| Key | Type | Example (DM live) | Example (QR IMAGE.LOAD) |
|---|---|---|---|
| `verticalBWG` | number | `11` | `-4` |
| `horizontalBWG` | number | `11` | `-3` |
| `xDimension` | string | `"20.3 mil"` | `"12.6 mil"` |
| `contrastUniformity` | number | `73` | `-1` |
| `contrastUniformityRow` | number | `12` | `-1` |
| `contrastUniformityCol` | number | `17` | `-1` |
| `MRD` | number | `67` | `-1` |

**CRITICAL CORRECTION (2026-05-25)**: `encodedCharacters` is **NOT** one of the 7 keys.
Prior documentation assumed it was present based on DebugGeneralKeys (v1.29) which did not
record all key names. Confirmed absent: `gnProp("encodedCharacters")` returns `""` on both
DM and QR, causing fallback to `encodationAnalysisArray.length` which is also wrong
(DM: eaLen=33, DMST=38; QR: eaLen=39, DMST=36). `EncodedCharacters` is definitively
unresolvable from push XML on fw 6.1.16_sr4 — see Appendix A.

**Dead paths in q.general** (confirmed empty on DM and QR):
`polarity`, `imagePolarity`, `image`, `colorMode`, `imageColor`,
`eci`, `ECI`, `eciMode`, `encodedChars`, `encodedCharacters` — ALL absent on fw 6.1.16_sr4.
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

**No ECLevel, dataMask, or version key present at top-level.** See §1 nested inventory
and §3 for complete path inventory. All paths exhausted as of v1.33 (2026-05-25).

### r.symbology nested key inventory (confirmed v1.33, scans #14/#15, 2026-05-24)

The v1.33 `DebugRSymbologyNested` probe enumerated sub-keys of the three nested objects
within `r.symbology`. No ECLevel, DataMaskPattern, ECI, or ImagePolarity anywhere.

**`r.symbology.size` sub-keys:**

| Key | DM 16×36 | QR 29×29 | Semantics |
|---|---|---|---|
| `x` | `36` | `29` | width in modules (columns for DM) |
| `y` | `16` | `29` | height in modules (rows for DM) |

QR version derivable: `(size.x − 21) / 4 + 1` → 29→v3 ✓ (redundant with MatrixSize string parse)

**`r.symbology.center` sub-keys:**

| Key | DM | QR | Semantics |
|---|---|---|---|
| `x` | `1553` | `201` | pixel X of symbol center within image frame |
| `y` | `485` | `199` | pixel Y of symbol center within image frame |

**`r.symbology.corners[0]` sub-keys:**

| Key | DM | QR | Semantics |
|---|---|---|---|
| `x` | `1271` | `57` | pixel X of first corner |
| `y` | `587` | `59` | pixel Y of first corner |

All 9 top-level keys + all nested sub-keys of `r.symbology` are now fully enumerated.
No decode-structural data (ECLevel, DataMaskPattern, version number) exposed anywhere.
Center/corner pixel coordinates are new data — potentially useful for D4 image annotation.

### r.image key inventory (COMPLETE 28-key, confirmed scan #12 DM live, 2026-05-24)

`r.image` contains camera **hardware metadata only** — no symbol quality data.
Full key inventory from DebugRImage on DM live scan (scan #12):

| Key | Type / Example (DM live) | Example (QR IMAGE.LOAD) |
|---|---|---|
| `id` | number `4007` | `0` |
| `index` | number `0` | `0` |
| `FoV` | `[obj]` | `[obj]` |
| `RoI` | `[obj]` | `[obj]` |
| `exposureTime` | number `32` | `0` ← secondary LiveScan discriminator |
| `gain` | number `1.00` | `0.00` |
| `autoExposure` | boolean `true` | `false` ← secondary LiveScan discriminator |
| `illEnabled` | boolean `true` | `false` |
| `illIntensity` | number `1` | — |
| `extIllEnabled` | boolean `true` | — |
| `extIllIntensity` | number `1` | — |
| `focusLength` | number `0` | — |
| `focusPower` | number `0.00` | — |
| `setupIndex` | number `0` | — |
| `inputStates` | `[arr.4]` | — |
| `filterTime` | number `0` | — |
| `creationTime` | number `435462838` | — |
| `creationTicks` | number `0` | — |
| `creationDate` | `[obj]` | — |
| `ptpTimeStamp` | `[obj]` | — |
| `mirrorAngleA` | number `0.000` | — |
| `mirrorAngleB` | number `0.000` | — |
| `mirrorPathIndex` | number `-1` | — |
| `intIllWavelength` | number `65535` (= 0xFFFF sentinel) | — |
| `intIllWavelengthMask` | number `0` | — |
| `multicolorWavelength` | `[arr.0]` (not equipped) | — |
| `multicolorRatio` | `[arr.0]` | — |
| `multicolorName` | `[arr.0]` | — |

**Secondary OpticsSource discriminator** (supplementary to CU/MRD=-1 primary):
`r.image.exposureTime=0` and `r.image.autoExposure=false` on IMAGE.LOAD scans.
Use CU=-1 AND MRD=-1 as the primary discriminant; r.image fields are secondary confirmation only.

`r.image` contains **NO** `ImagePolarity`, `ECI`, `ECLevel`, or `DataMaskPattern`.
These 4 fields are definitively unresolvable via `r.image` on fw 6.1.16_sr4.

### r-sibling complete inventory (confirmed QR scan #13, 2026-05-24)

`r` has the following top-level properties (from DebugRSiblings on scan #13):

| Key | Type | Notes |
|---|---|---|
| `decoded` | boolean | `true` on successful decode |
| `content` | string | Decoded data (first 30 chars in DebugRSiblings) |
| `decodeTime` | number | `402` (QR GUID IMAGE.LOAD) |
| `triggerTime` | number | `418` |
| `timeout` | number | `2000` |
| `readSetup` | number | `0` |
| `symbology` | object | = `r.symbology` (9 keys, probed) |
| `image` | object | = `r.image` (28 keys, probed) |
| `validation` | object | = `r.validation` (ecLvl=absent confirmed) |
| `source` | string | Device name, e.g. `"DM475-63530E-PIPS-Verif-Lab"` |
| `annotation` | string | Empty string on all observed scans |
| `label` | string | Empty string |
| `custom_svg` | string | Empty string |
| `barcodeAssignment` | object | **NEVER PROBED — v1.34 candidate** |
| `trucheck` | object | `r.trucheck` (all grade metrics — fully probed) |
| `metrics` | object | `r.metrics` (excluded by DebugRSiblings filter) |

**Total r properties**: 16 (14 shown by DebugRSiblings + trucheck + metrics excluded by filter).

**NEW UNEXPLORED PATH**: `r.barcodeAssignment=[obj]` — revealed by scan #13 DebugRSiblings.
Name suggests assignment/type configuration rather than decode results, but has never been
probed. Queue for v1.34 if v1.33 FIB/VIB probes do not resolve ECLevel/DataMaskPattern.

**Note — DM DebugRSiblings parse failure**: The DM scan #12 DebugRSiblings element is present
in the XML but its value contains GS1 control characters (`<0x1E>`, `<0x1D>`) from the
`content` field, which break naive XML regex parsers. VTCCP's streaming push XML parser handles
this correctly. The DM r-sibling inventory is expected to be identical to QR.

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

## 3. Fields permanently unresolvable from push XML on fw 6.1.16_sr4

**PROBE CAMPAIGN COMPLETE — 2026-05-25.** v1.33 results (scans #14/#15) have exhausted
all known firmware paths. All four open fields are confirmed permanently unresolvable
from the push script's JS scope on fw 6.1.16_sr4.

The firmware **does** know these values internally — it uses ECLevel and DataMaskPattern
to decode and grade QR codes, and it applies ImagePolarity during acquisition. However,
the JS push script environment exposes only grade results, not the decoded structural data.

### 3a — Complete path inventory (all exhausted)

| Field | All paths tried | Final status |
|---|---|---|
| `ECLevel` (QR) | `q.symbols=null`; `r.validation.ecLvl=absent`; AIM modifier=ECI presence not ECLevel; `r.symbology` 9 top-level keys=no ecLevel; `q.formatInformationBlock` sub-keys=`{grade, numericGrade}` only | **PERMANENTLY UNRESOLVABLE** |
| `DataMaskPattern` (QR) | `r.symbology` top-level=absent; `q.general` 7 keys=absent; `q.formatInformationBlock` sub-keys=`{grade, numericGrade}` only | **PERMANENTLY UNRESOLVABLE** |
| `ECI` value (e.g. `000003`) | `gnProp(eci/ECI/eciMode/encodedChars)=empty`; `r.symbology` top-level=absent; AIM modifier encodes presence not value; no sub-key path found | **PERMANENTLY UNRESOLVABLE** |
| `ImagePolarity` | `gnProp(polarity/imagePolarity/image/colorMode/imageColor)=empty`; `r.image=hardware metadata only`; not a decode property anywhere in JS scope | **PERMANENTLY UNRESOLVABLE** |

### 3b — Why q.formatInformationBlock was the final path

The QR Format Information Block encodes both ECLevel (2 bits) and DataMaskPattern (3 bits)
in the bitstream. `q.formatInformationBlock` is a `TrucheckMetric` object. Prior probes had
only read `.grade` (the letter). v1.33 enumerated ALL sub-keys:

**`q.formatInformationBlock` sub-keys (confirmed v1.33, scan #15, QR Grade A):**

| Key | Value | Notes |
|---|---|---|
| `grade` | `"A"` | Grade letter — already read by all prior versions |
| `numericGrade` | `4` | Float grade before letter mapping |

Only 2 sub-keys. The firmware grades the FIB but does not expose the decoded ECLevel
or DataMaskPattern values through the JS scope. **Path exhausted.**

**`q.versionInformationBlock` sub-keys (confirmed v1.33, scan #15, QR v3):**

| Key | Value | Notes |
|---|---|---|
| `grade` | `"-"` | No VIB for QR versions 1–6 |
| `numericGrade` | `4.5` | Firmware sentinel: "not applicable, passing" (above max 4.0) |

Only 2 sub-keys. No version number sub-key. **Path exhausted.**

### 3c — r.barcodeAssignment (only remaining unexplored object)

`r.barcodeAssignment=[obj]` was revealed by scan #13 and has never been probed.
Given the complete FIB/VIB result, this is the only remaining unexplored object in scope.
Name strongly suggests setup configuration rather than decode results.
**Queue as v1.34 low-priority probe; do not expect ECLevel/DataMaskPattern here.**

### 3d — Schema implications

Per Appendix A, these four fields must be omitted from push XML and marked as
"not available from push channel on fw 6.1.16_sr4" in VTCCP schema documentation.
Do not allocate Excel column space for permanently-empty fields.

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

## 13. Live scan catalog — 13 scans (updated 2026-05-25)

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
| 11 | 2026-05-24 | v1.32 | QR GUID (IMAGE.LOAD) | A | v1.32 confirmed; ANUPercent=0.8392; `DebugRImage` closes 4 missing fields |
| 12 | 2026-05-24 | v1.32 | DM GS1 16×36 (LIVE) | D | q.general exact 7-key names; EncodedCharacters dead (push=33,DMST=38); r.image 28-key inventory; OpticsSource=LiveScan (CU=73,MRD=67) |
| 13 | 2026-05-24 | v1.32 | QR GUID f0cffb39 (IMAGE.LOAD) | A | r-sibling 16-prop inventory; r.barcodeAssignment discovered; EncodedCharacters dead (push=39,DMST=36); OpticsSource=LoadedImage (CU=-1,MRD=-1) |
| **14** | **2026-05-24** | **v1.33** | **DM GS1 16×36 (LIVE)** | **D** | **v1.33 confirmed; DebugQFIBKeys=(not QR) ✓; DebugQVIBKeys=(not QR) ✓; DebugRSymbologyNested: size.x=36,y=16; center.x=1553,y=485; corners[0].x=1271,y=587** |
| **15** | **2026-05-24** | **v1.33** | **QR GUID f0cffb39 (IMAGE.LOAD)** | **A** | **PROBE CAMPAIGN COMPLETE: DebugQFIBKeys={grade=A,numericGrade=4} only — ECLevel/DataMask NOT exposed; DebugQVIBKeys={grade=-,numericGrade=4.5}; DebugRSymbologyNested: size.x=29,y=29; center.x=201,y=199; corners[0].x=57,y=59** |

---

## 14. Push script version history — device-confirmed installs

| Version | Confirmed | Symbology | Grade | Key changes / probes answered |
|---|---|---|---|---|
| v1.24 | 2026-05-18 | DM GS1 | — | Baseline; 10 probes confirmed |
| v1.28 | 2026-05-18 | DM GS1 | — | Post-B4 parser wiring confirmed |
| v1.29 | 2026-05-19 | DM + QR | A/F/D | 11 scans; q.symbols=null; AIM modifier=ECI presence; QR grades confirmed |
| v1.30 | — | — | — | VOIDED — written 2026-05-20, never installed; superseded by v1.32 |
| v1.31 | — | — | — | VOIDED — written 2026-05-20, never installed; superseded by v1.32 |
| v1.32 | 2026-05-24 | QR GUID + **DM GS1** | A + **D** | Bug #9 fixed; DebugRImage answered; q.general exact 7 keys named; EncodedCharacters dead path confirmed; r.image 28-key + r-sibling 16-prop full inventories; r.barcodeAssignment discovered |
| v1.33 | **2026-05-24** | **DM GS1 + QR GUID** | **D + A** | **DEVICE CONFIRMED (scans #14/#15). PROBE CAMPAIGN COMPLETE. FIB={grade,numericGrade} only — ECLevel/DataMask permanently unresolvable. r.symbology nested: size.x/y, center.x/y, corners[0].x/y confirmed.** |

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

### DebugRSymbologyNested (v1.33, scans #14/#15, 2026-05-24) — CLOSED

`r.symbology` nested objects fully enumerated. See §1 for complete table.
Finding: size.x/y = module dimensions; center.x/y and corners[0].x/y = pixel coordinates.
No ECLevel, DataMaskPattern, or version number in any sub-key. Probe complete; do not re-run.

### DebugQFIBKeys (v1.33, scan #15 QR, 2026-05-24) — CLOSED ★ DECISIVE

`q.formatInformationBlock` = `TrucheckMetric` with exactly **2 sub-keys**: `grade` and `numericGrade`.
- QR Grade A: `grade="A"`, `numericGrade=4`
- DM (non-QR): correctly emitted `(not QR)` per branch guard

The firmware computes the FIB grade internally using ECLevel + DataMaskPattern but does
NOT expose those values in the JS scope. **ECLevel and DataMaskPattern are permanently
unresolvable from push XML on fw 6.1.16_sr4.** Probe complete; do not re-run.

### DebugQVIBKeys (v1.33, scan #15 QR, 2026-05-24) — CLOSED

`q.versionInformationBlock` = `TrucheckMetric` with exactly **2 sub-keys**: `grade` and `numericGrade`.
- QR v3 (no VIB): `grade="-"`, `numericGrade=4.5` (firmware "not applicable" sentinel)
- DM (non-QR): correctly emitted `(not QR)` per branch guard

No version number sub-key. QR version remains derivable only from `MatrixSize` string
(or new: `r.symbology.size.x/y` integers). Probe complete; do not re-run.

### DebugANUSubkeys (v1.32 carried forward) — INFORMATIONAL

The sub-key probe confirmed `.percent` is the correct sub-key (ANUPercent=0.8392 working).
Sub-key name documentation is informational only — no action required.

### DebugValidationKeys (v1.32 carried forward) — INFORMATIONAL

Result not separately surfaced; all validation sub-keys superseded by FIB/VIB probe campaign.
No action required.

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

| Field | Reason | Lookup table / Status |
|---|---|---|
| `DataCodewords` | `q.symbols = null` | ECC200 table (DM) + QR codeword table by version |
| `ErrorCorrectionBudget` | `q.symbols = null` | Same table |
| `ErrorCapacityUsed` | Not exposed | `ErrorsCorrected × 2` (confirmed derivation) |
| `QR_Version` | Not in `r.symbology` top-level keys | Derive from `MatrixSize` (e.g. 29×29 → v3) |
| `EncodedCharacters` | `q.general.encodedCharacters` ABSENT (not one of the 7 keys); `encodationAnalysisArray.length` fallback also wrong (DM: 33 vs DMST 38; QR: 39 vs DMST 36) | **No known derivation from push XML on fw 6.1.16_sr4** — omit or leave empty; cannot be accurately computed without full encodation analysis parse |
| `QR_ECLevel` | All paths exhausted (v1.33 confirmed): q.symbols=null; r.validation=absent; AIM modifier=ECI presence; r.symbology top-level=absent; FIB sub-keys={grade,numericGrade} only — ECLevel not exposed | **PERMANENTLY UNRESOLVABLE** from push XML on fw 6.1.16_sr4 — omit from schema |
| `QR_MaskPattern` | All paths exhausted (v1.33 confirmed): FIB sub-keys={grade,numericGrade} only — DataMaskPattern not exposed | **PERMANENTLY UNRESOLVABLE** from push XML on fw 6.1.16_sr4 — omit from schema |
| `ECI` (value) | gnProp paths all empty; r.symbology top-level absent; FIB sub-keys absent | **PERMANENTLY UNRESOLVABLE** — omit from schema |
| `ImagePolarity` | gnProp paths all empty; r.image = hardware metadata only | **PERMANENTLY UNRESOLVABLE** — omit from schema |

---

## Appendix B — Tech debt and open items

| Item | Status | Notes |
|---|---|---|
| QR pattern grades in Excel sheet | PENDING | Parsed into VerificationRecord (v1.32) but ExcelWriter sheet mapper not yet updated |
| IMAGE.LOAD QR scan | **COMPLETE** (scans #13/#15) | CU=-1/MRD=-1 confirmed as LoadedImage discriminant on fw 6.1.16_sr4. OpticsSource logic working. |
| ISO 29158 / DPM schema | PENDING (C1) | Requires external document |
| Sensor/frame metadata UI | UNBLOCKED | Plan in `references/architecture/sensor-frame-metadata-plan.md` (check if exists) |
| Report HTML/PDF (D1) | BLOCKED on C3 | Webscan TruCheck docs required |
| CalibrationWarning in D1 | PENDING | `FieldCalibrated=false` on all 11 scans to date — definite risk |

---

## 16. DMCC SYMBOL.RESULT FULL = push script XML verbatim — CONFIRMED CLOSED (2026-05-25)

**Probe**: Raw TCP `GET SYMBOL.RESULT` after scan #15 (QR GUID f0cffb39, IMAGE.LOAD, Grade A).
**Result**: Response is the push-script-generated XML, returned verbatim by the firmware.
**Proof**: Response contains `<PushScriptDiag>v1.33 q=r.trucheck m=found</PushScriptDiag>` —
the push script's own diagnostic tag. The firmware does not maintain a separate "native DMCC XML"
format; `GET SYMBOL.RESULT` returns whatever the active push script last emitted.

**All four target fields empty in DMCC SYMBOL.RESULT FULL (confirmed):**

| Field | XML tag | Value |
|---|---|---|
| ECLevel | `<ErrorCorrectionType>` | `""` (empty) |
| DataMaskPattern | (no tag) | absent |
| ECI | (no tag) | absent |
| ImagePolarity | `<ImagePolarity>` | `""` (empty) |

**Additional confirmed-empty fields in SYMBOL.RESULT FULL:**

| Field | XML tag | Value |
|---|---|---|
| DataCodewords | `<DataCodewords>` | `""` (empty) |
| ErrorCorrectionBudget | `<ErrorCorrectionBudget>` | `""` (empty) |
| SCPercent | `<SCPercent>` | `""` (empty) |

**EncodedCharacters discrepancy confirmed**: `<EncodedCharacters>39</EncodedCharacters>` in both
push XML and SYMBOL.RESULT — both say 39. HTML report says 36. HTML is authoritative (matches
DMST display). This is a confirmed push script / firmware eaLen calculation bug on fw 6.1.16_sr4.

**Conclusion**: DMCC `GET SYMBOL.RESULT` is NOT an alternative data path. It returns identical
data to the push event XML. No additional fields are obtainable via this command.
The HTML scraping path is the only resolution for all four permanently unresolvable fields.

---

## 17. DMST HTML report field map — confirmed 2026-05-25

**Source**: `2026-05-24_23-03-58-752_1779678267324.html`
**Device**: DM475-63530E-PIPS-Verif-Lab, fw 6.1.16_sr4, QR GUID f0cffb39, Grade A, IMAGE.LOAD

### HTML structure overview

The report is a **single minified HTML line** with two main data tables:

1. **Header table** (cells 0–30): Multi-column layout. Labels and values in SEPARATE rows.
   Do NOT use consecutive `<td>` pairing here. Overall grade extracted via "D.D (L)" regex.

2. **Simple characteristics table** (cells 31–60): Clean `<td>Label</td><td>Value</td>` pairs.
   All four primary target fields live here.

3. **Grade parameters table** (cells 61+): 6-cell rows per parameter:
   `[label][secondary][pct%][numeric][letter][PASS/FAIL]`

### DateTime corruption

The in-page header shows `"Wed 31-Dec-1970 07:00:00"` — Unix epoch, corrupt.
**Always parse DateTime from the filename prefix**: `yyyy-MM-dd_HH-mm-ss-mmm_<random>.html`.

### Complete field extraction map (simple characteristics table)

| Cell index | Label | Confirmed value | Notes |
|---|---|---|---|
| [31/32] | "QR Size" | "29x29" | MatrixSize cross-validate |
| [33/34] | "Horizontal BWG" | "-3%" | strip "%" for decimal |
| [35/36] | "Vertical BWG" | "-4%" | strip "%" for decimal |
| [37/38] | "Encoded characters" | "36" | **HTML authoritative; push XML says 39 (WRONG)** |
| [39/40] | "Total Codewords" | "70" | cross-validate |
| [41/42] | "Data Codewords" | "44" | **empty in push XML; HTML authoritative** |
| [43/44] | "Error Correction Budget" | "26" | **empty in push XML; HTML authoritative** |
| [45/46] | "Errors Corrected" | "0" | cross-validate |
| [47/48] | "Error Capacity Used" | "0" | cross-validate |
| [49/50] | **"Error Correction Level"** | **"M"** | ★ PRIMARY TARGET |
| [51/52] | **"Data Mask Pattern"** | **"2"** | ★ PRIMARY TARGET |
| [53/54] | **"Image"** | **"Black on white"** | ★ PRIMARY TARGET (ImagePolarity) |
| [55/56] | "Nominal X Dim" | "12.6 mil" | cross-validate vs NominalXDim_2D |
| [57/58] | "Pixels per Module" | "9.75" | informational |
| [59/60] | **"ECI"** | **"000003"** | ★ PRIMARY TARGET |

### Grade parameters table row structure

| Row label | Cell+1 (secondary) | Cell+2 (pct) | Cell+3 (numeric) | Cell+4 (letter) |
|---|---|---|---|---|
| "1. Unused Error Correction (UEC)" | "" | "100.0%" | "4.0" | "A" |
| "2. Symbol Contrast (SC)" | "Rl/Rd (87/6)" | "nan%" | "4.0" | "A" |
| "3a. Modulation (MOD)" | "" | "" | "4.0" | "A" |
| "3b. Reflectance Margin (RM)" | "" | "" | "4.0" | "A" |
| "4. Axial Nonuniformity (ANU)" | "" | "0.8%" | "4.0" | "A" |
| "5. Grid Nonuniformity (GNU)" | "" | "0.0%" | "4.0" | "A" |
| "6. Fixed Pattern Damage (FPD)" | "" | "" | "4.0" | "A" |

**"nan%"** on SC: IMAGE.LOAD scans have no live illumination → SC% not computed → `null` in DmstHtmlReport.SCPercent.

### ImagePolarity label

The HTML label for image polarity is **"Image"** (not "ImagePolarity" or "Polarity").
Value: "Black on white" → `ImagePolarity.BlackOnWhite`; "White on black" → `ImagePolarity.WhiteOnBlack`.
No other values observed; "Inverted" is NOT used.

### No external library required

Parser uses `Regex.Matches` on `<td>` elements only. The single-line minified HTML makes
regex extraction reliable. HtmlAgilityPack is NOT required and has not been added.

### DataCodewords and ErrorCorrectionBudget — HTML eliminates C# table lookup need

Prior plan (v1.30 bug list #5/#6) was to derive DataCodewords and ErrorCorrectionBudget from
C# lookup tables (ECC200 table for DM, QR size table for QR) because these were empty in push XML.
**HTML report provides both values directly and authoritatively. C# table lookup is no longer
needed for any HTML-correlated scan.** The table lookup remains as a fallback for scans where
DMST is not running or the .html extension is not configured.

---

## 10. DMST-native wire protocol — Wireshark capture analysis (2026-05-25)

**Source**: `vtccp/architecture/gui-reference/wireshark-dmst-full-capture.txt` (7200 lines)
**Full analysis**: `vtccp/architecture/wireshark-protocol-analysis.md`
**Capture date**: 2026-05-25 (DM475-63530E-PIPS-Verif-Lab, fw 6.1.16_sr4)
**Confirmed**: Live scan trigger_index=55 (good_reads: 27→28), DM 16×36 GS1 label, Grade D

### 10.1 HTTP result-push protocol endpoints

The DMST-native result delivery uses HTTP pub/sub over a dedicated TCP channel (NOT port 44444):

| Endpoint | Direction | Size | Frequency | Content |
|---|---|---|---|---|
| `GET /events?enable` | DMST → device | — | Once (subscribe) | Subscription handshake |
| `HTTP/1.1 204 No Content` | device → DMST | 0 | Response to GET | Subscription confirmed |
| `PUT /status.xml` | device → DMST | ~4,625 B | ~1/second | Telemetry (read stats, timing) |
| `PUT /vs.cfg` | device → DMST | 288–400 B | Irregular | Config sync (AES-encrypted, unreadable) |
| `PUT /codes.xml` | device → DMST | ~9,415 B (monitor) / ~202,249 B (verify) | Per scan | Full result XML including push XML + trucheck data |
| `PUT /pcm_report.html` | device → DMST | 131–202 KB | Per verification scan | Full HTML report (sent BEFORE codes.xml) |

Device User-Agent: `DM475/6.1.16 (DeviceID=50)`

**VTCCP note**: `DmstListener` uses the DataMan Network Client mechanism (raw XML, not HTTP).
The HTTP endpoints above are used by DMST only. The HTML files that `DmstHtmlScraper` reads
are the disk-saved versions of the `PUT /pcm_report.html` body — identical content.

### 10.2 codes.xml — `origin` field discriminator

The `<result>` root element carries an `origin` attribute:

| Value | Meaning | codes.xml size | trucheck data complete? | pcm_report.html sent? |
|---|---|---|---|---|
| `"monitor"` | Continuous background monitoring scan | ~9,415 B | No (only CalibrationDate + OpticalVariant + NO DECODE) | No |
| `"common"` | Full triggered TruCheck verification | ~202,249 B (includes JPEG) | Yes (all tables) | Yes (sent first) |

### 10.3 trucheck_verificaiton_result XML — confirmed field inventory

**Tag name**: `<trucheck_verificaiton_result>` — "verificaiton" is a firmware misspelling (ai transposed).

**Top-level fields** (present on all scans):
- `<OpticalVariant>DM475V</OpticalVariant>` — exact model string (more precise than DEVICE.TYPE which may return family "DM470")
- `<CalibrationDate>5/20/2026 1:14:58 AM</CalibrationDate>`
- `<CalibrationState>0</CalibrationState>` — meaning of 0 not yet confirmed

**SymbolData timing fields** (present on `origin="common"` only):
- `<VerificaitonTime>183</VerificaitonTime>` — "Verificaiton" misspelled (183 ms total)
- `<PreDecodeTime>0</PreDecodeTime>`
- `<BlurTime>64</BlurTime>`
- `<ThreshTime>13</ThreshTime>`
- `<StickTime>0</StickTime>`
- `<LineSearchTime>13</LineSearchTime>`
- `<CanidateEvaluationTime>30</CanidateEvaluationTime>` — "Canidate" misspelled
- `<PostDecodeTime>51</PostDecodeTime>`
- `<ResultTime>0</ResultTime>`

**GradeInfo (full ISO formal notation)**:
```
<FormalGrade>1.0/16/660/45Q</FormalGrade>
```
Format: `numericGrade / aperture / wavelength / lighting`. Push XML `<FormalGrade>` gives `1/D`
(abbreviated form). The trucheck XML provides the complete ISO 15415 formal notation.

**Quality parameter table numbers** confirmed from this capture (DM 16×36):

1=UEC, 2=SC, 3a=MOD, 3b=RM, 4=ANU, 5=GNU, 6=FPD, 7=LLS, 8=BLS, 9=LQZ, 10=BQZ,
11a=ULQZ, 11b=URQZ, 12a=RUQZ, 12b=RLQZ, 13a=LQTTR, 13b=RQTTR, 14a=LQRTR, 14b=RQRTR,
15a=LQTCT, 15b=RQTCT, 16a=LQRCT, 16b=RQRCT, 17=AG (Average Grade), 18=DECODE

**General Characteristics section** — confirmed values (DM 16×36 GS1 live scan):

| Parameter | trucheck XML value | Push XML value | Status |
|---|---|---|---|
| Matrix Size | `16x36 (Data: 14x34)` | `16x36` | trucheck adds inner data region size |
| Encoded characters | **`38`** | `33` (wrong) | ★ **trucheck value is correct** |
| Total Codewords | `56` | `56` | Match |
| Data Codewords | **`32`** | `` (empty) | ★ **trucheck resolves bug #5** |
| Error Correction Budget | **`24`** | `` (empty) | ★ **trucheck resolves bug #6** |
| Errors Corrected | `7` | `7` | Match |
| Error Capacity Used | `14` | `14` | Match |
| Error Correction Type | `ECC 200` (space) | `ECC200` (no space) | Minor format difference |
| **Image** | **`Black on white`** | `` (empty) | ★ **ImagePolarity resolved** |
| Nominal X Dim | `20.3 mil` | `20.3 mil` | Match (per-scan measurement) |
| Pixels per Module | `15.96` | — (absent) | New field — not in push XML |
| Contrast Uniformity | `74 at module(12,17)` | `74` | trucheck adds location |
| MRD | `67% (73% - 6%)` | `67` | trucheck adds expanded form |

**ECLevel, DataMaskPattern, ECI**: NOT present in trucheck XML. FIB section has grade/numericGrade
only (consistent with v1.33 probe campaign findings). These three fields remain permanently
unresolvable from device protocol on fw 6.1.16_sr4.

### 10.4 New push XML fields confirmed from this capture

These fields appeared in the decoded full_string but were not previously inventoried:

| Field | Value (DM 16×36 live) | Notes |
|---|---|---|
| `<TQZGrade>` | `X` | Top Quiet Zone — not applicable on this DM orientation |
| `<RQZGrade>` | `X` | Right Quiet Zone — not applicable |
| `<ULQZGrade>` through `<LRQZGrade>` | `` (empty) | 6 additional QZ grades — likely DMV-8072V fields |
| `<HClockTrackGrade>` | `` | Horizontal clock track — empty on standard DM |
| `<VClockTrackGrade>` | `` | Vertical clock track — empty on standard DM |
| `<ULQTTRGrade>` through `<URQRTRGrade>` | `` | Transition ratio grades — empty on standard DM |

All new fields are empty on the standard 16×36 DM symbol. Likely populated for DMV-8072V
or DM variants with non-standard finder pattern topology.

### 10.5 NominalXDim — confirmed per-scan values

| Scan | Symbology | NominalXDim |
|---|---|---|
| #12 (DM 16×36 live, v1.32) | Data Matrix | `20.3 mil` |
| #11/#13/#15 (QR 29×29 IMAGE.LOAD) | QR | `12.6 mil` |
| This Wireshark scan (DM 16×36 live) | Data Matrix | `20.3 mil` |

NominalXDim is a per-scan measurement derived from the optical image — not a device constant.
It varies with symbol size and distance from camera.

