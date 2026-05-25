# Parser Alignment Gap — v1.25 Push Script vs DmstResultParser

**Date**: 2026-05-18 (updated 2026-05-24)
**Script version compared**: v1.32 (device-confirmed 2026-05-24)
**Parser version compared**: post-v1.32 QR wiring (VerificationXmlMap + DmstResultParser updated 2026-05-24)
**Reference XML**: `TestHarness/Fixtures/dmst_qr_grade_a_v132.xml` (device-confirmed QR Grade A)
**Gap status source**: `vtccp/architecture/firmware-confirmed-facts.md` (authoritative, living doc)

---

## Gap table legend

| Gap type | Meaning |
|---|---|
| `WIRED` | Parser reads element; VR field populated |
| `ABSENT-EMPTY` | Script emits element but always empty on this firmware; parser ignores correctly |
| `RESOLVED-B4` | Not wired before B4; wired as part of this task |
| `PENDING-PROBE` | Element present but correct parse path blocked on v1.25 DebugSymbols0/DebugPrintGrowth result |
| `DESIGN-INTENT` | Element present in script; intentionally not mapped to VR (probe/debug only) |
| `LEGACY-KEEP` | Parser reads element removed from v1.25 script; kept for backward compat with older XMLs |

---

## Block 1 — Identity / timing

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<SymbologyType>` (fw6.x) / `<SymbologyName>` (legacy) | all | `Symbology` | WIRED | map default |
| `<DecodedData>` | all | `DecodedData` | WIRED | v1.25: ECI prefix stripped for QR modifier 2/4/6 before emit |
| `<SymbologyId>` | v1.24 | `SymbologyId` | RESOLVED-B4 | AIM ID e.g. `]d1`, `]Q2` |
| `<SymbolQuality>` | v1.24 | `SymbolQuality` | RESOLVED-B4 | Decoder confidence 0–100 |
| `<SymbolAngle>` | v1.24 | `SymbolAngle` | RESOLVED-B4 | Rotation degrees; NOT reliable for LoadedImage detection |
| `<ModuleSizePx>` | v1.24 | `ModuleSizePx` | RESOLVED-B4 | r.symbology.moduleSize (pixels); replaces PixelsPerModule |
| `<Source>` | v1.24 | `DeviceName` fallback | RESOLVED-B4 | r.source device name; used when deviceContext.DeviceName null |
| `<CalibrationDate>` | v1.24 | `CalibrationDate` | WIRED (pre-B4) | Already parsed at lines 232–236 |
| `<FieldCalibrated>` | v1.24 | `FieldCalibrated` | RESOLVED-B4 | rp.status3D.fieldCalibrated; bool via Bool() helper |
| `<FactoryCalibrated>` | v1.24 | `FactoryCalibrated` | RESOLVED-B4 | rp.status3D.factoryCalibrated; bool via Bool() helper |
| `<MinPassGrade>` | v1.24 | `MinPassGrade` | RESOLVED-B4 | Letter grade e.g. "C", "NA" (NA = none configured) |
| `<MinPassRaw>` | v1.24 | `MinPassRaw` | RESOLVED-B4 | Numeric threshold (empty when NA) |
| `<ApplicationStandard>` | v1.24 | `ApplicationStandard` | RESOLVED-B4 | e.g. "GS1", "ISO 15434" |
| `<ApplicationPass>` | v1.24 | `ApplicationPass` | RESOLVED-B4 | Full string: "Pass" / "Fail (Quality)" / "Fail (X Dimension out of Range)" |
| `<ApplicationPassReason>` | v1.25 | `ApplicationPassReason` | RESOLVED-B4 | Reason suffix only: "Quality", "", "X Dimension out of Range" |
| `<OpticsSource>` | v1.25 | `OpticsSource` | RESOLVED-B4 | "LiveScan" \| "LoadedImage" — script derives from CU==-1 && MRD==-1 |
| `<JpegImageBase64>` | v1.25 | `JpegImageBase64` | RESOLVED-B4 | Full base64 JPEG; unblocks B6 (ImagesSheetWriter) |
| `<CustomNote>` | v1.24 | `CustomNote` fallback | RESOLVED-B4 | q.customNote device calibration note; sessionContext overrides |

---

## Block 2 — Grading summary

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<FormalGrade>` | all | `FormalGrade` | WIRED | |
| `<OverallGrade>` | all | `OverallGrade.Letter` | WIRED | |
| `<OverallGradeNumeric>` | all | `OverallGrade.Numeric` | WIRED | |
| `<GradingStandard>` | v1.19 | `Standard` fallback | WIRED (partial) | Parser primary = isoGradeInfo/Standard; `<Standard>` (line 845) also mapped via map.Standard default. `<GradingStandard>` (line 825) is a separate element not separately consumed — both carry the same grading standard string. Acceptable. |
| `<ApplicationStandard>` | v1.24 | `ApplicationStandard` | RESOLVED-B4 | See Block 1 |
| `<ApplicationPass>` | v1.24 | `ApplicationPass` | RESOLVED-B4 | See Block 1 |
| `<ApplicationPassReason>` | v1.25 | `ApplicationPassReason` | RESOLVED-B4 | See Block 1 |

---

## Block 3 — Verification conditions

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<ApertureRef>` | all | `Aperture` | WIRED | |
| `<Wavelength>` | all | `Wavelength` | WIRED | |
| `<Lighting>` | all | `Lighting` | WIRED | |
| `<Standard>` | all | `Standard` | WIRED | Also caught by isoGradeInfo path |

---

## Block 4 — 2D ISO 15415 quality parameters

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<UECPercent>` | all | `UEC_Percent` | WIRED | |
| `<UECGrade>` | all | `UEC_Grade` | WIRED | |
| `<SCPercent>` | all | `SC_Percent` | WIRED | |
| `<SCRlRd>` | all | `SC_RlRd` | WIRED | Rl/Rd formatted as "89/4" |
| `<SCGrade>` | all | `SC_Grade` | WIRED | |
| `<MODGrade>` | all | `MOD_Grade` | WIRED | |
| `<RMGrade>` | all | `RM_Grade` | WIRED | |
| `<ANUPercent>` | all | `ANU_Percent` | WIRED | **Bug #9 fixed v1.32**: v1.32 script reads `.percent` sub-key of `q.metrics.axialNonUniformity`. Raw value is already in % scale (0.8392 = 0.84%). Do NOT apply ÷100. Confirmed device value 2026-05-24. |
| `<ANUGrade>` | all | `ANU_Grade` | WIRED | |
| `<GNUPercent>` | all | `GNU_Percent` | WIRED | |
| `<GNUGrade>` | all | `GNU_Grade` | WIRED | |
| `<FPDValue>` | v1.19 | `FPD_Value` | RESOLVED-B4 | q.fixedPatternDamage.raw |
| `<FPDGrade>` | all | `FPD_Grade` | WIRED | |
| `<DecodeGrade>` | all | `DECODE_Grade` | WIRED | |
| `<AGValue>` | all | `AG_Value` | WIRED | Can be negative for loaded images; `decimal?` handles correctly |
| `<AGGrade>` | all | `AG_Grade` | WIRED | |
| `<DDGrade>` | v1.19 | `DD_Grade` | RESOLVED-B4 | q.distributedDamageGrade |
| `<AverageGrade>` | v1.19 | `AverageGrade` | RESOLVED-B4 | q.averageGrade letter |
| `<AverageGradeNumeric>` | v1.19 | `AverageGradeNumeric` | RESOLVED-B4 | q.averageGrade.raw |
| `<MinReflectance>` | v1.19 | `MinReflectance` | RESOLVED-B4 | Suppressed when grade=F AND raw=0 (firmware sentinel); non-zero values represent real minR measurements |

**Known discrepancy — AG vs AverageGrade**: v1.14 confirmed the device's push `<AGValue>/<AGGrade>` (parameter 17) consistently gives a HIGHER letter grade than DMST PDF report "Average Grade". Hypothesis: DMST PDF uses a different parameter subset for average. The push value (`q.printGrowth` = Print Growth) IS ISO 15415 parameter 17 (Print Growth/BWG). The DMST PDF "Average Grade" appears to derive from parameters 1–16 mean. Document this discrepancy in report output; do not attempt to reconcile — VTCCP emits the device's push values with notes.

---

## Block 5 — 2D matrix general characteristics

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<MatrixSize>` | all | `MatrixSize` | WIRED | v1.25 fixed QR branch (QZ offset 8, not 2) |
| `<HorizontalBWG>` | all | `HorizontalBWG` | WIRED | |
| `<VerticalBWG>` | all | `VerticalBWG` | WIRED | |
| `<EncodedCharacters>` | all | `EncodedCharacters` | WIRED | Note: count convention differs from DMST PDF (different counting — not a bug) |
| `<TotalCodewords>` | all | `TotalCodewords` | WIRED | |
| `<DataCodewords>` | all | `DataCodewords` | WIRED (always empty) | ABSENT-EMPTY: firmware does not split CW count; field stays null. `ErrorCapacityUsed = ErrorsCorrected × 2` can be derived if needed |
| `<ErrorCorrectionBudget>` | all | `ErrorCorrectionBudget` | WIRED (always empty) | ABSENT-EMPTY: firmware does not expose |
| `<ErrorsCorrected>` | all | `ErrorsCorrected` | WIRED | Counted from codewordArray.isCorrected |
| `<ErrorCapacityUsed>` | all | `ErrorCapacityUsed` | WIRED (always empty) | ABSENT-EMPTY: derivable as ErrorsCorrected × 2 |
| `<ErrorCorrectionType>` | all | `ErrorCorrectionType` | WIRED | DM→"ECC200". QR→ECLevel is **DEFINITIVELY DEAD** on fw 6.1.16_sr4: `q.symbols=null`, `r.symbology` has no ecLevel key, AIM modifier=ECI presence not ECLevel, `r.image` = hardware only. Field emits "QR" for QR scans — correct as a placeholder; ECLevel will never be resolvable from push XML on this firmware. |
| `<NominalXDim>` | all | `NominalXDim_2D` | WIRED | Units stripped. **v1.32 fix**: push XML emits `"12.6 mil"` (with units); parser now splits on space for both push and CharData paths. Confirmed 2026-05-24. |
| `<PixelsPerModule>` | ≤v1.24 | `PixelsPerModule` | LEGACY-KEEP | Removed from v1.25 script (always empty — r.trucheck doesn't expose it). Parser still reads it for backward compat with v1.24 and earlier XMLs. Use `ModuleSizePx` (r.symbology.moduleSize) as the live value instead. |
| `<ImagePolarity>` | all | `ImagePolarity` | WIRED | Parsed to enum |
| `<ContrastUniformity>` | all | `ContrastUniformity` | WIRED | −1 on loaded images (one component of OpticsSource detection) |
| `<MRD>` | all | `MRD` | WIRED | −1 on loaded images |
| `<ContrastUniformityRow>` | v1.19 | `ContrastUniformityRow` | RESOLVED-B4 | Row index of worst module (gnProp) |
| `<ContrastUniformityCol>` | v1.19 | `ContrastUniformityCol` | RESOLVED-B4 | Col index of worst module (gnProp) |
| `<OpticsSource>` | v1.25 | `OpticsSource` | RESOLVED-B4 | See Block 1 |
| `<JpegImageBase64>` | v1.25 | `JpegImageBase64` | RESOLVED-B4 | See Block 1 |

---

## Block 6 — 2D L-sides and quiet zones

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<LLSGrade>` | all | `LLS_Grade` | WIRED | |
| `<BLSGrade>` | all | `BLS_Grade` | WIRED | |
| `<LQZGrade>` | all | `LQZ_Grade` | WIRED | |
| `<BQZGrade>` | all | `BQZ_Grade` | WIRED | |
| `<TQZGrade>` | all | `TQZ_Grade` | WIRED | |
| `<RQZGrade>` | all | `RQZ_Grade` | WIRED | |
| `<TTRPercent>` | all | `TTR_Percent` | WIRED | |
| `<TTRGrade>` | all | `TTR_Grade` | WIRED | |
| `<RTRPercent>` | all | `RTR_Percent` | WIRED | |
| `<RTRGrade>` | all | `RTR_Grade` | WIRED | |
| `<TCTGrade>` | all | `TCT_Grade` | WIRED | |
| `<RCTGrade>` | all | `RCT_Grade` | WIRED | |

---

## Block 7 — 2D quadrant parameters (≥32×32)

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<ULQZGrade>` – `<LRQRCTGrade>` (24 elements) | v1.22 | All `*Grade` fields | WIRED (always empty) | Script emits `""` since v1.22 rollback. q.upperLeftPattern etc. are inert placeholder objects in this firmware's JS scope. No false data emitted. Will re-probe in a future version when DebugSymbols0 reveals if per-region data is accessible via symbols[0]. |

---

## Block 8 — 1D ISO 15416 summary

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<SymbolAnsiGrade>` | all | `SymbolAnsiGrade` | WIRED | |
| `<AvgEdge>` – `<AvgMinQZ>` (11 elements) | all | `Avg_*` fields | WIRED | All empty on DM475V (2D verifier); populated only on 1D scans |
| `<BWGPercent>` | v1.24 | `BWG_Percent` | WIRED | Still empty on device (v1.24 confirmed); m.printGrowth probe in v1.25 will diagnose |
| `<Magnification>` | all | `Magnification` | WIRED | |
| `<Ratio>` | all | `Ratio` | WIRED | Code 39 only |
| `<NominalXDim1D>` | all | `NominalXDim_1D` | WIRED | |
| `<ScanResults>` block | all | `ScanResults` | WIRED | 1D per-scan results |

---

## Block 9 — QR Code-specific (updated 2026-05-24, all probes answered)

| XML element | Script version | VR field | Gap type | Notes |
|---|---|---|---|---|
| `<ULPGrade>` | v1.29+ | `QR_ULP_Grade` | WIRED (v1.32) | Upper-Left Finder Pattern. Present when `SymbologyName="QR"`. Grade A scan: A. |
| `<URPGrade>` | v1.29+ | `QR_URP_Grade` | WIRED (v1.32) | Upper-Right Finder Pattern. |
| `<LLPGrade>` | v1.29+ | `QR_LLP_Grade` | WIRED (v1.32) | Lower-Left Finder Pattern. Emits "F" on defocus (scan #4 confirmed). |
| `<HCTGrade>` | v1.29+ | `QR_HCT_Grade` | WIRED (v1.32) | Horizontal Clock Track. |
| `<VCTGrade>` | v1.29+ | `QR_VCT_Grade` | WIRED (v1.32) | Vertical Clock Track. |
| `<ALPGrade>` | v1.29+ | `QR_ALP_Grade` | WIRED (v1.32) | Alignment Pattern (v2+). |
| `<VIBGrade>` | v1.29+ | `QR_VIB_Grade` | WIRED (v1.32) | Version Information Blocks. Emits `"-"` for v1/v3 QR (no VIB). Emits `"F"` on total fail. |
| `<FIBGrade>` | v1.29+ | `QR_FIB_Grade` | WIRED (v1.32) | Format Information Blocks. |
| `QR_Version` | — | `QR_Version` | DERIVABLE | Not in push XML. Derive from `MatrixSize` (e.g. 29×29 → v3). Not yet implemented. |
| `QR_ECLevel` | — | `QR_ECLevel` | DEAD | All firmware paths closed on fw 6.1.16_sr4. See §3 of `firmware-confirmed-facts.md`. |
| `QR_MaskPattern` | — | `QR_MaskPattern` | DEAD | Not present in `r.symbology` (9 keys confirmed). See `firmware-confirmed-facts.md` §1. |
| `SymbologyName="QR"` | v1.29+ | `SymbologyFamily` | WIRED (v1.32) | Push emits `"QR"` not `"QR Code"`. `"QR"` prefix added to `ClassifySymbology` map. |

**Note — DebugSymbols0 probe (v1.25)**: `q.symbols = null` confirmed by `DebugSymbols` probe
(v1.29, 2026-05-19). The symbols path is dead on fw 6.1.16_sr4. QR grade params are accessed
directly from the push XML elements (`<ULPGrade>` etc.), not via `q.symbols[0]`.

---

## Summary: RESOLVED-B4 count

23 fields newly wired in this task:

**Block 1 (9)**: SymbologyId, SymbolQuality, SymbolAngle, ModuleSizePx, DeviceName fallback (Source), FieldCalibrated, FactoryCalibrated, MinPassGrade, MinPassRaw

**Block 1 continued (6)**: ApplicationStandard, ApplicationPass, ApplicationPassReason, OpticsSource, JpegImageBase64, CustomNote fallback

**Block 4 (5)**: FPD_Value, DDGrade, AverageGrade, AverageGradeNumeric, MinReflectance

**Block 5 (3)**: ContrastUniformityRow, ContrastUniformityCol, (OpticsSource/JpegBase64 — counted above)

---

## Remaining gaps (updated 2026-05-24)

| Gap | Status | Notes |
|---|---|---|
| `BWGPercent` always empty | OPEN | `m.printGrowth` probe not yet confirmed on device. Still empty as of v1.32. |
| `<ErrorCorrectionType>` QR — emits "QR" placeholder | CLOSED (no fix possible) | ECLevel definitively dead on fw 6.1.16_sr4. "QR" is the correct permanent value for QR scans. |
| QR grade params (ULP/URP/LLP/HCT/VCT/ALP/VIB/FIB) | CLOSED — WIRED (v1.32) | All 8 wired in `VerificationXmlMap` + `DmstResultParser`. Confirmed from `dmst_qr_grade_a_v132.xml`. |
| `QR_Version` | OPEN — DERIVABLE | Not in push XML. Derive from `MatrixSize` (29×29 → v3). C# lookup table not yet written. |
| `QR_MaskPattern` | CLOSED (dead) | Not in `r.symbology` (9 keys confirmed). Unresolvable on fw 6.1.16_sr4. |
| `QR_ECLevel` | CLOSED (dead) | All 5 paths exhausted. See `firmware-confirmed-facts.md` §3. |
| Per-region DM quadrant grades (ULQZ/URQZ/RUQZ/RLQZ) | OPEN | Script emits `""` since v1.22 rollback. `q.symbols=null` means `q.symbols[0]` path is also dead. Revisit if future firmware exposes per-region data. |
| `ImagePolarity` | CLOSED (dead) | `q.general` dead paths confirmed; `r.image` = hardware metadata only. Unresolvable. |
| `ECI` (value, e.g. 000003) | CLOSED (dead) | All push paths empty. DMST shows it; push XML does not expose it on fw 6.1.16_sr4. |
| QR pattern grades in ExcelWriter sheet mapper | OPEN — TECH DEBT | Parsed into `VerificationRecord` (v1.32) but not yet written to Excel output columns. |
| ModulationValues / CodewordValues | COMPLETE (B7) | `ModValuesSheetWriter` + `CwValuesSheetWriter` + VerificationRecord + ExcelWriter all wired. |
