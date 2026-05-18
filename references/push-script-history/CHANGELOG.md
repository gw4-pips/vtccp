# DMST Push Script — Changelog

Every shipped version of `DmstPushScript_v1.js`. The "v1" suffix in the
filename refers to the push-script *generation*, not the version number; the
real version is in the `<PushScriptDiag>vX.YY ...</PushScriptDiag>` tag emitted
in every result XML.

Source of truth for the build: `vtccp/DeviceInterface/Dmst/DmstPushScript_v1.js`.
Archived copies in this directory.

Install ritual (every version):
1. DMST → Format Data → Scripting tab
2. Open Script → paste in the new version
3. Save → Write Settings to verifier
4. Trigger a scan → confirm `<PushScriptDiag>vX.YY ...</PushScriptDiag>` matches expected

---

## Versions

### v1.26 — 2026-05-18

Filed: `v1.26.js` (also `dist/DmstPushScript_v1.26.txt`).
Status: **authored, syntax-verified, awaiting device install**.
Confirm by: `<PushScriptDiag>v1.26 q=r.trucheck m=found</PushScriptDiag>`.

**Probe release.** Two new probes driven by v1.25 live DM scan findings
(catalog: `samples/live-scans/v1.25-2026-05-18-DM-Live-GS1Format06-catalog.md`).

#### v1.26 changes

**New probes:**
- `<DebugTrucheckKeys>` — ES3-safe `for-in` enumeration of all top-level keys
  of `q` (= `r.trucheck`). Target: find where ULQZ/URQZ/RUQZ/RLQZ and TTR/CTR
  sub-grades are accessible — X in push XML, A in DMST PDF. v1.25 confirmed
  these grades are computed by the device; they must live in an unexplored
  `q` sub-path. This probe reveals the full key set so v1.27 can target
  the correct path. Also useful for QR: any QR-specific sub-objects will
  appear here on a live QR scan.
- `<DebugMinRef>` — Probes `r.metrics.extremeReflectance` and `r.metrics.reflectMin`
  shapes. v1.25 `DebugMetricsKeys` confirmed both keys exist on `r.metrics`.
  `<MinReflectance>` is currently empty in push XML. Per v1.19 history note,
  some metrics objects are primitives (bare numbers), not `{raw,grade}` — probe
  checks both cases and reports type + value so v1.27 can wire the correct path.

**Retained probes:**
- `DebugModSize`, `DebugECCount` (formula sanity checks, always useful)
- `DebugMetricsKeys`, `DebugRSiblings` (regression baselines)
- `DebugSymbols0` — confirmed null for DM live scans (v1.25); only useful on
  a live QR scan. Retained for QR sessions.

**Dropped probes (answered by v1.25 live DM scan):**
- `DebugPrintGrowth` — answered: `raw=-1` sentinel; BWG comes from `q.general`
  (`HorizontalBWG`/`VerticalBWG`), not `r.metrics.printGrowth`. Field
  `BWGPercent` (singular, legacy) may be declared dead.
- `DebugImageShape` — answered: `id=0;index=0;FoV=[obj];RoI=[obj];
  exposureTime=32;gain=1.00;autoExposure=true;illEnabled=true;...`
- `DebugGS1` — needs `Data Format Check = GS1` device setting to populate.
  Defer until user sets up GS1 application standard and scans a GS1 symbol.
- `DebugDodUid` — needs DoD UID symbol + matching application standard setting.
- `DebugBarcodeAsgn` — answered: `result=-1` (barcode assignment not configured).

#### v1.26 expected install findings

| Element | Expected result | Action on result |
|---|---|---|
| `<PushScriptDiag>` | `v1.26 q=r.trucheck m=found` | Confirms install |
| `<DebugTrucheckKeys>` | Key list of q properties | Look for ULQZ/URQZ/TTR/CTR sub-object keys; drive v1.27 wires |
| `<DebugMinRef>` | `extR:number=... reflMin:number=...` or object shape | Determines MinReflectance wire path for v1.27 |
| `<DebugSymbols0>` | `(q.symbols[0] null)` on DM; actual data on QR | Run live QR scan for QR grade param discovery |

---

### v1.25 — 2026-05-18

Filed: `v1.25.js` (also `dist/DmstPushScript_v1.25.txt`).
Status: **authored, syntax-verified, awaiting device install**.
Confirm by: `<PushScriptDiag>v1.25 q=r.trucheck m=found</PushScriptDiag>`.

**Confirmed-wire release + QR foundation.** Six fields promoted from probe/deferred to first-class; one structural fix (MatrixSize QR branch); two new probe targets.

#### v1.25 changes

**New first-class elements:**
- `<JpegImageBase64>` — full `r.trucheck.jpegImage` base64 JPEG string. v1.24 DebugJpegProbe confirmed safe: ~9,912 chars live DM / ~22,260 chars loaded QR. Unblocks B6 (ImagesSheetWriter).
- `<OpticsSource>` — `LoadedImage` when `ContrastUniformity===-1 AND MRD===-1`; `LiveScan` otherwise. Confirmed across two loaded-image captures. SymbolAngle is NOT a reliable discriminator.
- `<ApplicationPassReason>` — reason suffix parsed from `ApplicationPass` (split on ` (`). `"Pass"` → `""`. `"Fail (Quality)"` → `"Quality"`. `"Fail (X Dimension out of Range)"` → `"X Dimension out of Range"`. Format failures will add more tokens when Data Format Check ≠ None.

**Fixes:**
- `<DecodedData>` ECI prefix stripped for QR codes with AIM modifier 2/4/6 (`\x00\x00\x26` = ECI 26 = UTF-8). Modifier 1/3/5 → no strip.
- `<MatrixSize>` QR branch: QZ offset = 8 (4 modules per side × 2). QR v3: sqrt(1369)=37, 37−8=29 ✓. v1.24 emitted 35 (wrong; used DM offset of 2).
- `<ErrorCorrectionType>` symbology-aware: DM → `"ECC200"`; QR → `ovProp("ecLevel")` or `q.ecLevel` fallback → `"QR"` until DebugSymbols0 reveals exact path.
- `<PixelsPerModule>` removed — always empty; `<ModuleSizePx>` (from `r.symbology.moduleSize`) is the sole confirmed source.

**New probes:**
- `<DebugSymbols0>` — full enum of `q.symbols[0]`. Targets: QR grade params (`ulpGrade`/`urpGrade`/`llpGrade`/`hctGrade`/`vctGrade`/`alpGrade`/`vibGrade`/`fibGrade`), `ecLevel`, `maskPattern`, `rows`, `cols`; DM per-region QZ grades (`ulqzGrade`/`urqzGrade`/`ruqzGrade`/`rlqzGrade`), `dataCodewords`, `ecCodewords`, `minLightReflectance`, `maxDarkReflectance`.
- `<DebugPrintGrowth>` — `m.printGrowth` shape diagnostic. Expected `{raw, grade}`; if `raw===-1` → sentinel (empty is correct); if null → accessor wrong; if bare number → use directly.

**Fixed probe:**
- `<DebugImageShape>` — wrapped in try/catch; v1.24 emitted nothing (silent error). Will now report `(error: ...)` if `r.image` is inaccessible.

**Dropped probes:**
- `DebugANUCase` — resolved: lowercase `axialNonuniformity` confirmed correct.
- `DebugJpegProbe` — superseded by `<JpegImageBase64>`.
- `DebugReaderProps` — `status3D.{fieldCalibrated,factoryCalibrated}` confirmed accessible in v1.24.

#### v1.25 expected install findings

| Element | Expected result | Action on result |
|---|---|---|
| `<PushScriptDiag>` | `v1.25 q=r.trucheck m=found` | Confirms install |
| `<JpegImageBase64>` | Long base64 string starting `/9j/` | Wire parser (B4), unblocks B6 |
| `<OpticsSource>` | `LiveScan` on device scan | Confirmed; `LoadedImage` on next image-load scan |
| `<ApplicationPassReason>` | `Quality` (or `""` if Pass) | Wire parser (B4) |
| `<MatrixSize>` on QR | `29x29` for v3 (not 35x35) | Fix confirmed |
| `<DebugSymbols0>` | Key list — look for `ulpGrade`, `ecLevel`, `maskPattern`, `dataCodewords` | Drive v1.26 scope |
| `<DebugPrintGrowth>` | Object shape of `m.printGrowth` | Fix BWGPercent in v1.26 |
| `<DebugImageShape>` | Key list or `(error: ...)` | Confirm r.image accessibility |

---

### v1.24 — 2026-05-18

Filed: `v1.24.js` (also `dist/DmstPushScript_v1.24.txt`).
Live-confirmed: `samples/live-scans/v1.24-2026-05-18-Probe-DataMatrix-GS1Format06.xml`.

**Probe release.** Status: **device-confirmed 2026-05-18** via `<PushScriptDiag>v1.24 q=r.trucheck m=found</PushScriptDiag>`. B1 closed.

#### v1.24 live findings (from confirmation scan)

Promotions that **worked** (populated correctly):
- `<SymbologyId>` = `]d1` ✓
- `<SymbolQuality>` = `41` ✓
- `<SymbolAngle>` = `1` ✓
- `<ModuleSizePx>` = `16.196012496948242` ✓
- `<CalibrationDate>` = `1/15/2026 3:04:14 PM` ✓
- `<FieldCalibrated>` = `false` ✓
- `<FactoryCalibrated>` = `false` ✓
- `<MinPassGrade>` = `NA` ✓ (`<MinPassRaw>` empty — sentinel for "no min-pass configured", expected)

Promotions that **failed** (still empty, need v1.25 fix):
- `<BWGPercent>` — still empty despite `mmPctAuto(printGrowth)` wire. The `r.metrics.printGrowth` accessor is either undefined on this build, or returns `{raw:-1}` which `mmPctAuto` collapses to empty. v1.25 must probe `r.metrics.printGrowth` shape explicitly.

Probe results (drive v1.25 scope):
- `<DebugJpegProbe>` = `type=string len=9912 preview=/9j/4AAQSkZJRgABAQ...` — **JPEG is a base64 string, 9912 chars ≈ 7.4 KB raw**. Well within Network Client buffer. **Full payload safe to emit in v1.25.**
- `<DebugANUCase>` = `lower=present(grade=D,raw=11.368...) upper=null` — **case mismatch resolved**: lowercase `axialNonuniformity` is correct; uppercase from comms guide doesn't exist on this build. Drop upper, keep lower first-class in v1.25.
- `<DebugBarcodeAsgn>` = `result=-1;stats=[obj]` — `r.barcodeAssignment` exists, returns -1 (unassigned), exposes `.stats` sub-object. Worth deep-probe in v1.25.
- `<DebugReaderProps>` = `name=...;trigger=[obj];stats=[obj];inputstr=;status3D=[obj]` — confirms `readerProperties` shape; `.status3D` is the path that yielded `FieldCalibrated`/`FactoryCalibrated`.
- `<DebugGS1>` = all 187 AIs `undefined` — **root cause now known**: Application Settings → Data Format Check was set to **None** during this scan, so the device never ran a GS1 check and `r.validation.gs1` is an unpopulated stub. This is a **device settings issue, not a JS API access-pattern problem**. v1.25 action: user must set Data Format Check = GS1 in Application Settings, scan a GS1-encoded symbol (GS1 Format 06 cal card is ideal), then enumerate `r.validation.gs1` keys — not guess property names.
- `<DebugDodUid>` = all 5 fields `undefined` — same situation: DodUid check requires a matching UID-formatted symbol AND the Application Standard set to a DoD-recognizing mode. Set Data Format Check and scan a MIL-STD-130 UID symbol before re-probing.
- `<DebugImageShape>` not present in this capture — check whether v1.24 actually emitted it (probe code may have errored silently) or whether `r.image` is `undefined` on this firmware. v1.25 must guard + report.

Foundation: comms-and-programming-guide 25.4.1.1 digest landed 2026-05-18,
formally confirming several v1.23 empirical findings and unblocking the
image-emission path. v1.23 baseline regression-passed across DMST 25.4.1.1
→ 26.1.0 (schema bit-identical, see B2 in session plan).

New first-class fields (promoted from v1.23 probes):
- `<SymbologyId>` ← `r.symbology.id` (AIM ID, e.g. `]d1`)
- `<SymbolQuality>` ← `r.symbology.quality` (0–100 decoder confidence)
- `<SymbolAngle>` ← `r.symbology.angle` (degrees)
- `<ModuleSizePx>` ← `r.symbology.moduleSize` (pixels/module)
- `<CalibrationDate>` ← `r.trucheck.calibrationDate` (promoted out of `<CustomNote>` workaround)
- `<FieldCalibrated>` ← `readerProperties.status3D.fieldCalibrated`
- `<FactoryCalibrated>` ← `readerProperties.status3D.factoryCalibrated`
- `<MinPassGrade>` / `<MinPassRaw>` ← `r.metrics.minPassGrade.{grade,raw}`
- `<BWGPercent>` ← `r.metrics.printGrowth.raw` via `mmPctAuto` (was empty in v1.23 — comms guide confirms `printGrowth` IS the source)

New probes:
- `<DebugImageShape>` — full enum of `r.image` (metadata only per comms guide; unblocks D4 image-load)
- `<DebugJpegProbe>` — `r.trucheck.jpegImage` LENGTH + 80-char preview. Full base64 payload **deferred to v1.25** pending size data — could be 50–300 KB which is 7–40× current XML size; need to verify Network Client + listener handle it before committing.
- `<DebugGS1>` — deep enum of `r.validation.gs1` (v1.23 shallow probe saw only `[obj]`)
- `<DebugDodUid>` — deep enum of `r.validation.dodUid`
- `<DebugBarcodeAsgn>` — deep enum of `r.barcodeAssignment` (undocumented sibling)
- `<DebugANUCase>` — case-mismatch resolver: `axialNonuniformity` (lower-u, empirically working) vs `axialNonUniformity` (upper-U, comms-guide spelling)
- `<DebugReaderProps>` — top-level enum of `readerProperties`

Retained (continuing baselines, bit-identical across DMST 25→26.1): `DebugModSize`, `DebugECCount`, `DebugMetricsKeys`, `DebugRSiblings`.

Dropped from v1.23 (questions definitively answered by filed v1.23 XML, see `samples/live-scans/v1.23-2026-05-18-*`):
`DebugSymbology` · `DebugCellDefects` · `DebugFPDefects` · `DebugDMCellDims` · `DebugValidation` · `DebugMetricShape`

`DebugMetricShape` finding worth carrying forward: `r.metrics.symbolContrast = {raw, grade}` — monolithic single-symbol grade with no hidden per-region structure. **ISO 15415 metrics in JS scope are flat; per-region grading definitively requires a separate data path (DMCC report engine), not a JS-scope re-probe.**

Diagnostic tag: `<PushScriptDiag>v1.24 q=r.trucheck m=found</PushScriptDiag>`.

---

### v1.23 — 2026-05-17

Filed: `v1.23.js` (also `dist/DmstPushScript_v1.23.txt`).

Live-confirmed: `samples/live-scans/v1.23-2026-05-18-Probe-DataMatrix-GS1Format06.xml`.

Added:
- `<Source>` emission from `r.source` (e.g. `DM475-63530E-PIPS-Verif-Lab`)
- `<DebugMetricsKeys>` — enumerates all keys of `r.metrics` with type stubs
- `<DebugRSiblings>` — enumerates all properties of `r.*` with type stubs
- `<DebugSymbology>` — deep dump of `r.symbology` (id, quality, contrast, moduleSize, corners, center, angle, size, failureCode)
- `<DebugCellDefects>`, `<DebugFPDefects>`, `<DebugDMCellDims>` — DPM-specific metric probes (NA on non-DPM scans, expected)
- `<DebugValidation>` — `r.validation` shape (state, method, failurePos, gs1, dodUid, matchString)
- `<DebugMetricShape>` — confirms universal `{raw, grade}` shape for metric objects
- Diagnostic tag: `<PushScriptDiag>v1.23 q=r.trucheck m=found</PushScriptDiag>`

v1.23 was the major probe-payload release. Surfaced the complete 30-metric
enumeration and 12-r-sibling enumeration; `r.image` confirmed present (sibling)
but accessor not yet probed.

---

### v1.16 — 2026-05-17 (earlier)

Filed: `v1.16.js`.

Mid-sprint snapshot. Probe envelope expansion before the full 30-key enumeration.
Used as a reference for the v1.18+ delta.

---

### v1.11 — 2026-03-30

Filed via raw XML capture: `samples/live-scans/probe-history/v1.11-2026-03-30-FullLive.xml` (6,961 chars).

Original push-script vintage that produced the XML the C# `DmstResultParser`
was written against. **This is the parser's reference grammar — fields added
since v1.11 are not yet known to the parser.**

---

### v1.10 — 2026-03-30

Filed via raw XML capture: `samples/live-scans/probe-history/v1.10-2026-03-30-FullLive.xml` (3,551 chars).

Initial probe iteration. Smaller payload than v1.11.

---

## Pending versions

- **v1.25** — scope driven by v1.24 DM confirmation findings AND the QR/loaded-image capture (see below):
  - **Commit to full `<JpegImageBase64>`** emission. v1.24 confirmed `r.trucheck.jpegImage` is a base64 string of ~9.9 KB for a typical scan — comfortably within Network Client buffer. Just emit the value as-is (it's already a string).
  - **Drop `axialNonUniformity` (upper-U), keep `axialNonuniformity` (lower-u) first-class.** Upper returned null; lower returned `{grade:D,raw:11.368}`. Replace the v1.24 `DebugANUCase` probe with the resolved wire.
  - **Re-probe `r.metrics.printGrowth` shape** to fix the still-empty `<BWGPercent>`. v1.24's `mmPctAuto(printGrowth)` wire yielded empty — need to verify whether `printGrowth` is defined and what its actual shape is on this build.
  - **Re-probe `r.validation.gs1`** — but FIRST set Application Settings → Data Format Check = **GS1**, then scan a GS1-encoded symbol. The all-undefined result in v1.24 was because Data Format Check = None → no GS1 check ran → stub object. Once set to GS1, enumerate `r.validation.gs1` keys directly rather than guessing names. Similarly for `r.validation.dodUid`: requires a DoD UID symbol and matching standard setting.
  - **Re-probe `r.metrics.minPassGrade` accessor** — UI shows Overall Pass Grade = 1.5 but `<MinPassGrade>` returned "NA". The property name or object path is wrong; probe `r.trucheck.minPassGrade`, `r.settings`, and enumerate `r.trucheck` top-level keys to find where the threshold lives.
  - **Deep-probe `r.barcodeAssignment.stats`** — top-level `result=-1` is "unassigned"; the `.stats` sub-object may carry useful diagnostics.
  - **Verify `<DebugImageShape>` emission** — it didn't appear in the v1.24 capture. Either `r.image` is undefined (probe should report that explicitly) or the probe code errored silently (needs try/catch guard).
  - Drop `r.decodeTime` / `r.triggerTime` wires (not in any v1.23 enumeration — were a misconception in the prior session-plan notes; the comms guide doesn't document them on `r` either).
  - **Add `<OpticsSource>` field** — emit `LiveScan` vs `LoadedImage` based on `ContrastUniformity === -1 && MRD === -1` (both must be −1). Two loaded-image captures confirm this: URL QR had SymbolAngle=360°, email QR had SymbolAngle=0° — both loaded images, so `SymbolAngle` is NOT a reliable discriminator. ContrastUniformity and MRD are always −1 on loaded images (not computed from optics) and always real values on live scans.
  - **Strip ECI prefix from `<DecodedData>` for QR** — QR codes with ECI mode 26 (UTF-8) emit a `\000026` header before the payload; strip it for clean display. Add a helper function in v1.25. ECI is indicated by AIM modifier digit: `]Q2`/`]Q4`/`]Q6` → ECI present (strip); `]Q1`/`]Q3`/`]Q5` → no ECI. See `references/standards/AIM-identifiers-honeywell-KB000063426.md` for full modifier table.
  - **QR-specific sentinel pattern confirmed**: ContrastUniformity=−1, MRD=−1, LLS/BLS/QZ grades=X on QR/loaded-image scans — these are "not applicable" sentinels, NOT errors. Parser must treat −1 as N/A, not as a grade failure.
  - **`]Q1` vs `]Q2` AIM ID modifier**: modifier digit encodes ECI presence. `]Q1` = QR without ECI (plain ASCII/byte — no prefix stripping needed). `]Q2` = QR with ECI designator (strip `\000026` UTF-8 header). Parser must branch on the modifier character, not the whole ID.
  - **`AGValue` can be negative on loaded images**: email QR capture showed AGValue=−0.5 with AGGrade=A. Parser must accept negative AG as valid (not an error sentinel); display as-is. Loaded-image scoring artifact.
  - **`ApplicationPass` reason tokens** — three confirmed so far: `Pass` / `Fail (Quality)` / `Fail (X Dimension out of Range)`. Parser should split on ` (` to extract the reason: `Quality` = grade below threshold; `X Dimension out of Range` = NominalXDim outside [Min, Max] X Dimension setting; format failures (GS1/HiBCC/ISO 15434) will add more when Data Format Check is set. Store as separate `ApplicationPassReason` field.
  - **Image upload accepts JPEG, not PRN**. D4 (WPF image-load flow) must write/convert to JPEG before pushing to device via DMCC.
  - **DMST PDF report cross-reference** (from `v1.24-2026-05-18-QR-Email-DMSTReport-catalog.md`) reveals additional v1.25 QR probe targets missing from push XML:
    - **Unit Serial** (`1A1903PP010754`) — not in any push output; probe `r.readerProperties.serialNumber` or `status3D` sub-path
    - **DataCodewords / ErrorCorrectionBudget / ErrorCapacityUsed** — all empty in push XML, populated in DMST report (44/26/0); probe `r.trucheck.symbols[0].{dataCodewords,ecCodewords,...}`
    - **ErrorCorrectionType wrong for QR**: push emits `ECC200` (Data Matrix label). QR uses L/M/Q/H. Probe `r.trucheck.symbols[0].ecLevel` or equivalent
    - **Data Mask Pattern** (value=2 for email QR) — absent from push; probe `r.trucheck.symbols[0].maskPattern` or similar
    - **QR-specific grade params** absent from push: ULP, URP, LLP (finder patterns), HCT, VCT (clock tracks), ALP (alignment), VIB, FIB — probe `r.trucheck.symbols[0].{ulp,urp,llp,hct,vct,alp,vib,fib}Grade` or equivalent
    - **SCPercent = NaN for loaded images** — push emits empty; parser must treat as N/A (not zero/error)
    - **MatrixSize=35×35 is wrong for QR v3** — DMST report confirms 29×29 (QR v3). DebugModSize sqrt=37 = 29 + 4 quiet-zone rows each side. For v1.25: probe `r.trucheck.symbols[0].{rows,cols}` for true module count
    - **DMST epoch timestamp bug**: DMST PDF shows `31-Dec-1970` for loaded images; push XML `<DateTime>` is correct. VTCCP must use push timestamp, not DMST report metadata
  - **Live DM scan cross-reference** (from `v1.24-2026-05-18-DM-Live-15434-catalog.md`) — ISO 15434 enabled scan, same UPCE-44960 calibration symbol. Additional v1.25 targets and schema corrections:
    - **DataCodewords / ECBudget / ECCapacity empty for LIVE DM too** — not a QR or loaded-image artifact. Push script has no working accessor for any of these three fields. Probe `r.trucheck.symbols[0].{dataCodewords,ecCodewords,ecCapacityUsed}` in v1.25
    - **ErrorCapacityUsed is computable**: ECBudget=24, ErrorsCorrected=7, ECCapacityUsed=14. Formula: `ECCapacityUsed = ErrorsCorrected × 2` (each RS correction uses 2 EC codewords). Fallback computation valid until probe finds the JS path
    - **ULQZ / URQZ / RUQZ / RLQZ grades are empty in push XML** — DMST shows all Grade A. Push schema fields exist but push script returns null/undefined. V1.25 probe `r.trucheck.symbols[0].{ulqz,urqz,ruqz,rlqz}` or equivalent
    - **TTR / CTR sub-grades**: DMST shows LQTTR/RQTTR/LQRTR/RQRTR and LQTCT/RQTCT/LQRCT/RQRCT all Grade A. Push XML shows TTRGrade=X, RTRGrade=X, TCTGrade=X, RCTGrade=X (aggregate X = "not broken down"). V1.25 probe for individual sub-grades
    - **PixelsPerModule field is redundant and empty** — `<ModuleSizePx>` is already correctly populated from `r.symbology.moduleSize`. `<PixelsPerModule>` is a second schema field targeting the same value via a different (unfound) accessor. Remove `<PixelsPerModule>` in B4 parser alignment — use only `<ModuleSizePx>`
    - **ISO 15434 data does NOT populate DebugGS1** — confirmed. `r.validation.gs1` only activates for GS1 format; Format 06 DI data is not recognized as GS1 AIs. VTCCP's in-house 15434 parser + DI lookup table (E1) is the only source for field-by-field breakdown
    - **ISO 15434 field-by-field parse is absent from push XML** — DMST shows the full structural + DI breakdown (header, term char, format type, sep char, each DI + value). This is a VTCCP differentiator: the push channel gives only the aggregate ApplicationPass; VTCCP's E1 parser gives per-field detail
    - **DI mapping confirmed from DMST 15434 parse** (ANSI MH10.8.2 Format 06): `18V` = Other Agency, `1P` = Supplier Part Number, `S` = Serial Number. Add to VTCCP DI lookup table
    - **DebugJpegProbe len=9868** for live 16×36 DM at aperture 17 — much smaller than loaded QR (19K–22K). V1.25 full JPEG emission is confirmed safe for live DM scans
    - **Average Grade (AG) discrepancy**: DMST shows AG=4.0/A; push XML AGGrade=B (mathematically correct per ISO 15415 §5: average of 8 primary grades = (2+4+4+2+1+4+4+4)/8=3.125→B). DMST appears to compute AG from a different parameter subset. Document in A4 parser-alignment-gap. Push XML AGGrade=B is the correct ISO 15415 value
    - **EncodedCharacters discrepancy**: push=33, DMST=38 (38 = all bytes including GS/RS/EO). Do not use `<EncodedCharacters>` for DMST comparison; use `<TotalCodewords>` instead (56=56 ✓)
    - **DebugModSize for DM confirmed 1-module quiet zone per side**: 16×36 → (16+2)×(36+2) = 18×38 = 684 ✓. `sqr=false` because rectangular. Rule: DM modulation grid = (rows+2)×(cols+2); QR grid = (version_size+8)²
    - **MRD breakdown components**: DMST shows "71% (79%−8%)" = Rl_min=79, Rd_max=8 at worst module. Push XML has MRD=71 ✓ but `<MinReflectance>` is empty. V1.25: probe `r.trucheck.symbols[0].{minLightReflectance,maxDarkReflectance}` or equivalent
  - The 21-unwired-metrics bulk extract concept is **deprecated**: v1.23 enumerated `r.metrics` and confirmed most "unwired" entries were either already wired (different name), DPM-only (NA on non-DPM scans), or `{raw:-1, grade:NA}` sentinels. Per-metric wire-up now happens case-by-case, not bulk.

---

## Parser-alignment note

The C# `DmstResultParser.cs` (657 lines, last touched 2026-03-27) was written
against the v1.10-v1.11 XML grammar. Since then we've added at minimum:

- `<PushScriptDiag>`, `<Source>`, `<FormalGrade>` reformatting
- All `<Debug*>` envelopes
- The 21 unwired metric promotions still pending in v1.24

When v1.24 lands, opening `architecture/parser-alignment-gap.md` and going
through every field becomes the natural next sync point.
