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
- `<DebugGS1>` = all 187 AIs `undefined` — `r.validation.gs1` does NOT expose AIs as direct properties. Need different access pattern (likely a method like `.getAI(n)` or a `.parsed` sub-object). v1.25 probe.
- `<DebugDodUid>` = all 5 fields `undefined` — same conclusion as GS1; need different access pattern.
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

- **v1.25** — scope now driven by v1.24 live confirmation findings:
  - **Commit to full `<JpegImageBase64>`** emission. v1.24 confirmed `r.trucheck.jpegImage` is a base64 string of ~9.9 KB for a typical scan — comfortably within Network Client buffer. Just emit the value as-is (it's already a string).
  - **Drop `axialNonUniformity` (upper-U), keep `axialNonuniformity` (lower-u) first-class.** Upper returned null; lower returned `{grade:D,raw:11.368}`. Replace the v1.24 `DebugANUCase` probe with the resolved wire.
  - **Re-probe `r.metrics.printGrowth` shape** to fix the still-empty `<BWGPercent>`. v1.24's `mmPctAuto(printGrowth)` wire yielded empty — need to verify whether `printGrowth` is defined and what its actual shape is on this build.
  - **Re-probe `r.validation.gs1` and `r.validation.dodUid` access patterns.** Direct property access yielded all-undefined; try method-style access (`.getAI(n)`, `.parsed`, `.fields`, etc.) and key-enumeration.
  - **Deep-probe `r.barcodeAssignment.stats`** — top-level `result=-1` is "unassigned"; the `.stats` sub-object may carry useful diagnostics.
  - **Verify `<DebugImageShape>` emission** — it didn't appear in the v1.24 capture. Either `r.image` is undefined (probe should report that explicitly) or the probe code errored silently (needs try/catch guard).
  - Drop `r.decodeTime` / `r.triggerTime` wires (not in any v1.23 enumeration — were a misconception in the prior session-plan notes; the comms guide doesn't document them on `r` either).
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
