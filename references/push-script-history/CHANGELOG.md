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

**Probe release.** Status: drafted, not yet device-confirmed (B1 pending).

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

- **v1.25** — planned, depends on v1.24 device confirmation (B1) returning JPEG size data:
  - Commit to full `<JpegImageBase64>` payload emission once Network Client + listener confirmed to handle the size.
  - Resolve `<DebugANUCase>` outcome — drop the loser, keep the winner first-class.
  - Promote whichever of `<DebugGS1>` / `<DebugDodUid>` / `<DebugBarcodeAsgn>` returned structured payloads.
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
