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

- **v1.24** — planned. Scope per last session:
  - Wire `r.symbology.moduleSize` → `<PixelsPerModule>`
  - Wire `r.symbology.id` → new `<AIMId>`
  - Wire `r.symbology.angle` → new `<SkewAngle>`
  - Wire `r.symbology.quality` → new `<DecoderQuality>`
  - Wire `r.decodeTime` / `r.triggerTime`
  - Bulk-extract all 21 unwired metrics as `<MetricName_Raw>` / `<MetricName_Grade>`
  - Probe `r.image` accessor + format
  - Probe `r.validation.gs1` and `.dodUid` shapes
  - Probe `r.barcodeAssignment`

---

## Parser-alignment note

The C# `DmstResultParser.cs` (657 lines, last touched 2026-03-27) was written
against the v1.10-v1.11 XML grammar. Since then we've added at minimum:

- `<PushScriptDiag>`, `<Source>`, `<FormalGrade>` reformatting
- All `<Debug*>` envelopes
- The 21 unwired metric promotions still pending in v1.24

When v1.24 lands, opening `architecture/parser-alignment-gap.md` and going
through every field becomes the natural next sync point.
