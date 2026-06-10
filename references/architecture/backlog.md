# VTCCP Implementation Backlog

Items logged here are confirmed-feasible, materials in hand, not yet started.
Order within each section = rough priority (top = sooner).

---

## GS1 / Push Script

### GS1-1 — applicationStdArray: full push emission + C# parsing + Excel surfacing
**Confirmed by**: Scan #18 (v1.37, 2026-06-09)  
**Sub-key schema**: `name` (string), `data` (string), `check` ("PASS"/"FAIL") — 3 keys per element  
**Element count**: 1 per GS1 format-check row (13 for a 4-AI pharma code)  
**Absent on non-GS1**: `len=-1`

**Work required**:
1. **Push script (v1.38)**: emit all elements (not just first 6) as structured XML.
   Two options:
   - Option A: flat elements `<GS1Row_0_Name>`, `<GS1Row_0_Data>`, `<GS1Row_0_Check>` per row
   - Option B: single encoded element `<GS1FormatRows>0|GS1 Header|<F1>|PASS;1|AI:GTIN|01|PASS;...</GS1FormatRows>`
   Prefer Option B — compact, single element, easy to split on `;` in C#.
2. **C# parser** (`DmstResultParser` / `VerificationXmlMap`): parse `<GS1FormatRows>` into
   `List<GS1FieldResult>` on `VerificationRecord`. Each entry: `Name`, `Data`, `CheckResult`.
3. **Excel** (`CwValuesSheetWriter` or new `GS1CheckSheetWriter`): GS1 Format Check table —
   Name / Value / Result columns; FAIL cells in red fill.
4. **VerificationRecord**: add `List<GS1FieldResult>? GS1FormatRows` nullable property.

**Note**: `ApplicationPass` (overall format pass/fail) already captured. This adds per-field detail.  
**Does NOT replace** `ApplicationStandard` / `ApplicationPass` / `ApplicationPassReason` — those stay.

---

### GS1-DL-1 — applicationStdArray behavior on GS1 Digital Link QR (len=0 — unresolved)
**Observed**: Scan #19 (2026-06-09, GS1 DL QR, Grade F). `DebugApplicationStdArray: len=0`.  
Contrast: Scan #18 (pharma GS1 DM, Grade C): `len=13`.

**Open question**: Is len=0 because:
- **(A) Grade F suppresses it** — firmware cannot walk field-by-field AI checks without a
  decoded codeword stream (DECODE=X). GS1 parser bails before populating the array.
- **(B) GS1 Digital Link uses a different pipeline** — DL URIs are parsed at the URL level,
  not as FNC1 AI streams. The DL parser recognizes the format as valid GS1 (hence
  `ApplicationPass=Fail(Quality)`, not `Fail(Data Format)`) but does not produce a
  row-per-field `applicationStdArray`.
- **(C) ★ MOST LIKELY — GS1 DL parsing is a DM395V / fw7 feature, not present on DM475V fw6**
  The DM475V runs fw 6.1.16_sr4. GS1 Digital Link support in the firmware's GS1 application
  parser may have been introduced in fw7 (confirmed shipping on DM395V). On fw6, the device
  may only recognize DL by URL pattern-match (enough to set `ApplicationStandard=Custom` and
  `ApplicationPassReason=Quality`) but has no DL AI-parsing engine → `applicationStdArray`
  is always empty for DL on this firmware/device combination.

**If (C) is correct**: `applicationStdArray` will NEVER be populated for GS1 DL QR on
fw 6.1.16_sr4 regardless of symbol quality. A passing DL QR scan would confirm this.
GS1 DL AI extraction on this device must come entirely from `DecodedData` via the gs1-syntax-engine
client-side library — the firmware offers no per-field data.

**If (C) is wrong** and a passing DL QR gives len>0 → revise to hypothesis (A).

**Impact on GS1-1**: GS1-1 scoped to `applicationStdArray` on traditional GS1 (FNC1 / `]d2` DM).
GS1 DL on fw6 needs a separate client-side extraction path. Do NOT block GS1-1 on this.

---

### GS1-DL-2 — VTCCP GS1 Digital Link detection rule (https:// prefix is NOT sufficient)
**Confirmed by**: Scan #19 (GS1 DL QR, `]Q1`, `DecodedData=https://id.gs1.org/...`)

**Finding**: GS1 Digital Link QR codes carry `SymbologyId=]Q1` — the same AIM ID as non-GS1
QR. There is NO symbology-layer distinction between DL QR and plain QR in the push XML stream.

**❌ WRONG approach — do NOT use**: Checking if `DecodedData` starts with `https://` (or
`http://`, `www.`). An enormous number of non-GS1 QR codes encode plain URLs — product
pages, contact cards, Wi-Fi credentials, payment links, marketing links. A URL prefix test
would produce massive false positives.

**✓ Correct detection strategy**: Pass `DecodedData` to the gs1-syntax-engine
(`vtccp/lib/gs1-syntax-engine/` v1.4.0) Digital Link validator. The library
can attempt to parse the URI as a GS1 DL and return AIs. If the parse succeeds with
at least one valid AI extracted → it is GS1 DL. If parse fails → treat as plain URL QR.

In practice the gs1-syntax-engine checks that:
- The path contains one of the GS1 primary key AIs (01/8006/8013/8010/255/8017/8018/414/417/8004)
  in the correct position (first non-stem path segment pair)
- Subsequent path segments follow valid GS1 AI structure

`ApplicationStandard=Custom` is a secondary confirming signal but cannot be used alone —
it is also set on GS1 DM (`]d2`) and is absent (or different) on non-GS1 QR.

**AIs extracted from scan #19 DL URI** (`https://id.gs1.org/01/09506000164960/22/80/10/ABC`):
- AI 01 = GTIN-14: `09506000164960`
- AI 22 = Internal Product Variant: `80`
- AI 10 = Batch/Lot: `ABC`

**Note on fw6 / DM395V**: See GS1-DL-1 hypothesis (C). Client-side DL parsing via
gs1-syntax-engine may be the ONLY path for AI extraction on this device/firmware regardless
of what the firmware's own GS1 parser does.

---

### GS1-DL-3 — FormalGrade two-token format on total fail (C# parser gap)
**Observed**: Scan #19: `FormalGrade=0/F` (2 tokens).  
All prior grading scans: `FormalGrade=2.3/C` or (per Wireshark) `1.0/16/660/45Q` (4 tokens).

**Problem**: Current C# parser may assume a fixed token count for FormalGrade. Total-fail
scans return only 2 tokens (numericGrade/letterGrade). Parser must handle both:
- 2-token: `0/F` → numericGrade=0, letterGrade=F, aperture/wavelength/lighting=null
- Standard: `2.3/C` or `1.0/08/660/45Q`

**Work required**: Defensive split on `/` in `FormalGrade` parsing; handle 2-token case
without throwing or silently truncating.

---

### GS1-2 — ANU raw unit formula (DM scans)
**Confirmed by**: Scans #17/#18 (2026-06-09) — push=11.5/7.4, DMST=0.1%  
**Prior**: Scan #10 (grade-D QR) push=3.9, DMST=3.9% (1:1).  
**Root cause**: `axialNonuniformity.raw` is a pixel-domain measurement on DM, not a percent.
DMST applies an unknown normalization. `mpa()` reads `.raw` and applies a (n>1 → use; n≤1 → ×100)
heuristic that is wrong for DM ANU.

**Work required**:
- Add probe `ekv(pk(q,"axialNonuniformity"), "q.axialNonuniformity")` to a future script version.
- Sub-keys expected: `.raw`, `.grade`, possibly `.percent` or `.normalized`.
- If `.percent` key exists and matches DMST value → use it; do not use `.raw` for ANU.
- If only `.raw` exists → the conversion formula is device-specific and must be reverse-engineered
  from more scans (ANU raw vs DMST percent pairs across range).

**Impact**: ANUPercent in Excel is currently wrong on DM scans. ANUGrade is correct (grade computed
by firmware, not by us). Clinical severity: low (grade is right; display value is wrong).

---

## Excel / Codeword Sheet

### CW-1 — Corrected codeword visual marker in CwValuesSheetWriter
**Confirmed by**: v1.36 HTML report `*` prefix on corrected codewords (scans #17/#18)  
**Current state**: `isCorrected` flag already on `CodewordValuesData` model; used to count
`ErrorsCorrected` total. Not visually indicated in the codeword grid cells.  
**DMST convention**: `*` prefix before codeword value in the Codewords table.

**Work required**:
- In `CwValuesSheetWriter`, when writing a codeword cell: if `isCorrected == true`, either:
  - Prefix the codeword value string with `*` (matches DMST convention), or
  - Apply a distinct cell fill (e.g. yellow or orange) in addition to the existing isBlack shading.
- Recommend both: `*` prefix + light orange fill for maximum visibility.
- No model change needed — `isCorrected` already present.

---

## Parser / Schema

### PARSE-1 — ECC200 lookup table: add 24×24 row
**Confirmed by**: Scan #17/#18 (24×24 outer, 22×22 data region, 36 data CW, 24 ECC CW)  
**Current table**: has 22×22 entry (30 data, 20 ECC).  
**Add**: `{matrixOuter: 24, matrixData: 22, dataCW: 36, eccCW: 24}`  
Note: push XML `<MatrixSize>24x24</MatrixSize>` = outer dimension. Key lookup must use outer dim.

---

### PARSE-2 — DDGrade mismatch investigation
**Observed**: `DDGrade` emits `X` on grade-C DM scan; DMST HTML shows DFPD (Distributed FPD) = A.  
**Possible causes**:
  (a) `distributedDamageGrade.grade` is literally `X` — firmware quirk for a metric that isn't
      applicable or is calculated differently from DFPD
  (b) `distributedDamageGrade` is a different metric than DMST's "Distributed FPD (DFPD)"
**Work required**: Add `ekv(pk(q,"distributedDamageGrade"), "q.distributedDamageGrade")` probe
to a future script version to expose sub-keys. If `grade=X` is confirmed, clarify what the
field represents vs DMST's DFPD row.

---

## Trigger / Result Capture

### TRIG-0 — ★ ACTIVE BUG — Trigger results not arriving (CP trigger AND TC trigger both dead)
**Observed**: 2026-06-10. Neither VTCCP (CP) software trigger nor DMST (TC) Verify button
trigger results are arriving on the HTTP subscriber. Both paths are silent.

**Context**: Earlier in this session, external production-line trigger results WERE arriving
(that was the problem TRIG-1 was written to solve). Now neither external nor software triggers
produce results on the subscriber.

**Most likely cause (2026-06-10 update)**: Dual live mode interference — DMST and VTCCP
were both open simultaneously. DMST may have opened its own `GET /events?enable` subscription,
intercepting the result stream before VTCCP received it, OR one client's connection displaced
the other on the device side. When VTCCP was closed and only DMST was open (or vice versa),
results may have resumed normally.

**Other possible causes**:
- HTTP subscriber connection dropped or was not re-established after a mode change
- TRIGGER.TYPE left in a non-standard state from a prior session
- Device reboot or config reset cleared push XML script or subscriber state
- Port 44444 connection saturation (too many open connections from dual live mode)

**First things to check when revisiting**:
1. Is the HTTP subscriber connection still alive? (Send a keepalive or reconnect)
2. What is `TRIGGER.TYPE` currently set to on the device?
3. Does triggering from the DMST Verify button produce a result on the HTTP channel?
   (This was previously the most reliable path — Wireshark confirmed)
4. Is the push XML script still installed on the device?

**Do not debug further until user is ready to revisit.**

---

### TRIG-1 — Pending-trigger correlation flag (filter external trigger results)
**Problem**: The HTTP subscriber captures ALL verification results regardless of trigger source.
When the production line is live and its external hardware trigger fires the DM475V, VTCCP
currently intercepts and records those results — which is undesirable. VTCCP must only record
scans that it explicitly initiated.

**Fix (advice, not yet implemented)**:
- Add `bool _pendingVerification` flag to `DeviceSession`, default `false`.
- When VTCCP sends a trigger command → set `_pendingVerification = true`.
- On result arrival in HTTP subscriber:
  - `_pendingVerification == true` → process + clear flag.
  - `_pendingVerification == false` → silently discard.
- Add a timeout (e.g. 5 s) that auto-clears the flag if no result arrives — prevents a stale
  `true` from capturing the next production scan after a trigger failure.
- No DMCC, firmware, or TRIGGER.TYPE changes required. Filter lives entirely in VTCCP.

---

### TRIG-2 — Passive capture / audit mode (possible future feature)
**Concept**: VTCCP operates as a silent observer while the production line runs — capturing
every hardware-triggered verification result without operator action. Useful for compliance
audit trails, batch quality records, or remote monitoring.

**Status**: POSSIBLE FUTURE FEATURE. NOT a default or normal ops mode. Must be an explicitly
toggled mode, clearly labelled, never active unless the operator enables it.

**Open design questions** (do not resolve yet):
- Where in the UI does the operator enable/disable audit mode?
- Does audit mode write to a separate session/file from operator-triggered verification?
- What happens to the pending-trigger flag (TRIG-1) when audit mode is active?

---

### TRIG-3 — Dual live mode conflict (VTCCP live + DMST live simultaneously)
**Observed**: When both VTCCP's live view and DMST's live view are open at the same time,
both are polling the device and both are receiving results. Conflicts observed include:
result interception, possible frame contention.

**Open question (do not resolve yet)**: Who has control when both are open?
- Does DMST's `GET /svg_image.img` polling interfere with VTCCP's `GET /events?enable`
  subscription? (Likely no — they are separate connections.)
- Does a scan triggered from either side get captured by both? (Very likely yes — both
  subscribers will receive it.)
- Can VTCCP detect that DMST is connected? (Unknown — no DMCC query for active connections.)

**Note only**: This conflict needs a defined policy before dual-mode use is allowed.
Do not design VTCCP assuming it is the sole client of the device.

---

## Live View (VTCCP Camera Panel)

### LV-1 — ROI setting in VTCCP live view
**Context**: `r.image.RoI` (left/top/right/bottom in sensor pixels) is already returned per scan
in push XML (DebugRImageRoI confirmed this). The device's `DECODER.ROI` DMCC parameter
controls the active decode region.

**Feature**: Allow the operator to set the ROI directly from within VTCCP's live view panel —
e.g. drag a rectangle on the live image, or enter pixel coordinates — and have VTCCP issue a
`DMCC SET DECODER.ROI` to the device.

**Why**: When both VTCCP and DMST live modes were observed open simultaneously, it became
apparent that VTCCP's live view needs the same framing capability DMST has. Without ROI
control, VTCCP cannot direct the decoder to a specific region of the field of view.

**Prerequisite**: Software trigger must be working (TRIG-1 resolved) and live view image
delivery path must be confirmed. This is a D4-scope feature.

---

### LV-2 — Centered crosshair superimposed on live view at all times
**Feature**: A fixed centered crosshair (X-pattern, full-width horizontal + full-height
vertical lines intersecting at image center) rendered over the live view image at all times,
regardless of zoom or ROI.

**Purpose**: Alignment aid — operator can position the symbol under the verifier by centering
it on the crosshair before triggering a scan. Standard practice in optical verification setups.

**"At all times"**: Crosshair is always visible — not toggled, not hidden during scan, not
dependent on any scan state. It is a permanent overlay on the live camera feed.

**Implementation note (advice only)**: In WPF, render as a `Canvas` overlay on top of the
`Image` control. Two `Line` elements — horizontal center and vertical center — with a
contrasting color (e.g. semi-transparent red or green) and thin stroke. Scales automatically
with the Image control's layout bounds regardless of image resolution.

---

## Reference Catalog

### CAT-1 — Scan catalog index update
Add scans #17 and #18 to the master scan index / session plan table.  
Files: `v1.36-2026-06-09-DM-24x24-GradeC-GS1-4AI-pharma.xml` and
`v1.37-2026-06-09-DM-24x24-GradeC-GS1-4AI-pharma.xml`.
