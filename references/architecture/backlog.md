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

**New evidence from scan #20 (2026-06-10)**: USPS DataMatrix (`]d1`, non-GS1, Grade A) gave
`len=1` — a single applicationStdArray element for the "FNC1 Required" check row that failed.
This rules out "non-GS1 symbology" as a cause of len=0. The `]d1` non-GS1 DM IS processed
by the GS1 parser and CAN produce non-zero applicationStdArray. Therefore, the DL QR's
len=0 is specifically due to Grade F (hypothesis A) OR the DL pipeline not generating
per-field rows (hypotheses B/C). The non-GS1 DM result does not distinguish between these.

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

`ApplicationStandard=Custom` **cannot be used as a GS1 indicator at all.**
Scan #20 (USPS `]d1` non-GS1 DM) confirmed `ApplicationStandard=Custom` is set whenever
the GS1 application parser is active — regardless of whether the symbol is GS1 or not.
It reflects "GS1 parser ran" not "symbol is GS1 format".

**AIs extracted from scan #19 DL URI** (`https://id.gs1.org/01/09506000164960/22/80/10/ABC`):
- AI 01 = GTIN-14: `09506000164960`
- AI 22 = Internal Product Variant: `80`
- AI 10 = Batch/Lot: `ABC`

**Note on fw6 / DM395V**: See GS1-DL-1 hypothesis (C). Client-side DL parsing via
gs1-syntax-engine may be the ONLY path for AI extraction on this device/firmware regardless
of what the firmware's own GS1 parser does.

---

### GS1-DL-3 — FormalGrade push XML is ALWAYS 2-token (not a fail anomaly)
**Revised by**: Scan #20 (Grade A USPS DM): `FormalGrade=3.5/A` — also 2 tokens.

**Confirmed pattern** across all scans:
- `3.5/A` — Grade A, scan #20
- `2.3/C` / `2.4/C` — Grade C, scans #17/#18
- `0/F` — Grade F, scan #19

Push XML FormalGrade format is ALWAYS `numericGrade/letterGrade` (2 tokens).
The 4-token format (`3.5/08/660/45Q` with aperture/wavelength/lighting) appears only in
the HTML report and the Wireshark codes.xml HTTP channel — NOT in push XML.

**Work required**: C# parser must always treat push XML FormalGrade as 2-token.
Aperture, wavelength, and lighting are NOT in push XML FormalGrade — they are separate
push XML elements (`ApertureRef`, `Wavelength`, `Lighting`).

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

**Scope extended by scan #20**: ANUPercent mismatch also occurs on rectangular DM (`]d1`, 12×36
USPS symbol) — push=59.5, DMST=0.6%. This is not limited to GS1 DM (`]d2`). Affects ALL DM
symbology variants on fw 6.1.16_sr4.

**Impact**: ANUPercent in Excel is currently wrong on all DM scans. ANUGrade is correct (grade
computed by firmware, not by us). Clinical severity: low (grade is right; display value is wrong).

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

### PARSE-4 — ECC200 lookup table: add 22×22 outer row (20×20 data region)
**Confirmed by**: Scan #21 (2026-06-10, 22×22 outer, Grade C)  
**From HTML**: TotalCW=50, DataCW=30, ECCW=20  
**Add**: `{matrixOuter: "22x22", dataCW: 30, eccCW: 20, totalCW: 50}`

Note: This is the 22×22 OUTER symbol (20×20 data region). It is DISTINCT from the entry
documented for scans #17/#18, which has a **24×24 outer** symbol with a 22×22 DATA region.
Push XML `<MatrixSize>22x22</MatrixSize>` vs `<MatrixSize>24x24</MatrixSize>` — do not conflate.
The existing comment in PARSE-1 ("has 22×22 entry — 30 data, 20 ECC") refers to THIS entry,
not the 24×24 entry. Verify that the C# table already has `"22x22"` → 30/20 and that PARSE-1
adds `"24x24"` → 36/24 as a separate row.

---

### PARSE-3 — ECC200 lookup table: add 12×36 row (rectangular USPS postage DM)
**Confirmed by**: Scan #20 (2026-06-09, USPS 12×36 DM, Grade A)  
**From HTML**: TotalCW=40, DataCW=22, ECCW=18  
**Add**: `{matrixOuter: "12x36", dataCW: 22, eccCW: 18, totalCW: 40}`  
Note: rectangular DM sizes use both dimensions — key lookup must handle `NxM` format,
not just a single integer. The C# table must index on the full string `"12x36"` (from
push XML `<MatrixSize>12x36</MatrixSize>`).

---

### PARSE-1 — ECC200 lookup table: add 24×24 row
**Confirmed by**: Scan #17/#18 (24×24 outer, 22×22 data region, 36 data CW, 24 ECC CW)  
**Current table**: has 22×22 entry (30 data, 20 ECC).  
**Add**: `{matrixOuter: 24, matrixData: 22, dataCW: 36, eccCW: 24}`  
Note: push XML `<MatrixSize>24x24</MatrixSize>` = outer dimension. Key lookup must use outer dim.

---

### DM-RECT-1 — Rectangular DM: 12 grade parameters not emitted by push XML
**Confirmed by**: Scan #20 (12×36 USPS DM, Grade A). Push XML emits all rectangular-specific
grade fields as empty; HTML report has all 12 as Grade A.

**Parameters missing from push XML (rectangular DM only)**:

| HTML row | Parameter | Push XML field |
|---|---|---|
| 11a | Upper Left Quiet Zone (ULQZ) | `ULQZGrade` — empty |
| 11b | Upper Right Quiet Zone (URQZ) | `URQZGrade` — empty |
| 12a | Right Upper Quiet Zone (RUQZ) | `RUQZGrade` — empty |
| 12b | Right Lower Quiet Zone (RLQZ) | `RLQZGrade` — empty |
| 13a | Left Top Transition Ratio (LQTTR) | `ULQTTRGrade` / `ULQTTRPercent` — empty |
| 13b | Right Top Transition Ratio (RQTTR) | `URQTTRGrade` / `URQTTRPercent` — empty |
| 14a | Left Right Transition Ratio (LQRTR) | `LLQRTRGrade` / `LLQRTRPercent` — empty |
| 14b | Right Right Transition Ratio (RQRTR) | `LRQRTRGrade` / `LRQRTRPercent` — empty |
| 15a | Left Top Clock Track (LQTCT) | `ULQTCTGrade` — empty |
| 15b | Right Top Clock Track (RQTCT) | `URQTCTGrade` — empty |
| 16a | Left Right Clock Track (LQRCT) | `LLQRCTGrade` — empty |
| 16b | Right Right Clock Track (RQRCT) | `LRQRCTGrade` — empty |

**Also**: `TQZGrade` and `RQZGrade` emit `X` on rectangular DM — these are square-DM
fields (top/right quiet zone for L-shaped finder pattern) that do not apply to rectangular
symbols. Do NOT treat `X` as a letter grade in this context.

**Additionally**: `DDGrade=X` on Grade A rectangular DM while HTML shows DFPD=A — same
mystery as PARSE-2, now confirmed to occur on rectangular DM as well.

**Work required**:
- Future push script version must probe rectangular-DM-specific sub-keys from
  `q.trucheck` or the grade object to extract these 12 parameters.
- The push XML fields are already declared (they exist as empty elements); the probe
  script just needs to populate them from the appropriate JS paths.
- VTCCP parser and Excel sheet must accept and display all 12 rectangular-specific fields.
- Parser must guard: if symbol is rectangular DM → render ULQZ/URQZ etc. in results;
  if square DM → render TQZ/RQZ; never display `X` as a printed grade for inapplicable fields.

---

### DM-BIN-1 — Binary DecodedData (Base256 encoding) — safe handling required
**Confirmed by**: Scan #20 (USPS DM, Base256 encoding, codeword 231 = switch to Base256)

**Problem**: When the DM symbol uses Base256 encoding, `DecodedData` contains raw binary
bytes (values 0x00–0xFF). These are not printable UTF-8 strings. VTCCP must handle
this gracefully in:
1. **XML serialization**: Binary bytes in a push XML `<DecodedData>` element can corrupt
   the XML parser (NUL bytes are illegal in XML 1.0). The current push XML delivery
   already contains these — the parser must tolerate them.
2. **C# string representation**: Do not expose raw binary as a `string` to Excel or report
   writers. Either hex-dump, escape as `<xNN>`, or detect and flag as binary content.
3. **GS1 parsing**: Do NOT attempt GS1 AI parsing on binary `DecodedData`. Gate on
   `ErrorCorrectionType=ECC200` AND `SymbologyId=]d1`/`]d2` AND content is printable.
4. **Display in VTCCP UI**: Show a safe representation — e.g. `[Binary data: 20 bytes]`
   followed by a hex preview of the first N bytes, then the printable suffix `000524`.

**Detection heuristic**: If `DecodedData` contains bytes < 0x20 (other than tab/CR/LF),
treat as binary. The USPS example has NUL (0x00), STX (0x02) as clear indicators.

**Note**: The Encodation Analysis table in the HTML report (`ErrorCorrectionType` field)
will show `ECC 200` for both binary and text DM. The mode switch is codeword 231 (Base256)
in the Codewords table — not currently accessible from push XML.

---

### GS1-IMP-1 — Implied FNC1 parser (missing-FNC1 `]d1` DM with GS1-structured data)
**Confirmed by**: Scan #21 (2026-06-10, `]d1` DM, data `011065316030393010201703209017190831`)
**Context**: One of the most common GS1 DM production defects. The symbol encodes correct
GS1 AI data but the FNC1 first-position codeword was omitted. DMST TC fails immediately
("Application Header is Expected") and shows no AI parse at all. VTCCP should do better.

**What VTCCP should do**:
When a `]d1` (non-GS1) DM result arrives and `ApplicationPass=Fail(Data Format)`:
1. Check if `DecodedData` is all-printable ASCII (not binary — gate on DM-BIN-1 check first)
2. Check if `DecodedData` starts with a 2-4 digit string that is a known GS1 primary key AI
   (01, 00, 414, 417, 8004, etc.) — see gs1-syntax-engine AI table
3. If yes → run the **implied-FNC1 backtracking parser**:
   - Use the GS1 AI length table to identify fixed-length AIs unambiguously
   - For variable-length AIs (10, 21, 30, etc.) try each possible length, check if
     remainder starts with a valid AI, recurse
   - Collect ALL complete parses (no unmatched remainder)
4. Report:
   - If exactly 1 valid parse → show AIs with label **"⚠ Implied GS1 (FNC1 missing)"**
   - If 2+ valid parses → show all candidates with label **"⚠ Ambiguous — FNC1 missing"**
   - If 0 valid parses → not GS1, show raw data only
5. Always surface `ApplicationPass=Fail(Data Format)` — implied parse is supplemental info,
   NOT a pass override. The symbol is still a GS1 format failure.

**Scan #21 parse result — TWO valid parses (ambiguous)**:

Parse A (4 AIs):
```
AI 01  GTIN-14   10653160303930  (check digit ✓)
AI 10  Lot       201703
AI 20  Variant   90
AI 17  Expiry    190831  →  2019-08-31
```

Parse B (3 AIs):
```
AI 01  GTIN-14   10653160303930  (check digit ✓)
AI 10  Lot       2017032090
AI 17  Expiry    190831  →  2019-08-31
```

Both consume all 36 characters exactly. Both pass AI lint rules. Cannot be resolved
without additional context (knowledge of whether AI 20 is expected in this product's label).

**Implementation approach**:
- The gs1-syntax-engine v1.4.0 (`vtccp/lib/`) has an AI table and lint validators.
  Use it for AI recognition and value validation.
- The backtracking parser is a small recursive function — roughly 50–100 lines in C#.
  Input: the decoded string (with implied FNC1 prepended as `\x1D`). Output: list of
  `List<GS1FieldResult>` parse candidates.
- This is independent of `applicationStdArray` (which the firmware doesn't populate for
  `]d1` missing-FNC1 scans — confirmed len=1 with only the failure row on scan #21).

**ANU note from scan #21**:
On this scan, push XML `ANUPercent=2` matched DMST `2.0%` exactly (1:1). Prior DM scans
with larger ANU values (11.5, 7.4, 59.5) all showed large mismatches. Suggests the
mismatch is non-linear — small ANU values agree; large values diverge. This narrows the
GS1-2 investigation (see that item).

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

### DMST-1 — DMST TC instability after Grade F QR scan (post-verify image lost)
**★ CONFIRMED REPEATABLE — QR scan → DMST TC post-verify image lost, every time (2026-06-10)**

**Occurrences observed**:
1. Scan #19 (Grade F GS1 DL QR) → post-verify image lost → config restore required
2. Any QR scan after restore → post-verify image lost again immediately → second config restore

**Pattern**: It is NOT specific to Grade F or total-fail QR symbols. Any QR scan on this
device/firmware/DMST version combination reproduces the failure. Config restore recovers it;
the next QR scan loses it again. DM scans do NOT trigger the issue.

**Root cause (working hypothesis)**: fw 6.1.16_sr4 + this DMST version have an incompatibility
in the QR result display pipeline — possibly the QR post-verify image is delivered in a format
or path that DMST's TC panel does not handle, causing the image widget to enter a bad state that
persists until a config restore clears it.

**This is a DMST software bug, not a DM475V firmware issue and not a VTCCP issue.**
Do not attempt to diagnose further — it is not VTCCP's component.

**Operational consequence**: QR scans on this lab unit require a config restore after each
scan if DMST TC post-verify image display matters. For VTCCP probe sessions: run all DM scans
first, then QR scans at the end of the session (or accept the post-verify image loss for QR).

**Note for VTCCP**: VTCCP's result display pipeline must be hardened against Grade F / total-fail
scans AND against any scan where the post-verify image is absent or malformed. The UI must
degrade gracefully — do not throw or leave the display in an indeterminate state.

**Device recovery**: Config restore recovers DMST TC each time. No permanent device damage.
Two restorations confirmed 2026-06-10.

---

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
