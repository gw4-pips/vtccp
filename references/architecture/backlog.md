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

## Reference Catalog

### CAT-1 — Scan catalog index update
Add scans #17 and #18 to the master scan index / session plan table.  
Files: `v1.36-2026-06-09-DM-24x24-GradeC-GS1-4AI-pharma.xml` and
`v1.37-2026-06-09-DM-24x24-GradeC-GS1-4AI-pharma.xml`.
