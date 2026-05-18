# ISO/IEC 15415:2024 (3rd Edition, DRAFT) — Digest

**Full title**: Information technology — Automatic identification and data
capture techniques — Bar code symbol print quality test specification —
Two-dimensional symbols

**Edition**: Third (Draft International Standard, dated 2024-10-16; track-change
PDF), cancels and replaces 2nd ed (2011)

**Source file**: `ISO-IEC-15415-ed3-2024-print-quality-2D-DRAFT.pdf`

**Scope**: Defines the methodology for **measuring and grading 2D barcode
symbol print quality** (matrix + multi-row). This is what the verifier
*does*. ISO/IEC 15426-2 governs whether the verifier *conforms*.

**Editorial caveat**: This is the DIS (Draft International Standard) version
with track-changes from the 2023 working draft → 2024 approval cycle. "Very
close to but not the current published version" per project owner. Should be
substantively identical to the published version for our reference purposes,
but minor wording or table-number drift is possible.

---

## 1. Parameter definitions (matrix symbols)

ISO/IEC 15415 uses a **numeric scale 4.0 to 0.0 in steps of 0.1**.
Historically these mapped to A/B/C/D/F letter grades; **the 3rd edition
formalizes decimal grade as the primary metric** (continuous grading,
0.1 resolution).

| Parameter | Symbol | What it measures | Grade 4.0 boundary | Grade 0.0 boundary | Mandatory |
|---|---|---|---|---|---|
| **Decode** | — | Reference Decode Algorithm successfully reads symbol | Decodable | Undecodable | Yes |
| **Symbol Contrast** | SC (ΔRsc) | Rmax − Rmin reflectance difference | ≥ 70 % | < 20 % | Yes |
| **Modulation** | MOD (Mmod) | Module contrast ÷ symbol contrast (uniformity of light/dark modules) | ≥ 0.50 | < 0.20 | Yes |
| **Fixed Pattern Damage** | FPD | Damage to finder patterns, quiet zones, clock tracks | No damage | Total failure | Yes |
| **Axial Nonuniformity** | ANU (MANU) | Deviation in X/Y aspect ratio | ≤ 0.06 | > 0.20 | Yes |
| **Grid Nonuniformity** | GNU (MGNU) | Deviation of module centers from ideal grid | ≤ 0.38 | > 0.75 | Yes |
| **Unused Error Correction** | UEC (MUEC) | Fraction of error correction capacity remaining after decode | ≥ 0.62 | < 0.20 | Yes |
| **Reflectance Margin** | RM | How close modules are to binarization threshold (combined w/ MOD in Ed.3) | High margin | No margin | Optional* |
| **Print Growth** | PG | Deviation of actual module size from nominal (growth/loss) | ≤ ±0.10 | > ±0.60 | **Ed.3: now graded** |
| **Contrast Uniformity** | CU | Local contrast variations across symbol | High | Low | Optional |

\* "Optional" parameters may be **required by specific symbology or
application standards** even when not mandated by 15415 itself.

---

## 2. Reportable fields (schema baseline)

A compliant verifier report **MUST** include:

1. **Overall Symbol Grade** — numeric (e.g. `2.8`); may include historical
   letter (e.g. `B`).
2. **Grading String** — `Grade/Aperture/Wavelength/Lighting`
   (e.g. `2.8/05/660/45Q`).
   - **Grade** — 0.0 to 4.0
   - **Aperture** — reference number (`05` = 5 mil) or diameter in mils
   - **Wavelength** — nm (e.g. `660`) or `W` for white light
   - **Lighting** — angle + orientation suffix (see §4 below)
3. **Individual parameter grades** — all mandatory parameters listed above.
4. **Decoded data** — raw string extracted from the symbol.

**VTCCP relevance**: The `<FormalGrade>1/D</FormalGrade>` field we see in
DMST XML is exactly the format `Grade/Lighting` (truncated form). The full
canonical form `Grade/Aperture/Wavelength/Lighting` is what `1/D` represents
in shorthand. The standard's full string form is closer to
`1.0/17/660/45Q` given our v1.23 scan's `ApertureRef=17 Wavelength=660
Lighting=45Q`.

---

## 3. Reference decode algorithm

The standard **does NOT** define decoding for every symbology. Instead it
defers to the **symbology specification**:
- Data Matrix → ISO/IEC 16022
- QR Code → ISO/IEC 18004
- Aztec → ISO/IEC 24778
- PDF417 → ISO/IEC 15438

Role of the reference decode:
- Define how to find the grid
- Identify module centers
- Perform error correction

**Binarization** uses the specific thresholding algorithm in **Annex A**
(histogram-based, Otsu-like — see §8).

---

## 4. Optical setup → DMST XML field mapping

The optical setup determines the "physics" that the grade is computed
against. **These are the values that MUST be reported alongside the grade**
because the grade is meaningless without knowing the optical conditions.

| Concept | DMST XML field | Notes |
|---|---|---|
| Effective aperture | `<ApertureRef>17</ApertureRef>` | NOT a physical lens opening — a mathematical convolution (blurring) applied to the image. Common values: `02` (2 mil), `05` (5 mil), `08` (8 mil), `17` (17 mil — used in our calibrated scans). |
| Light wavelength | `<Wavelength>660</Wavelength>` | Standard is 660 nm (red). White light reports as `W` with required color temperature (e.g. 5400K). |
| Lighting geometry | `<Lighting>45Q</Lighting>` | Angle (45° default for printed labels; 30°/90° for DPM or special textures) + suffix. |

**Lighting suffix codes:**
- `Q` — 4-sided (omnidirectional, default)
- `T` — 2-sided
- `S` — 1-sided

---

## 5. Symbol grade computation — the "lowest common denominator" rule

The overall symbol grade is the **lowest grade received by any of the
mandatory parameters**. A symbol with all 4.0s but one 1.8 gets an overall
grade of 1.8.

**Continuous grading** (new in Ed. 3): grades are reported at 0.1 resolution,
no longer rounded to whole-number letter grades. This matches what we see
in v1.23 XML: `<UECPercent>41.7</UECPercent>` `<UECGrade>C</UECGrade>` —
the verifier internally has the decimal grade but DMST emits the legacy
letter form. **VTCCP should probably preserve both representations** —
many existing systems still consume letter grades, but the standard now
prefers decimals.

---

## 6. Application standards (GS1, FDA UDI, MIL-STD-130, etc.)

Application standards layer **pass/fail thresholds** on top of 15415's
measurement methodology.

- **Minimum pass grade**: typically `1.5` (C) or `2.5` (B) depending on
  industry
- **GS1**: requires grade ≥ 1.5 for most retail; ≥ 2.5 for healthcare /
  FDA UDI
- **Data validation**: adds requirements for data formatting (FNC1, AI
  prefixes) — **not part of 15415** but commonly reported by verifiers
  via `<ApplicationStandard>` field

**Key new metric in our v1.23 enumeration**: `r.metrics.minPassGrade` —
this exposes the application-standard's threshold directly. **VTCCP should
promote this to a first-class column** so the report can compute
`OpticsCompliant ∧ OverallGrade ≥ minPassGrade → PASS` deterministically
without relying on the firmware's `<ApplicationPass>` field.

---

## 7. Key technical changes in 3rd edition

1. **Decimal grading** — formalized 0.1 resolution (was integer 0-4 in
   Ed. 2)
2. **Histogram-based threshold** — more robust binarization (Annex A)
3. **Print Growth now graded** — formal parameter, no longer just
   informative
4. **Modulation + Reflectance Margin combined logic** — Parameter Overlay
   (Annex E) — error correction "masks" defects by evaluating cumulative
   codeword grades
5. **Removal of 5-rotation averaging** — replaced with single-scan
   requirement, relying on ISO 15426-2 verifier conformance for stability
   (this is significant — it means a conforming verifier scans once, not
   five times averaged, in Ed. 3)

**VTCCP-relevant implication of #5**: The DMST XML may report only a single
scan's worth of metrics in Ed.3 mode, not averaged values. The empty `Avg*`
columns in our 1D scan output (`AvgEdge`, `AvgRlRd`, etc.) may stay empty
on 2D scans by design under Ed.3. They were never meant for 2D — they're
the 1D averaged-profile fields per ISO 15416. **Confirms our scratchpad
note that Avg* fields are 1D-specific.**

---

## 8. Annexes

| Annex | Status | Contents |
|---|---|---|
| **A** | Normative | Histogram-based thresholding algorithm (Otsu-like variance minimization to find optimal global threshold) |
| **B** | Informative | Troubleshooting guide for low grades ("Low SC? Check substrate opacity"; "Low MOD? Check ink coverage"; etc.) |
| **C** | Informative | Guidance on selecting aperture + lighting for application specs |
| **D** | Informative | Substrate characteristics (paper, plastics, metals — how each affects measurements) |
| **E** | Informative | **Parameter Overlay** — method to account for error correction "masking" defects by evaluating cumulative codeword grades. Important — this is how Ed.3 handles the Modulation + Reflectance Margin combined logic. |

---

## 9. Key glossary terms

| Term | Meaning |
|---|---|
| **X-Dimension** | Nominal width of a single module. Reported in mil (1/1000 in) by Cognex (`<NominalXDim>21.4 mil</NominalXDim>`) |
| **Reflectance (R)** | Lightness of a surface, 0-100 % |
| **Ecap** | Total error correction capacity of the symbology |
| **AIDC** | Automatic Identification and Data Capture (the ISO/JTC 1/SC 31 domain) |
| **Reference Decode** | The standardized decode algorithm per symbology spec |
| **Convolution aperture** | The mathematical "blur" applied to the image before measurement — simulates a physical aperture but is done in software |

---

## 10. Standards this depends on / cross-references with

| Standard | Provides | Relationship |
|---|---|---|
| ISO/IEC 15416 | Linear (1D) print quality | Multi-row symbols (PDF417) use 15416 row-by-row methodology |
| ISO/IEC 15426-2 | **Verifier conformance** | Validates that hardware accurately implements 15415 |
| ISO/IEC 16022 | Data Matrix symbology + reference decode | Symbology under test |
| ISO/IEC 18004 | QR Code symbology + reference decode | Symbology under test |
| ISO/IEC 24778 | Aztec symbology + reference decode | Symbology under test |
| ISO/IEC 15438 | PDF417 symbology + reference decode | Multi-row symbology under test |
| **ISO/IEC 29158** | **DPM (Direct Part Marking)** print quality methodology | **Use instead of 15415 for laser/dot-peen/etched marks**. The "AIM-DPM" mode in our verifier. |
| ISO/IEC 19762 | AIDC harmonised vocabulary | Terms + definitions |

---

## 11. Why this matters for VTCCP

1. **Schema validation**: Every mandatory parameter from §1 must have a
   first-class column in our 167-column schema. Checklist:
   - Decode → `<DecodeGrade>` ✓
   - Symbol Contrast → `<SCPercent>` / `<SCGrade>` / `<SCRlRd>` ✓
   - Modulation → `<MODGrade>` ✓
   - Fixed Pattern Damage → `<FPDValue>` / `<FPDGrade>` ✓
   - Axial Nonuniformity → `<ANUPercent>` / `<ANUGrade>` ✓
   - Grid Nonuniformity → `<GNUPercent>` / `<GNUGrade>` ✓
   - Unused Error Correction → `<UECPercent>` / `<UECGrade>` ✓
   - Reflectance Margin → `<RMGrade>` ✓
   - Print Growth → likely `<BWGPercent>` (currently empty); **wire from
     `r.metrics.printGrowth` in v1.24**
   - Contrast Uniformity → `<ContrastUniformity>` ✓

2. **Continuous-grade preservation**: Ed.3 prefers decimal grades. We
   currently emit letter grades from DMST. **v1.25+ should preserve
   `r.metric.{name}.raw` decimal value alongside letter grade.**

3. **Application-pass logic**: `r.metrics.minPassGrade` is the threshold
   for `ApplicationPass` boolean. Promote to first-class column +
   compute pass/fail independent of firmware's verdict (defensive
   computation).

4. **Grading string format**: VTCCP report should display the canonical
   `Grade/Aperture/Wavelength/Lighting` string (e.g. `1.0/17/660/45Q`)
   alongside individual fields. This is the standard's preferred
   summary form.

5. **Trade-dress / differentiation**: Standardized terms (SC, MOD, FPD,
   ANU, GNU, UEC, RM, PG) are vendor-neutral. Using them in our report
   instead of Cognex-specific labels reinforces "we conform to the
   standard, we don't clone the vendor's UI."

6. **DPM hand-off**: When `ApplicationStandard=AIM-DPM`, the verifier
   switches from 15415 to **ISO/IEC 29158** methodology. The DPM-only
   metrics in our 30-key enumeration (`cellDefects`, `finderPatternDefects`,
   `dataMatrixCellWidth/Height`, `horizontalMarkMisplacement`,
   `verticalMarkMisplacement`) live in 29158, NOT 15415. **We need the
   29158 standard for the DPM column semantics.**

---

## 12. Quick-reference card

The 7 mandatory matrix parameters + their grade interpretation:

```
DECODE        — yes/no
SC            — Rmax − Rmin             → A: ≥70 %    F: <20 %
MOD           — module:symbol contrast  → A: ≥0.50    F: <0.20
FPD           — fixed pattern damage    → A: none     F: total
ANU           — X/Y aspect deviation    → A: ≤0.06    F: >0.20
GNU           — module grid deviation   → A: ≤0.38    F: >0.75
UEC           — EC capacity remaining   → A: ≥0.62    F: <0.20

OVERALL GRADE = MIN(all mandatory parameter grades)
```

---

## What's NOT here

- The mathematical formulae for each parameter (~10 pages each in the PDF
  — refer to the source when needed)
- The histogram thresholding algorithm pseudocode (Annex A — read PDF
  directly if implementing; we won't)
- Full Parameter Overlay calculation method (Annex E — same)
- 1D-specific methodology (lives in ISO/IEC 15416, not in our library yet)
- DPM-specific methodology (lives in ISO/IEC 29158, not in our library yet)
