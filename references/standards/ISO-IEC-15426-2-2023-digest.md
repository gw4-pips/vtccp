# ISO/IEC 15426-2:2023 — Digest

**Full title**: Information technology — Automatic identification and data
capture techniques — Bar code verifier conformance specification — Part 2:
Two-dimensional symbols

**Edition**: Third (2023-12), cancels and replaces 2nd ed (2015)

**Source file**: `ISO-IEC-15426-2-2023-verifier-conformance-2D.docx`

**Scope**: Defines test methods and minimum accuracy criteria for **verifiers**
that grade 2D symbols (matrix + multi-row) using the ISO/IEC 15415
methodology. **This is the conformance standard for the device, not the
print-quality standard for the symbol** (that's 15415). VTCCP consumes
verifier output; this standard tells us what tolerances and reportable values
to expect *from* a conforming verifier.

---

## What this standard governs

A verifier is "conforming" if it:

1. **Performs all mandatory functions** (per Section 6.3, see below).
2. **Reports values within the tolerances of Table 1** when measuring primary
   reference test symbols.
3. **Documents** its operational parameters (Section 10).

The arithmetic mean of:
- **10 measurements** for multi-row symbols (PDF417, etc.)
- **5 measurements** for matrix symbols (Data Matrix, QR, Aztec, etc.)

must fall within the Table 1 tolerances.

---

## Table 1 — Tolerances for measured parameter values (the bible)

| Parameter | Symbology type | Tolerance |
|---|---|---|
| Rmax and/or Rs | Both | **±5 % reflectance** |
| Rmin and/or Rb | Both | **±3 % reflectance** |
| Unused Error Correction (UEC) | Both | **±0** (no tolerance — derived from exact codeword count) |
| Decodability | Multi-row | ±0.08 |
| Defects | Multi-row | ±0.08 |
| Codeword yield | Multi-row | ±0.08 |
| Grid Nonuniformity | Matrix | ±0.06 |
| Axial Nonuniformity | Matrix | ±0.02 |
| Contrast Uniformity (Modulation) | Matrix | ±0.08 (for the contrast uniformity value per A.3.2) |
| Fixed Pattern Damage | Matrix | Within calibrated grade boundaries |

**Implication for VTCCP**: When loaded-image flow is implemented and we report
"this would have been an X grade", we should also disclose that loaded-image
flow is *outside the conformance tolerance regime* — there's no Table 1
tolerance for that scenario, because the optics aren't the verifier's own.
This is the formal grounding for our `OpticsCompliant` flag.

---

## Section 6.3 — Mandatory functions

### 6.3.1 — Multi-row bar code verifiers (PDF417 etc.)

Must be capable of:
- Collecting reflectance from points along multiple scan paths
- Establishing scan reflectance profiles
- Analysing the scan reflectance profiles
- Reporting individual scan reflectance profile **parameter grades** AND
  **profile grades**
- Reporting **codeword yield** value + grade
- Reporting **unused error correction** value + grade
- Determining + reporting an **overall symbol grade**
- Reporting the **decoded data**

### 6.3.2 — Matrix symbol verifiers (Data Matrix, QR, Aztec)

Must be capable of:
- Collecting reflectance from a sample area framing the symbol + quiet zones
  (extends **20X beyond quiet zones** for certain symbol sizes)
- Establishing reference greyscale image + binarised image (per 15415)
- Decoding per the applicable reference decode algorithm
- Reporting **individual values + grades** for each 15415 parameter
- Reporting an **overall symbol grade**
- Reporting the **decoded data**

**Method of reporting is not specified** — display, printed report, or
electronic communication all qualify. Our `<DMCCResponse>` over TCP is a
valid reporting method.

---

## Section 6.4 — Optional (not required for conformance) functions

These are permitted enhancements:
- Reporting number of scan reflectance profiles / images the overall grade
  is based on
- Reporting symbology verified
- Reporting all decoded symbol characters or codewords
- Print-out / display of all (or selected) scan profiles / images

**VTCCP relevance**: Almost everything in our 167-column schema beyond the
mandatory parameters falls under "optional functions" per this standard.
This is fine — optional doesn't mean non-standard, it means not required
for conformance.

---

## Section 7 — Operational requirements (specs the manufacturer must declare)

The standard doesn't fix specific values; it requires the manufacturer to
**state** them in documentation:

- Operating temperature range
- Storage temperature range
- Relative humidity range (condensing / non-condensing)
- Power supply parameters
- Ambient light conditions (the standard explicitly names: high-efficiency
  fluorescent, sodium vapour, mercury vapour, red neon, direct sunlight as
  "typical light sources which can cause problems")

For battery-powered units: must warn or stop when battery gets low.

---

## Section 8.3 — Primary reference test symbols (PRTS)

**Definition**: Bar code symbols intended for testing verifier accuracy,
manufactured to tolerances **10x better** than the verifiers being tested,
by methods traceable to national standards.

Each PRTS must be supplied with a statement of:
- Symbology used
- Data encoded
- Measurement aperture(s) or effective resolution
- Spectral characteristics (peak wavelength or colour temperature) of
  calibration light
- Overall symbol grade per 15416 or 15415
- Individual parameter grade + value (if the symbol emphasises a particular
  parameter)

PRTS shall be produced on materials with **negligible dimensional variation**
under 10-30 °C, 30-70% RH ambient conditions.

**Secondary test symbols** are permitted for routine QA — these are symbols
graded against a verifier that has itself been checked against PRTS. **They
do NOT support conformance declarations** but are fine for periodic
calibration checks by users.

**VTCCP relevance**: The DMV-CCC, DMV-DMCC, DMV-GS1CC, DMV-AICC calibration
cards (per DM475V manual Section 4.1.1.1, lines 1188-1202 of extracted .txt)
are Cognex's secondary test symbols, traceable to NIST. The June 2023
calibration date on the UPCE-44960 sample in our v1.23 scan is the "last
checked against PRTS" stamp.

---

## Section 8.4 — Test report (what verifier output should look like)

When a verifier is tested for conformance, the report must include:

- Symbology(ies) tested
- Overall symbol grade as measured + grade defined for the PRTS
- Values for individually measured parameters
- Confirmation that measured values are within Clause 5 tolerances
- Copies of reports output by the verifier under test (printed reports or
  screen prints)

**The output reports themselves are part of the conformance evidence trail**
— which means the *format* of verifier output reports is at least obliquely
specified-by-reference. This bears on our reverse-report architecture: a
report containing fewer fields than the underlying data set isn't conforming
to "report all individual parameter values".

---

## Section 10 — Equipment specification (mandatory documentation)

The manufacturer **should** specify:
- Which symbologies the verifier can verify (and which optional features)
- Measuring apertures / effective resolutions available
- Symbol size ranges (relative to X dimension ranges)
- Illumination source spec (peak wavelength, colour temperature)
- Means of reflectance calibration
- Means of reporting and recording results
- Optional functions (per 6.4)
- Ability to average results from repeated scans
- Interfacing capabilities
- Programming / configuration specifications

**VTCCP relevance**: When VTCCP logs verifier output, it should also log
which device + firmware + apertures + lighting were used. The DMST XML
already provides `<ApertureRef>17</ApertureRef>`, `<Wavelength>660</Wavelength>`,
`<Lighting>45Q</Lighting>` — all three are reporting items required by
this clause. Our existing schema captures them.

---

## Annex A (normative) — Primary reference test symbols

### A.2 Multi-row symbologies (PDF417 per 15438)

Two sets, X = 0.200 mm and X = 0.500 mm, both at grades 4 and 1 for each of:
- Symbol contrast
- Modulation
- Defects (spots + voids)
- Decodability (edge-to-similar-edge)
- Unused Error Correction
- Codeword Yield

### Table A.1 — Multi-row parameter values

| Parameter | Grade 4 | Grade 1 |
|---|---|---|
| Symbol contrast | ≥73.75 % | 25-35 % |
| Defects | ≤0.1375 | 0.2625-0.2875 |
| Decodability | 0.65 | 0.28-0.34 |
| Codeword yield | ≥72.75 % | 51.75-55.25 % |
| Unused error correction | ≥0.65 | 0.28-0.34 |

### A.3 Matrix symbologies (Data Matrix per 16022)

Two sets, X = 0.200 mm and X = 0.500 mm, grades 4 and 1 for:
- Symbol contrast
- Grid Nonuniformity
- Axial Nonuniformity

Plus separate symbols for:
- **A.3.2** — Modulation/contrast uniformity (single "dark widow" symbol
  encoding `///00`, dark module shrunk to 5/9 X dimension, measured with
  0.8% effective aperture)
- **A.3.3** — UEC (specific Figure A.2 symbol with 8 introduced errors,
  yielding UEC = 0.4286 → grade 2.4)
- **A.3.4** — Fixed Pattern Damage (per-symbology test symbols for Data
  Matrix Figure A.3, QR Figure A.4, Aztec Figure A.5)

### Table A.5 — Matrix parameter values

| Parameter | Grade 4 | Grade 2 | Grade 1 |
|---|---|---|---|
| Symbol contrast | ≥73.75 % | — | 25-35 % |
| Grid non-uniformity | ≤0.35 | — | 0.66-0.72 |
| Axial non-uniformity | ≤0.055 | — | 0.105-0.115 |
| Unused error correction | NA | 0.43 | NA |
| Fixed pattern damage | AG = 4 | — | AG = 2.6 |

`AG` = average grade for the five fixed pattern segments evaluated for Data
Matrix.

### New in 3rd edition

Per Foreword:
- Tolerances for certain parameters **clarified**
- **Fixed pattern damage test symbol for Aztec code added** (Figure A.5)

---

## Annex B (normative) — Verification requirements for primary reference test symbols

Primary verification = measurement by a device that mimics commercial
verifier methodology with **national-standard-traceable** performance.

- Linear bar code / multi-row: **high resolution scanning microdensitometer**
- Matrix symbols: **high resolution imaging system**

### Required precision of the primary verification device

- Linear distance measurement: chrome-on-glass linear encoder, laser
  interferometer, or equivalent, traceable to chrome-on-quartz national
  standard. Repeatability: ±0.5 μm worst-case for 39 elements over 5 scans.
- Imaging systems: **minimum 10 pixels per module per axis, preferably 20**.
  For Data Matrix X = 0.150 mm, that's ≥66 pixels/mm (1694 pixels/in),
  preferably 133 pixels/mm (3387 pixels/in).
- Reflectance: ≥10-bit ADC, traceable to national reflectance tile standard.
  Rmax/Rmin repeatability: ±0.5 % reflectance worst case over 5 scans.

**Explicit note**: These resolution and traceability requirements apply
**only to the primary verification device, NOT to commercial verifiers**.
Commercial verifiers can use less resolution as long as 15415 repeatability
is achieved.

---

## Why this matters for VTCCP

1. **Tolerance disclosure**: Our reports should be able to surface Table 1
   tolerances next to each measured parameter, so users can compare
   their grade boundaries to the conformance window.

2. **Conformance boundary**: When we add `OpticsSource = Loaded-Image`, we're
   explicitly outside the conformance regime of this standard. The
   `OpticsCompliant = false` flag should cite this standard by reference:
   "Loaded image — not measured under ISO/IEC 15426-2 conformance
   conditions."

3. **Reportable-field minimum**: Our 167-column schema is far broader than
   the standard's mandatory minimum. The mandatory minimum is a useful
   "lite report" mode for users who want the conformance-required subset
   only.

4. **Calibration card chain of trust**: The June 2023 calibration date on
   the UPCE-44960 sample we keep scanning is the user's evidence trail
   under Annex B's traceability requirements. VTCCP captures this in
   `<CustomNote>` field today; should probably promote to first-class
   `<CalibrationDate>` + `<CalibrationCardSerial>` columns.

5. **We are not building a verifier**. The whole standard is unambiguous
   that primary verification requires NIST-traceable instruments 10x better
   than the commercial verifier. VTCCP's contribution is in the data
   management + reporting layer downstream of a conforming verifier.

---

## Cross-references this standard depends on

| Standard | Provides |
|---|---|
| ISO/IEC 15415 | The 2D print-quality methodology this conformance spec validates against. **See `ISO-IEC-15415-ed3-2024-digest.md`** (sibling file). |
| ISO/IEC 15416 | Linear bar code print quality (referenced for multi-row parameters that derive from scan reflectance profiles) |
| ISO/IEC 15426-1 | Verifier conformance for linear symbols (sister standard) |
| ISO/IEC 15438 | PDF417 symbology spec — defines the multi-row test symbols |
| ISO/IEC 16022 | Data Matrix symbology spec |
| ISO/IEC 18004 | QR Code symbology spec |
| ISO/IEC 24778 | Aztec Code symbology spec |
| ISO/IEC 19762 | AIDC harmonised vocabulary (terms + definitions) |
| ISO 2859-1 | Sampling procedures (informative, for manufacturer's QA) |

---

## What's NOT here

- Specific test symbol artwork files / patterns (refer to figures in the
  actual PDF; we have the docx form so figures may not have rendered)
- The print-quality calculation algorithms themselves (those live in 15415,
  which this standard validates verifiers against)
- 1D-only conformance details (those live in 15426-1; not in our library yet)

---

## Editorial caveats

- This particular file is the WG-comments-incorporated version (filename:
  `Updatedbased_on_WW_comments`). It's "very close to but not the current
  published version" per project owner. The actually-published 2023
  version is what conforming verifiers are tested against; this draft
  may differ in minor wording but should be substantively identical for
  our reference purposes.
- DOCX format means most figures/tables render as text only — for the
  visual layouts of Figures A.1-A.5 (test symbols), refer to the
  published PDF when acquired.
