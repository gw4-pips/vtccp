# DMST TruCheck — All Tab States, Post-Scan (DM Grade D)

**Device**: DM475-63530E-PIPS-Verif-Lab  
**Firmware**: 6.1.16_sr4  
**Scan timestamp**: Mon 25-May-2026 07:18:36 (817ms) AM  
**Unit Serial**: 1A1903PP010754  
**Operator**: GW4  
**Company**: Product Identification and Processing Systems, Inc.  
**Symbol**: DM GS1 16×36 (the recurring Grade D test symbol)  
**Push script active**: v1.33  

---

## Tab 1 — General Characteristic

**Screenshot**: `dmst-general-characteristic-tab.png`

| Field | Value | Notes |
|---|---|---|
| Matrix Size | 16x36 (Data: 14x34) | Overall incl. finder+QZ; data region excl. finder |
| Horizontal BWG | 11% | |
| Vertical BWG | 11% | |
| Encoded characters | 38 | Push XML says 33 (eaLen wrong — confirmed bug) |
| Total Codewords | 56 | 32 data + 24 ECC |
| Data Codewords | 32 | |
| Error Correction Budget | 24 | |
| Errors Corrected | 7 | |
| Error Capacity Used | 14 | = 7 × 2 ✓ (ErrorsCorrected × 2 confirmed again) |
| Error Correction Type | ECC 200 | |
| Image | Black on white | Label for ImagePolarity — confirmed consistent |
| Nominal X Dim | 20.3 mil | This scan (DM live); prior HTML sample (QR) was 12.6 mil |
| Pixels per Module | 15.97 | Slight rounding vs Report tab (15.86) — same value |
| Contrast Uniformity | 74 at module(7,28) | **NEW FORMAT**: value + location qualifier |
| MRD | 67% (73% - 6%) | **NEW FORMAT**: MRD% (Rl% - Rd%) — Rl=73%, Rd=6% |

### Contrast Uniformity format — new discovery

"74 at module(7,28)" — DMST displays the CU value with the location of the worst module
as a `(column, row)` or `(row, column)` coordinate. The push XML emits only the raw
numeric value `74` (the `contrastUniformity` field). VTCCP currently does not capture the
location — it could be added as a separate `ContrastUniformityModule` field if needed.

### MRD format — new discovery

"67% (73% - 6%)" expresses: `MRD% (MaxReflectance% - MinReflectance%)`.
The push XML emits MRD as a numeric value; Rl and Rd appear separately in the push XML
as `<MaxReflectance>` and `<MinReflectance>`. DMST is composing this display string.

---

## Tab 2 — Data Detail

**Screenshot**: `dmst-data-detail-tab.png`

### Left panel

**Data (displayed):**
```
]>
<RS>06<GS>18VD89536<GS>1P8902A<GS>S3122A02965<RS><EO>
```

**Unicode Data:** identical

**Data Format Check — GS1 Application Data Format: FAIL**
- `<F1> Required at beginning of data` | Data: `[` | Check: FAIL
- Error: "Application Header is Expected."

The `[` is the AIM symbology identifier (Data Matrix prefix). GS1 expects `]d1` or similar
structured FNC1 framing — this symbol has GS1 AIs (06, 18VD, 1P, S) but begins with `[`
not `]d1...FNC1`, so GS1 format check fails. ISO grade is unaffected (4. = A on all
structure grades). This is the `ApplicationPass=Fail (Data Format)` field in push XML.

**ASCII Values:**
```
091 041 062 030 048 054 029 049 056
086 068 056 057 053 051 054 029 049
080 056 057 048 050 065 083 051 049
050 050 065 048 050 057 054 053
030 004
```
37 bytes total. ASCII 091=[, 041=), 062=>, 030=<RS>, etc. These are the per-character
decimal ASCII codes. Matches Encoded characters = 38 in General Characteristic (37 visible
+ the implicit AIM prefix character = 38 total including the leading ] sequence).

**Codewords** (56 total = 32 data + 24 ECC):
```
092 042 063 031 136 030 148 087
*069 *219 *183 *055 030 050 081 219
132 *066 *030 *084 161 152 066 132
226 054 031 005 129 045 195 090
107 149 089 060 057 044 200 195
082 134 017 211 227 021 091 158
239 024 033 160 108 040 223 204
```
`*=Fixed by Error Correction`. The starred codewords (069, 219, 183, 055, 066, 030, 084)
= 7 corrected codewords. Matches Errors Corrected=7 and Error Capacity Used=14 ✓.

### Right panel — Modulation grid (Data Detail view)

The Data Detail tab also shows the modulation grid. Confirmed the red `0` values appear
in a diagonal stripe through the symbol — consistent with ANU (Axial Nonuniformity) failure.
The ANU measures how consistently modules are sized across the axes; a diagonal stripe of
low-modulation cells is the signature pattern.

---

## Tab 3 — Quality Detail

**Screenshot**: `dmst-quality-detail-tab.png`

### Verification Grade header

| Overall | Aperture | Wavelength | Lighting | Formal |
|---|---|---|---|---|
| 1.0 (D) | 16 | 660 | 45Q | 1.0/16/660/45Q |

**Formal Grade string format confirmed**: `{numericGrade}/{aperture}/{wavelength}/{lighting}`.
This IS parseable. Components:
- `1.0` — numeric floor grade
- `16` — aperture reference number (not millimeters)
- `660` — illumination wavelength (nm), 660nm = red LED ✓ for DM475V-LBL
- `45Q` — lighting geometry (45-degree quadrant = fixed optics on DM475V-LBL)

### ISO15415 Quality Parameters

**Important**: Quality Detail shows numeric grade and Pass/Fail ONLY. No letter grade column.
Letter grades only appear in the Report tab. See Tab 6 for the letter-grade comparison.

| # | Parameter | Secondary | % | Numeric | Pass/Fail |
|---|---|---|---|---|---|
| 1 | Unused Error Correction (UEC) | | 41.7% | 2.0 | PASS |
| 2 | Symbol Contrast (SC) | Rl/Rd (83/4) | 78% | 4.0 | PASS |
| 3a | Modulation (MOD) | | | 4.0 | PASS |
| 3b | Reflectance Margin (RM) | | | 2.0 | PASS |
| 4 | Axial Nonuniformity (ANU) | | 11.2% | **1.0** | **FAIL** |
| 5 | Grid Nonuniformity (GNU) | | 8.5% | 4.0 | PASS |
| 6 | Fixed Pattern Damage (FPD) | | | 4.0 | PASS |
| 7 | Left 'L' Side (LLS) | | | 4.0 | PASS |
| 8 | Bottom 'L' Side (BLS) | | | 4.0 | PASS |
| 9 | Left Quiet Zone (LQZ) | | | 4.0 | PASS |
| 10 | Bottom Quiet Zone (BQZ) | | | 4.0 | PASS |
| 11a | Upper Left Quiet Zone (ULQZ) | | | 4.0 | PASS |
| 11b | Upper Right Quiet Zone (URQZ) | | | 4.0 | PASS |
| 12a | Right Upper Quiet Zone (RUQZ) | | | 4.0 | PASS |
| 12b | Right Lower Quiet Zone (RLQZ) | | | 4.0 | PASS |
| 13a | Left Top Transition Ratio (LQTTR) | | 0% | 4.0 | PASS |
| 13b | Right Top Transition Ratio (RQTTR) | | 0% | 4.0 | PASS |
| 14a | Left Right Transition Ratio (LQRTR) | | 0% | 4.0 | PASS |
| 14b | Right Right Transition Ratio (RQRTR) | | 0% | 4.0 | PASS |
| 15a | Left Top Clock Track (LQTCT) | | | 4.0 | PASS |
| 15b | Right Top Clock Track (RQTCT) | | | 4.0 | PASS |
| 16a | Left Right Clock Track (LQRCT) | | | 4.0 | PASS |
| 16b | Right Right Clock Track (RQRCT) | | | 4.0 | PASS |
| 17 | Average Grade (AG) | | 4.0 | 4.0 | PASS |
| 18 | DECODE | | | 4.0 | PASS |

**ANU=1.0 is the grade floor** — pulls Overall from any higher value to 1.0 (D).

SC secondary field format: "Rl/Rd (83/4)" — max reflectance=83, min reflectance=4.

**GNU: 8.5%** here vs 0.0% in the prior QR HTML sample — different symbols, expected.

---

## Tab 4 — Advanced Detail

**Screenshot**: `dmst-advanced-detail-tab.png`

Full raw modulation values grid — every module in the symbol with its numeric score (0–100).

- **Row index range**: -1 to 16 (18 rows; row -1 is the finder pattern border)
- **Column index range**: -1 to 36 (38 columns; column -1 is the finder pattern border)
- **Legend**: A=50-100 (dark green), B=40-49 (light green), C=30-39 (yellow), D=20-29, F=0-19 (red)
- **Red `0` cells**: the error-corrected modules — appear in a diagonal stripe from top-left
  region toward mid-right, confirming the ANU pattern. These are the 7 corrected codewords.
- Majority of modules score 82-99 (solid A). The nonuniformity is localized.

This is identical data to the modulation array in `q.modulationArray` in the push XML.
The `-1` row/column indexing matches the `isBlack` / `raw` cell structure in our
`ModulationValuesData` model (the finder pattern border cells are included).

---

## Tab 5 — Histogram

**Screenshot**: `dmst-histogram-tab.png`

Black background. Two plots:

### Histogram of Symbol Region
Full-image grayscale distribution of all pixels within the symbol ROI. Shows a very
strong peak near 0 (dark pixels — the module area fill), then a long right tail through
mid-tones up to near 255 (the label background). The bimodal shape is characteristic of
a high-contrast black-on-white label scan.

### Histogram of Module Centers
Grayscale distribution of module-center pixels only (one sample per module).
- X-axis runs: 100 — 50 — 40 — 30 — 20 — Threshold — 20 — 30 — 40 — 50 — 100
  (mirrored, dark on left / light on right of threshold)
- Grade zones color-coded: A (green, >50), B, C (yellow), D, F (red, <20)
- **Pattern**: most dark modules cluster at the far left (>80, grade A), most light
  modules cluster at far right (>80, grade A). Very few modules in the B–F range.
  This is a well-formed symbol. The ANU failure is geometry (axis alignment), not
  a reflectance issue — confirmed by the clean bimodal histogram.

---

## Tab 6 — Report (5 screenshots)

### Report header (screenshot 1: `dmst-report-tab-1-header.png`)

**Report header block:**
- Cognex yellow logo (top-left)
- Title: "Verification Report"
- Unit Serial: **1A1903PP010754**
- Verified: Mon 25-May-2026 07:18:36(817ms) AM
- Firmware Version: **6.1.16_sr4**
- Last Calibrated: 5/20/2026 1:14:58 AM

**Report Summary:**

| Field | Value |
|---|---|
| Data | `]><RS>06<GS>18VD89536<GS>1P8902A<GS>S3122A02965<RS><EO>` |
| Symbology | DataMatrix |
| Device Name | DM475-63530E-PIPS-Verif-Lab |
| Company | Product Identification and Processing Systems, Inc. |
| Operator | **GW4** |

**Verification Grades table:**

| Standard | Grade | Aperture | Wavelength | Lighting | Formal Grade |
|---|---|---|---|---|---|
| ISO 15415:2011 | 1.0 (D) | 16 | 660 | 45Q | 1.0/16/660/45Q |
| Custom | Fail (Quality) | | | | |

**Standard cited as ISO 15415:2011** — this is the edition the operator has configured
the verifier to use, not a firmware-determined value. The grading standard edition is a
user-selectable setting in DMST TruCheck, not intrinsic to fw 6.1.16_sr4. VTCCP must
read and echo whichever edition string the device reports — do not hard-code "2011".

**Image section**: Symbol photograph visible in report (lower-left) — confirms that the
Report tab DOES include the captured image even though the Main tab image pane went gray.

**General Characteristics in report** (right of image):

| Field | Report value | General Characteristic tab value | Delta |
|---|---|---|---|
| Pixels per Module | 15.86 | 15.97 | −0.11 — minor rounding at different scroll snapshot |

All other fields identical to General Characteristic tab.

---

### Report — DataMatrix Quality Parameters (screenshot 2: `dmst-report-tab-2-quality-params.png`)

The Report tab has a FOUR-column quality parameter table: `% value | Numeric | Letter | Pass/Fail`

The Quality Detail tab has THREE columns: `% value | Numeric | Pass/Fail` (no letter column).

**Letter grade column confirmed — numeric-to-letter mapping:**

| Numeric | Letter |
|---|---|
| 4.0 | **A** |
| 3.0 | **B** (inferred — not present in this scan) |
| 2.0 | **C** |
| 1.0 | **D** |
| 0 | **F** |

| # | Parameter | % | Numeric | Letter | Pass/Fail |
|---|---|---|---|---|---|
| 1 | UEC | 41.7% | 2.0 | **C** | PASS |
| 2 | SC | 78% | 4.0 | **A** | PASS |
| 3a | MOD | | 4.0 | **A** | PASS |
| 3b | RM | | 2.0 | **C** | PASS |
| 4 | ANU | 11.2% | 1.0 | **D** | FAIL |
| 5 | GNU | 8.5% | 4.0 | **A** | PASS |
| 6–16 | (all) | | 4.0 | **A** | PASS |
| 17 | AG | 4.0 | 4.0 | **A** | PASS |
| 18 | DECODE | | 4.0 | **A** | PASS |

The push XML emits numeric grades. The letter is derived by DMST for display.
VTCCP derives letters from the `GradingResult.LetterGradeString` property which applies
the same numeric→letter mapping.

Note: UEC=2.0 and RM=2.0 are both Grade C (not D), but ANU=1.0 (Grade D) is the floor
that determines Overall = 1.0 (D).

---

### Report — Encodation Analysis (screenshots 3–4)

**Screenshots**: `dmst-report-tab-3-encodation.png`, `dmst-report-tab-4-encodation-codewords.png`

The complete decoded codeword-by-character analysis:

| Codeword | Mode | Result |
|---|---|---|
| 092 | ASCII | [ |
| 042 | ASCII | ) |
| 063 | ASCII | > |
| 031 | ASCII | `<RS>` |
| 136 | ASCII | 06 |
| 030 | ASCII | `<GS>` |
| 148 | ASCII | 18 |
| 087 | ASCII | V |
| 069 | ASCII | D |
| 219 | ASCII | 89 |
| 183 | ASCII | 53 |
| 055 | ASCII | 6 |
| 030 | ASCII | `<GS>` |
| 050 | ASCII | 1 |
| 081 | ASCII | P |
| 219 | ASCII | 89 |
| 132 | ASCII | 02 |
| 066 | ASCII | A |
| 030 | ASCII | `<GS>` |
| 084 | ASCII | S |
| 161 | ASCII | 31 |
| 152 | ASCII | 22 |
| 066 | ASCII | A |
| 132 | ASCII | 02 |
| 226 | ASCII | 96 |
| 054 | ASCII | 5 |
| 031 | ASCII | `<RS>` |
| 005 | ASCII | `<EO>` |
| 129 | ASCII | ASCII PAD |
| 045 | ASCII | ASCII PAD |
| 195 | ASCII | ASCII PAD |
| 090 | ASCII | ASCII PAD |
| 107,149,089,060,057,044,200,195,082,134,017,211,227,021,091,158,239,024,033,160,108,040,223,204 | ECC | (24 ECC codewords) |

**Encodation breakdown**: 28 data characters + 4 ASCII PAD codewords = 32 data codewords.
Then 24 ECC codewords. Total = 56 ✓.

**Encoded characters in report = 38** (counting from ] through <EO> as the GS1/AIM encoded
string length, including the leading `]>` = 2 chars, plus the 36 data characters = 38).
Push XML eaLen=33 is wrong for this scan as well.

**q.encodationAnalysisArray shape** confirmed: each element = `{codeword, mode, result}`.
The `mode` field = "ASCII" throughout (pure ASCII encodation, no C40/Text/Base256 here).

---

### Report — Modulation Grid (screenshots 5: `dmst-report-tab-5-modulation-grid.png`)

The Report tab embeds the full modulation grid at the bottom. Identical data to
Advanced Detail tab — the same `q.modulationArray` rendered as a heatmap.

Red `0` cells form a clear diagonal stripe through rows 5–11, columns 4–8 approximately.
This is the ANU failure signature: a band of low-modulation modules indicates that the
module size varies systematically along one scan axis (likely the Y-axis, since
horizontal/vertical BWG are both 11% which is high).

---

## Confirmed data cross-references for firmware-confirmed-facts.md

| Finding | Value | Source |
|---|---|---|
| ErrorCapacityUsed = ErrorsCorrected × 2 | 7 × 2 = 14 ✓ | General Characteristic + Codewords |
| ANU = 1.0 (D) → Overall = 1.0 (D) floor | ✓ | Quality Detail |
| Formal Grade string format | `numericGrade/aperture/wavelength/lighting` | Report header |
| Letter grade mapping: 4.0=A, 2.0=C, 1.0=D | ✓ | Report Quality Parameters |
| Quality Detail has no letter column | ✓ | Report tab has it; Quality Detail tab does not |
| CU format: "74 at module(7,28)" | value + location qualifier | General Characteristic |
| MRD format: "67% (73% - 6%)" | MRD% (Rl% - Rd%) | General Characteristic |
| SC format: "Rl/Rd (83/4)" | max=83, min=4 | Quality Detail |
| Unit Serial | 1A1903PP010754 | Report header |
| ISO edition in report header | ISO 15415:2011 — user-configured setting, not firmware-determined | Report header |
| Operator in report | GW4 | Report Summary |
| Report image: present even when Main tab is gray | ✓ | Report tab header |
| q.encodationAnalysisArray mode field | "ASCII" | Encodation Analysis |
| q.modulationArray grid in Report tab | identical to Advanced Detail | Report bottom |
