# TruCheck Verification Settings — Report Settings Tab

**Screenshot**: `TruCheck-Verification-Settings-ReportSettings-2026-05-18.png`
**Captured**: 2026-05-18
**Context**: DMST 26.1.0 connected to DM475-63530E-PIPS-Verif-Lab [10.10.10.7]

---

## What the screenshot shows

TruCheck Verification Settings dialog → **Report Settings** panel.

Controls visible:

| Control | State in screenshot | Notes |
|---|---|---|
| **Generate Report** | ☑ checked | Top-level enable/disable for DMST's built-in report |
| **Image** | ☑ checked | Include symbol image in report |
| **General Characteristics** | ☑ checked | Symbology, decode data, grading standard, pass/fail |
| **Quality Detail** | ☑ checked | Per-parameter grade table (SC, MOD, ANU, GNU, UEC, RM, FPD, AG, DEC) |
| **Data Detail** | ☑ checked | Matrix size, ECC, codeword counts, error correction used |
| **Encodation Analysis** | ☑ checked | Character encoding breakdown |
| **Modulation Values Table** | ☑ checked | Per-module modulation grid |
| **ASCII Values** | ☑ checked | Decoded payload as ASCII character table |
| **Codeword Values** | ☑ checked | Raw codeword byte table |

Other settings panels visible in the left nav:
- **Application Settings** — **now cataloged**: `TruCheck-Verification-Settings-ApplicationSettings-2026-05-18.md`
- Calibration Settings — not yet captured
- Trending Settings — not yet captured
- User Information
- Report Settings ← (active)
- Navigation

---

## Significance for VTCCP

### 1. Framework alignment — VCCS report sections

These 8 report sections define the information architecture of a
professional barcode verification report. They map almost directly to what
VTCCP's D1 report should cover:

| DMST Section | VTCCP equivalent | Notes |
|---|---|---|
| Image | Captured symbol image (or loaded-image thumbnail) | Plus OpticsCompliance banner |
| General Characteristics | Header block: symbology, AIM ID, decoded data, grade, pass/fail, standard | VCCS-distinctive layout |
| Quality Detail | ISO parameter table (SC/MOD/FPD/ANU/GNU/UEC/RM/PG grades + numerics) | Use ISO terms, not Cognex vendor labels |
| Data Detail | Data Detail block: matrix size, ECC type, codewords, error correction used | |
| Encodation Analysis | Encodation block | Lower priority for initial report |
| Modulation Values Table | Out of scope for VTCCP v1 | Device-internal rendering; not in push output |
| ASCII Values | Decoded payload display | Already available from DecodedData field |
| Codeword Values | Out of scope for VTCCP v1 | Not in push output; would require DMCC report pull |

### 2. Unknown: do these checkboxes affect push-script output?

**Not yet tested.** The user notes uncertainty about whether unchecking
a Report Settings section suppresses corresponding data in the push-script
XML output, or whether these controls only affect DMST's own graphical
report rendering (the PDF/print report, not the push channel).

**Hypothesis**: these controls most likely affect only the DMST GUI report,
not the push channel. The push script receives `r.*` values from the
verifier's internal measurement engine, which runs regardless of report
section visibility. But this is unconfirmed.

**B5 probe action** (add to session plan): toggle one section off (e.g.
uncheck "Quality Detail"), trigger a scan, confirm whether the corresponding
push-XML fields (UECPercent, MODGrade, etc.) still appear. One scan to
confirm or refute the hypothesis.

### 3. Application-standard checks — already in place in DMST

**Correction to original note**: these are NOT a future roadmap item —
they are already implemented and shipping in the current DMST/TruCheck
firmware. Evidence already present in the v1.24 push-XML captures:

- `<ApplicationStandard>Custom</ApplicationStandard>` — the selected standard
- `<ApplicationPass>Fail (Quality)</ApplicationPass>` — result of the check
  (DM cal card scan)
- `<ApplicationPass>Pass</ApplicationPass>` — result on the QR loaded-image scan

The **Application Settings** panel (visible in the left nav of the
TruCheck Verification Settings dialog, not yet captured in detail) is
where the user configures which application standard to check against
(GS1, MIL-STD-129/130, ISO 15434, Custom, etc.). A follow-up screenshot
of that panel is pending.

**Implication for VTCCP**:
- `<ApplicationStandard>` and `<ApplicationPass>` are already first-class
  push fields — the parser should wire them (B4 scope).
- The `r.validation.gs1` JS object exists (v1.24 confirmed) and is
  populated when the device runs a GS1 check. The all-undefined AI probe
  result means the API surface is not property-bag style, not that GS1
  checking is absent — a different access pattern may expose the check
  result.
- VTCCP's own GS1 syntax check (using `gs1-syntax-engine`) is an
  **independent, additional** check — the device's check is a pass/fail
  flag against the configured standard; VTCCP's check gives per-AI
  detail that DMST doesn't surface in the push channel.

---

## Related files

- `references/external-libraries/gs1-syntax-engine.md` — GS1 syntax engine catalog
- `vtccp/lib/gs1-syntax-engine/` — engine downloaded, v1.4.0
- `architecture/` — D1 report design work (pending)
