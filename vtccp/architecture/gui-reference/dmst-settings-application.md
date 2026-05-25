# DMST TruCheck Verification Settings — Application Settings

**Panel**: TruCheck Verification Settings → Application Settings (left nav, first item)  
**Firmware observed**: 6.1.16_sr4 (DM475-63530E-PIPS-Verif-Lab)  
**Logged**: 2026-05-25  
**Screenshots**:
- `dmst-settings-application-aperture-dropdown.png` — main panel with Aperture Setting dropdown open
- `dmst-settings-acas-upper.png` — Advanced Custom Application Standard Settings expanded (upper half)
- `dmst-settings-acas-lower.png` — Advanced Custom Application Standard Settings expanded (lower half)

---

## Navigation structure

The Settings dialog has six left-nav sections:
1. **Application Settings** ← this document
2. Calibration Settings
3. Trending Settings
4. User Information
5. Report Settings
6. Navigation

---

## Grading Standard (top radio group)

| Option | Notes |
|---|---|
| **ISO 15415/6** (default, selected) | 2D matrix (ISO 15415) and 1D linear (ISO 15416) symbologies |
| **ISO 29158 (AIM-DPM)** | Direct Part Mark grading standard — switches the entire grading regime |

This is a top-level mode switch. It governs which grading parameters and parameter labels
appear throughout the UI and report. VTCCP must track which standard is active per session.

---

## Application Standard

### Select Standard dropdown

Observed value: **Custom**. Other options exist in the dropdown (not yet fully enumerated —
likely includes GS1, HIBCC, and standards-body presets that pre-populate fields below).

### Main parameters

| Field | Observed value | Notes |
|---|---|---|
| **Dot Peen** | ☐ unchecked | Enables DPM dot-peen mode when checked |
| **Min X Dimension (mils)** | 8 | Lower bound on X-dim for pass/fail |
| **Max X Dimension (mils)** | 30 | Upper bound on X-dim for pass/fail |
| **Overall Pass Grade** | 1.5 | Minimum numeric grade for overall pass; dropdown (likely 1.0–4.0 in 0.5 steps) |

### Data Format Check (radio group)

| Option | Notes |
|---|---|
| **None** (selected) | No application-level data format validation |
| GS1 | Validate against GS1 Application Identifier syntax |
| HIBCC | Health Industry Bar Code Council format check |
| ISO 15434 | Data Identifier format check |

These map directly to the `ApplicationStandard` / `ApplicationPass` / `ApplicationPassReason`
fields in push XML. VTCCP design rule (user-confirmed): `ApplicationPass` is informational
only — `OverallGrade` (ISO 15415/15416) is always the operative grade.

### Aperture Setting dropdown

Three options, confirmed from open-dropdown screenshot:

| Option | Notes |
|---|---|
| **User Set** | Operator manually specifies aperture size |
| **Auto 50%/80%** | Firmware selects aperture automatically, bounded at 50%/80% of X-dim |
| **Auto Aperture** (observed, highlighted) | Firmware fully auto-selects aperture |

The aperture (also called "aperture diameter" in ISO context) determines the sampling area
used during reflectance measurement. It is part of the formal grade notation:
`numericGrade/aperture/wavelength/lighting` (e.g., `1.0/16/660/45Q` — aperture=16).
DMCC key for this setting: not yet confirmed from A1 digest.

---

## Advanced Custom Application Standard Settings

DMTC label: "Advanced Custom Application Standard" (chevron toggle, collapsed by default).  
**VTCCP label**: **Advanced Custom Application Standard Settings** — use this name exactly.  
Do NOT call it "Advanced Custom Application Standard" without "Settings", and do NOT use
"Advanced Calibration" (reserved term with different meaning in Webscan TruCheck legacy).

When expanded, this section allows per-parameter pass/fail threshold overrides for both
1D and 2D metrics, independent of the grading standard selected above.
All fields default to **N/A** (no override — grade-standard thresholds apply).

### 1D Pass Thresholds (left column)

Full field inventory from both screenshots:

| Field | Observed | Notes |
|---|---|---|
| Edge | N/A | |
| Minimum Reflectance | N/A | |
| Symbol Contrast | N/A | |
| Minimum Edge Contrast | N/A | |
| Modulation | N/A | |
| Defect | N/A | |
| Decode | N/A | |
| Decodability | N/A | |
| Quiet Zone | N/A | |
| Unused Error Correction | N/A | (1D) |
| Bar Width Growth (%) | N/A | Text field, not dropdown |
| Inspection Zone Height (mils) | N/A | Text field, not dropdown |

### 2D Pass Thresholds (right column)

Full field inventory from both screenshots:

| Field | Observed | Notes |
|---|---|---|
| Unused Error Correction | N/A | (2D) |
| Symbol Contrast / Cell Contrast | N/A | Dual-label: Symbol Contrast for label stock; Cell Contrast for DPM |
| Modulation / Cell Modulation | N/A | Same dual-label convention |
| Reflectance Margin | N/A | |
| Axial Nonuniformity | N/A | |
| Grid Nonuniformity | N/A | |
| Fixed Pattern Damage | N/A | |
| Average Grade / Distributed Damage Grade | N/A | Dual-label: AVG for label stock; DDG for DPM |
| Minimum Reflectance | N/A | (2D) |
| Matrix Size | N/A × N/A | Two text fields (rows × cols) |
| Bar Width Growth (%) | N/A | (2D) |
| Contrast Uniformity | N/A | |

### Grading Standard Versions

Three independent version dropdowns — one per symbology family:

| Standard family | Observed selection | Notes |
|---|---|---|
| **1D** | ISO 15416:2016 | Current edition (supersedes 2000 and earlier) |
| **2D** | ISO 15415:2011 | This is the user-configurable edition that appears in the report header. Do NOT hard-code "2011" in VTCCP — echo whatever the device reports. |
| **DPM** | ISO 29158:2020 | Current edition of the AIM-DPM / ISO 29158 standard |

These dropdowns directly determine the edition string in the ISO formal grade notation and
printed on the verification report. VTCCP must read and echo the active version per
symbology — it is NOT firmware-determined.

### QR Quiet Zone (radio group, 2D section)

| Option | Notes |
|---|---|
| **ISO 18004 QZ Requirement (4)** (selected) | Requires 4-module quiet zone on all sides — the stricter requirement |
| ISO 16480 QZ Requirement (1) | Requires 1-module quiet zone — the relaxed requirement for space-constrained applications |

This determines how the Quiet Zone grade parameter is evaluated for QR Code symbols.

### Clear Advanced Parameters button

Full-width button at the bottom of the ACAS section. Resets all 1D and 2D threshold
overrides back to N/A (reverts to grade-standard defaults). Does NOT affect Grading
Standard Versions or QR Quiet Zone selections.

---

## Bottom buttons

| Button | Notes |
|---|---|
| **Reset Defaults** | Resets entire Application Settings panel to factory defaults |
| **OK** | Commits all settings and closes the dialog |

---

## DMCC key mapping status

| Field | DMCC key | Status |
|---|---|---|
| Grading Standard (15415/6 vs 29158) | Unknown | Not yet confirmed — check A1 digest |
| Select Standard | Unknown | Likely a preset/template key |
| Dot Peen | Unknown | Likely TRUCHECK.DOT-PEEN or similar |
| Min/Max X Dimension | Unknown | |
| Overall Pass Grade | Unknown | |
| Data Format Check | `UPC-EAN.SUPPLEMENT` pattern? | Likely separate key — check A1 |
| Aperture Setting | Unknown | Aperture value appears in formal grade string |
| ACAS threshold fields | Unknown | One DMCC key per threshold field, likely |
| Grading Standard Versions (1D/2D/DPM) | Unknown | These drive the edition string in reports |
| QR Quiet Zone | Unknown | |

All DMCC key mapping for this panel is pending A1 digest review.  
**Do not implement this panel until DMCC keys are confirmed.**

---

## Open questions

| # | Question | Status |
|---|---|---|
| OQ-1 | What are all the options in the "Select Standard" dropdown (beyond "Custom")? | Not yet captured — needs screenshot with dropdown open |
| OQ-2 | What is the full range of values in the "Overall Pass Grade" dropdown? | Likely 1.0–4.0 in 0.5 steps; unconfirmed |
| OQ-3 | What DMCC commands read/write the Aperture Setting? | Pending A1 digest review |
| OQ-4 | What DMCC commands read/write the Grading Standard Version dropdowns (1D/2D/DPM)? | Pending A1 digest review — critical for report header edition string |
| OQ-5 | Does "ISO 29158 (AIM-DPM)" grading mode change which push XML fields are populated? | Likely yes — DPM-specific fields replace or supplement 2D fields |
