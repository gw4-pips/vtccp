# DMST TruCheck Verification Settings — Application Settings

**Document version**: v1.1
**Revised**: 2026-08-21
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

Observed value: **Custom**. The complete user-facing option list is:

| Option | Notes |
|---|---|
| GS1 | GS1 application standard |
| HIBCC | Health Industry Bar Code Council application standard |
| UDI (GS1 or HIBCC) | UDI application standard |
| UID (MIL-STD-130) | Unique Identification marking standard |
| Custom | Operator-configurable standard; current workstation selection |
| Auto | TruCheck automatically selects the applicable standard |
| Cryptocode | Cryptocode application standard |

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
| **Auto 50%** | Firmware selects the 50% aperture mode automatically |
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

Dropdown options confirmed from screenshots (2026-05-25):

| Standard family | Observed selection | All confirmed options |
|---|---|---|
| **1D** | ISO 15416:2016 | ISO 15416:2016 · ISO 15416:2025 |
| **2D** | ISO 15415:2011 | ISO 15415:2011 · ISO 15415:2024 |
| **DPM** | ISO 29158:2020 | **ISO 29158:2011** · ISO 29158:2020 · ISO 29158:2025 |

The DPM dropdown is the only one with three options (2011, 2020, 2025).
Do NOT hard-code any edition string in VTCCP — echo whatever the device reports.

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

Keys confirmed from `DmccCommand.cs` (A1 digest sourced, fw 6.1.10+):

| Field | DMCC key | Values | Status |
|---|---|---|---|
| Grading Standard (top-level) | `TRUCHECK.GRADING-STANDARD` | 0=ISO 15415/6, 1=ISO 29158:2020 | ✓ Confirmed |
| Select Standard (Application Standard) | `TRUCHECK.APPLICATION-STANDARD` | 0=GS1, 1=HIBCC, 2=UDI (GS1 or HIBCC), 3=UID (MIL-STD-130), 4=Custom, 5=Auto, 6=Cryptocode | ✓ Confirmed |
| Dot Peen | `TRUCHECK.DOT-PEEN` | ON / OFF | ✓ Confirmed |
| Min X Dimension | `TRUCHECK.APPLICATION-CUSTOM-MINIMUM-X-DIM` | [1–1000] thousandths of an inch | ✓ Confirmed |
| Max X Dimension | `TRUCHECK.APPLICATION-CUSTOM-MAXIMUM-X-DIM` | [1–1000] thousandths of an inch | ✓ Confirmed |
| Overall Pass Grade | `TRUCHECK.APPLICATION-CUSTOM-PASS-GRADE` | [0–40], no decimal | ✓ Confirmed |
| Data Format Check | `TRUCHECK.APPLICATION-CUSTOM-DATA-PARSING-STANDARD` | 0=None, 1=GS1, 2=HIBCC, 3=ISO 15434 (UI label) | ✓ Confirmed |
| Aperture Setting (mode) | `TRUCHECK.APERTURE` | 0=User Set, 1=Auto 50%, 2=Auto Aperture | ✓ Confirmed |
| Aperture Size (User Set only) | `TRUCHECK.APERTURE-SIZE` | [1–300] ten-thousandths of an inch | ✓ Confirmed |
| Grading Standard Versions (1D/2D/DPM) | Unknown — NOT yet in DmccCommand.cs | Separate per-family version keys required | **Pending** — search A1 digest |
| ACAS individual threshold fields | Unknown | One key per threshold, likely | **Pending** — search A1 digest |
| QR Quiet Zone | Unknown | | **Pending** — search A1 digest |

**Application Standard mapping correction**: the DMCC reference reverses the firmware values
for Custom and Auto. The active DM475V firmware returns `4` for the UI’s **Custom** selection
and `5` for **Auto**; use the verified device behaviour, not the reference’s reversed labels.

**Aperture Setting label**: the currently captured TruCheck UI uses **Auto 50%** for raw
`TRUCHECK.APERTURE = 1`; use that user-facing label in VCCS reports.

---

## Design notes for VTCCP Command Pilot

### Reset Defaults — granularity problem

DMST's "Reset Defaults" button on this panel resets **the entire TruCheck environment**,
not just the Advanced Custom Application Standard Settings section. This is coarse and
dangerous — a user adjusting only ACAS thresholds should not inadvertently wipe all
Application Standard, Aperture, and GS1-table settings.

**VTCCP must offer more granular reset options:**

| Reset scope | What it resets |
|---|---|
| Reset Advanced Parameters | ACAS threshold overrides only (equivalent to DMST's "Clear Advanced Parameters" button) |
| Reset Application Standard | Select Standard, Min/Max X-Dim, Pass Grade, Data Format, Aperture |
| Reset Grading Standard Versions | 1D/2D/DPM edition dropdowns only |
| Reset All Application Settings | Full equivalent of DMST Reset Defaults |

Each scope should be a distinct, clearly-labeled button or menu action with a confirmation
step before firing. Never a single undifferentiated "Reset Defaults."

### UI improvement directive

> "The Command Pilot UI needs to be a significant improvement over this DM TC UI."

The DMST Application Settings panel is a flat scrollable dialog with no grouping hierarchy,
no inline help, and no visual differentiation between primary settings and advanced overrides.
The ACAS section is a wall of 20+ N/A dropdowns with no context.

VTCCP Command Pilot design targets:
- Progressive disclosure: show only relevant fields for the selected Application Standard
- Inline help text per field (tooltip or sidebar panel) explaining the ISO context
- Visual grouping by standard family (1D / 2D / DPM) with family-level enable/disable
- ACAS thresholds: show only fields that deviate from N/A, collapsed by default
- Reset controls: per-section, with confirmation, with "what will change" preview
- Live validation: flag when min X-dim ≥ max X-dim, invalid pass grade, etc.

---

## Open questions

| # | Question | Status |
|---|---|---|
| OQ-1 | What are all the options in the "Select Standard" dropdown (beyond "Custom")? | Not yet captured — needs screenshot with dropdown open |
| OQ-2 | What is the full range of values in the "Overall Pass Grade" dropdown? | Likely 1.0–4.0 in 0.5 steps; unconfirmed |
| OQ-3 | What DMCC commands read/write the Aperture Setting? | Pending A1 digest review |
| OQ-4 | What DMCC commands read/write the Grading Standard Version dropdowns (1D/2D/DPM)? | Pending A1 digest review — critical for report header edition string |
| OQ-5 | Does "ISO 29158 (AIM-DPM)" grading mode change which push XML fields are populated? | Likely yes — DPM-specific fields replace or supplement 2D fields |
