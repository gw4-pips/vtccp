# DMST TruCheck Calibration Window

**Trigger**: Stylized C (©-like) icon at top-left of the DMST TruCheck toolbar  
**Firmware observed**: 6.1.16_sr4 (DM475-63530E-PIPS-Verif-Lab)  
**Logged**: 2026-05-25  
**Screenshots**: `dmst-trucheck-calibration-main.png`, `dmst-trucheck-calibration-custom-xdim.png`

---

## Overview

The TruCheck Calibration dialog performs reflectance calibration of the DM475V using a
physical calibration card. Calibration establishes the Rmax (specular white reference) and
Rmin (specular black reference) reflectance values that ISO 15415/15416 grading depends on.

`FieldCalibrated` in the push XML (`<FieldCalibrated>`) reflects whether a calibration has
been performed in the current session. All observed scans to date return `false` — meaning the
device uses its factory calibration when this dialog has not been invoked.

---

## Standard Calibration (collapsed Advanced panel)

**Screenshot**: `dmst-trucheck-calibration-main.png`

Three-step workflow presented left-to-right:

| Step | UI element | Notes |
|---|---|---|
| 1. Enter Rmax and Rmin | Two numeric text fields labeled **RMax** / **RMin** | Default observed: RMax=88, RMin=5. These match values printed on the physical NIST-traceable calibration card shipped with the device. |
| 2. Click Go Live | **Go Live** button | Activates the live camera feed in the large preview area. The large rectangle above the controls is the live-view frame (shown blank/gray when camera is not live). |
| 3. Center symbol, then click Start Calibration | **Start Calibration** button | Operator centers the calibration card in the frame and fires calibration. |

**Calibration card R values**: The dialog says "Enter the Rmax and Rmin values *on the Calibration Card*."
This implies the values come from a printed/certified label on the NIST-traceable calibration target.
However, the window is silent about *which* calibration card is referenced when a custom
target is used in the Advanced panel (see below). Clarification pending from Webscan founder.

---

## Custom X Dimension (DMTC "Advanced Calibration" — expanded)

**Screenshot**: `dmst-trucheck-calibration-custom-xdim.png`

The "Advanced Calibration" label expands via a chevron toggle to reveal:

| Field | Notes |
|---|---|
| **Custom X Dimension (mils)** text field | User enters the known X-dimension (module pitch) of a non-NIST-traceable calibration target |
| **"Note: Calibration will not be NIST traceable."** | Displayed inline below the field |

### What this mode does

Allows calibration using any target whose X-dimension is known — even if it lacks a
NIST-traceable reflectance certificate. The resulting calibration is non-NIST-traceable.
The Rmax/Rmin fields (from step 1) are still entered; the question of which card's R values
are used when a custom X-dim target is substituted for the certified card is **unresolved** —
the window provides no guidance. Clarification pending.

### VTCCP naming — do NOT use "Advanced Calibration"

In legacy Webscan TruCheck software, "Advanced Calibration" referred to a procedure with
a completely different meaning and scope. VTCCP must not reuse the "Advanced Calibration"
label for this feature. Candidate names:

- **Custom Target Calibration**
- **Custom X-Dim Calibration**
- **Non-Traceable Calibration**

Final name pending decision. The key differentiator to surface in VTCCP UI is the
NIST-traceability distinction — standard = traceable (certified card), custom = non-traceable.

---

## DMCC backend

Calibration in DMST TruCheck is driven by DMCC — the same DMCC commands VTCCP uses.
The specific commands have not yet been confirmed from the A1 digest. When VTCCP
implements this feature, the relevant DMCC keys to identify are:

- Command to trigger calibration with Rmax/Rmin/XDim parameters
- Command to query current calibration state (feeds `FieldCalibrated`)
- Command to read back current Rmax/Rmin from the device

**Do not implement until**: (a) DMCC key mapping is confirmed from A1 digest, and
(b) scope decision on Custom Target mode is made after Webscan founder clarification.

---

## Open questions

| # | Question | Status |
|---|---|---|
| OQ-1 | When Custom X Dimension is used, what calibration card do the Rmax/Rmin values refer to — the same NIST card, or the custom target? The window is silent. | **Pending** — clarification needed from Webscan founder |
| OQ-2 | What DMCC command(s) does DMST issue to the device when "Start Calibration" is clicked? | Not yet confirmed — check A1 digest |
| OQ-3 | What DMCC command reads back the current calibration state → feeds `FieldCalibrated`? | Not yet confirmed — `FieldCalibrated` seen as false on all scans |
| OQ-4 | Final VTCCP name for the "Custom X Dimension" sub-mode | Pending decision — do not use "Advanced Calibration" |
