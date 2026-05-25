# DMST TruCheck Calibration Window

**Trigger**: Stylized C (©-like) icon at top-left of the DMST TruCheck toolbar  
**Firmware observed**: 6.1.16_sr4 (DM475-63530E-PIPS-Verif-Lab)  
**Logged**: 2026-05-25  
**Screenshots**:
- `dmst-trucheck-calibration-main.png` — dialog at open (camera off)
- `dmst-trucheck-calibration-custom-xdim.png` — Advanced panel expanded
- `dmst-trucheck-calibration-golive.png` — after Go Live (live feed + targeting overlays)
- `dmst-trucheck-calibration-inprogress.png` — calibration in progress (progress bar)
- `dmst-trucheck-calibration-status-dialog.png` — Calibration Status sub-dialog (incomplete result)

---

## Overview

The TruCheck Calibration dialog performs **two distinct calibration procedures** simultaneously:

| Component | What it calibrates | Requires |
|---|---|---|
| **Rmin / Rmax** | Reflectance references (specular black / specular white) | Certified Rmin/Rmax values from the physical calibration card's label, transcribed by the operator |
| **Pixmil** | Pixel-to-mil conversion (X-dimension scale factor) | A recognized Cognex/Webscan certified calibration target physically present in the field of view |

**Rmin/Rmax source**: The values are not operator estimates — they are certified reference
values printed on the NIST-traceable conformance test card or Cognex calibration card, given
to one decimal place precision (e.g., 88.x / 5.x). The operator reads the values off the
physical card and transcribes them into the RMax/RMin fields. The firmware accepts them as integers (the one-decimal precision on the card label is not
practically significant) and treats them as canonical until the operator explicitly enters
new values — they are sticky across sessions until changed.

Calibration is **complete** only when both components succeed. When the target in the field
of view is not a recognized certified calibration card, Rmin/Rmax are committed from the
operator-entered values but Pixmil is not updated — the dialog reports
**"Calibration incomplete (Rmin/Rmax updated, Pixmil is uncalibrated)"**.

`FieldCalibrated` in the push XML (`<FieldCalibrated>`) reflects calibration state.
All observed scans to date return `false` — meaning the device uses factory calibration.
Whether a partial calibration (Rmin/Rmax only) sets `FieldCalibrated=true` is unconfirmed.

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

## Go Live state

**Screenshot**: `dmst-trucheck-calibration-golive.png`

After clicking "Go Live", the camera activates and the live feed fills the preview area.
Two visual overlays appear:

| Overlay | Description |
|---|---|
| **Red dashed rectangle** | Targeting ROI — the region the firmware expects the calibration symbol to occupy. The symbol must be inside this box before firing "Start Calibration". |
| **Red crosshair (+)** | Frame center indicator. Its position relative to the symbol shows how much centering adjustment is needed. |

The screenshot shows the calibration card symbol (DataMatrix on white stock) partially
overlapping the ROI but not yet centered — the crosshair is below and left of the symbol.
Operator physically adjusts the card until the symbol is centered within the dashed box,
then clicks "Start Calibration".

---

## In-progress state

**Screenshot**: `dmst-trucheck-calibration-inprogress.png`

After clicking "Start Calibration":

- "Calibrating..." text label appears below the live view
- An **orange/red progress bar** fills left-to-right beneath it
- Go Live and Start Calibration buttons are disabled during calibration
- The live feed remains visible (symbol now better centered in frame)

Duration is short — a few seconds at most. The progress bar is the only feedback.

---

## Calibration Status sub-dialog

**Screenshot**: `dmst-trucheck-calibration-status-dialog.png`

Upon completion (success or partial), a modal **Calibration Status** sub-dialog appears
overlaid on the main calibration window. The progress bar below is fully filled (orange).

### Observed result — incomplete calibration (unrecognized target)

> **Calibration incomplete (Rmin/Rmax updated, Pixmil is uncalibrated)**
>
> Save calibration to non-volatile memory?  [ Yes ]  [ No ]

The blue **?** icon (not a red ✗) indicates a warning state rather than a hard failure —
the calibration partially succeeded.

**What happened**: The symbol in the FOV was a GS1 DataMatrix label (the recurring test
symbol), not a Cognex/Webscan certified calibration card. The firmware:
- Committed Rmin/Rmax from the certified values transcribed by the operator (88 / 5) — **succeeded**
- Could not determine Pixmil because the target was unrecognized — **not updated**

**Non-volatile memory prompt**: The "Save calibration to non-volatile memory?" question
appears regardless of whether calibration was complete or partial. Choices:

| Button | Effect |
|---|---|
| **Yes** | Writes current Rmin/Rmax (and Pixmil if updated) to device NVM — persists across power cycles and DMST restarts |
| **No** | Calibration values remain in volatile RAM only — lost when DMST is closed or device reboots |

In this session the operator chose **No** to preserve the existing factory calibration —
the partial values were not committed to NVM.

### Confirmed partial calibration sequence (DMCC implications)

The two-component model is now confirmed from firmware behavior:

```
Start Calibration →
  DMCC: write Rmin=5, Rmax=88   → always executes (user-entered values)
  DMCC: identify calibration target in FOV → requires certified card
    if recognized: compute Pixmil from target dimensions + known X-dim → write Pixmil
    if unrecognized: skip Pixmil update → "Pixmil is uncalibrated"
  → Calibration Status dialog: complete or incomplete
  → Prompt: save to NVM? Yes/No
```

**`FieldCalibrated` flag behavior** (unconfirmed — needs a recognized-target run):
- Almost certainly requires BOTH Rmin/Rmax AND Pixmil to be updated before returning `true`
- A partial update (Rmin/Rmax only) likely leaves `FieldCalibrated=false`

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
| OQ-1 | When Custom X Dimension is used, what Rmax/Rmin values does the operator enter — values from the custom target's own cert label, or the same NIST card values regardless? The window is silent; the "incomplete" result confirms Rmin/Rmax ARE written from user-entered values even without a recognized target, but the intended source of those values in the custom-target case is unclear. | **Pending** — clarification needed from Webscan founder |
| OQ-2 | What DMCC command(s) does DMST issue to the device when "Start Calibration" is clicked? Likely separate SET commands for Rmin, Rmax, and Pixmil. | Not yet confirmed — check A1 digest |
| OQ-3 | Does a partial calibration (Rmin/Rmax updated, Pixmil uncalibrated) result in `FieldCalibrated=true` in the next push XML? Hypothesis: no — requires both components. | **Unconfirmed** — needs a recognized-target calibration run followed by a scan |
| OQ-4 | Final VTCCP name for the "Custom X Dimension" sub-mode | Pending decision — do not use "Advanced Calibration" |
| OQ-5 | What DMCC command reads back the current calibration state → feeds `FieldCalibrated`? | Not yet confirmed — `FieldCalibrated` seen as false on all scans to date |
