# DM475V Reference Manual — Notes

**File**: `reference-manual-25.4.1.1.pdf` (7.9 MB, 2,527 lines extracted to .txt)
**Source**: Cognex, document version 25.4.1.1
**Device**: DataMan 475V Verifier
**Firmware in use**: 6.1.16_sr4

`.txt` extraction also filed for grep convenience.

---

## Sections we've referenced

| Section | Lines (in .txt) | Used for |
|---|---|---|
| Calibration + SBG limits | 1095-1213 | Optics-source architecture (V vs non-V, what fields suppress when uncalibrated) |
| Standards-Based Grading section | ~1315-1400 | Confirmed SBG is for non-V readers; V readers ARE ISO verifiers |
| Calibration card list (NIST + Cognex symbols) | 1188-1202 | DMV-CCC / DMV-DMCC / DMV-GS1CC / DMV-AICC identification |
| Trigger / Live-mode mechanics | 950-1020 | Understanding trigger flow + push-script firing |
| Scripting templates (CSV + HTML) | ~1330-1395 | Reference for our DmstPushScript_v1.js |
| FTP transfer (Image / Result / Report) | 1959-1980 | Image-out (push); image-in (pull) NOT documented |
| Guided Focus Feedback | 1256-1268 | Future remote calibration UX |

---

## What this manual does NOT cover

(Look in the Comms+Programming Guide instead.)

- DMCC command reference / syntax / full command list
- Push-script JavaScript API (`r.*` object, `r.metrics`, `r.symbology`)
- Result XML schema details
- Image-load / replay / playback capability (silent on this — needs Comms guide)
- SDK API surface

## Important corrections / clarifications

1. **SBG is for non-V readers.** The manual discusses SBG limits in the
   context of any reader; do not misread this as "SBG mode on a V reader."
   V readers (DM475V, DMV-8072V) are full ISO-compliant verifiers when
   calibrated and never operate in SBG mode. SBG is the mode that lets
   non-V production scanners (DM150/260/370/etc.) report grade-like numbers
   with the optics-uncharacterized caveat. See
   `architecture/optics-source-model.md` (pending) for the full model.

2. **Calibration suppresses Aperture / SC / MR when absent**, regardless of
   reader tier. This is the firmware-level field-suppression rule we mirror
   for our loaded-image flow.
