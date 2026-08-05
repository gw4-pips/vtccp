---
name: TRUCHECK.APPLICATION-STANDARD mapping
description: The DMCC reference integer-to-name mapping for APPLICATION-STANDARD contradicts observed firmware behaviour on DM475V fw 6.1.16_tc9.
---

## Rule
Treat **5 = Auto** for DM475V fw 6.1.16_tc9. Value 4 meaning is currently unknown.

**Why:** DMCC reference rev 26.1.0.27 documents 4=Auto, 5=Custom. But a DM475V unit
factory-reset to fw 6.1.16_tc9 returns `GET TRUCHECK.APPLICATION-STANDARD` = 5, while
the VCCS TruCheck Verification Settings UI (which reads the same device) shows "Auto"
in the Application Standard dropdown. Observed UI behaviour overrides the reference.

**How to apply:** Do not use the reference mapping when writing or interpreting
TRUCHECK.APPLICATION-STANDARD values on this firmware. Use the UI label as ground truth.
DmccCommand.cs comment has been corrected accordingly.

## Evidence
- DM475V-DPM 866D76 factory-default dump: APPLICATION-STANDARD = 5
- DM475V-DPM 866D76 post-calibration dump: APPLICATION-STANDARD = 5
- VCCS TruCheck Verification Settings screenshot (2026-08-05): shows "Auto"
- DMCC reference (rev 26.1.0.27): claims 4=Auto, 5=Custom — contradicted by above
