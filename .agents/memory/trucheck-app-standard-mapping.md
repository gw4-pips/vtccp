---
name: TRUCHECK.APPLICATION-STANDARD mapping
description: The DMCC reference integer-to-name mapping for APPLICATION-STANDARD is inverted vs actual firmware behaviour on DM475V fw 6.1.16_tc9.
---

## Rule
**4 = Custom, 5 = Auto** on DM475V fw 6.1.16_tc9. The DMCC reference has them swapped.

**Why:** Empirically confirmed 2026-08-05:
- Unit returned 5 while VCCS UI showed "Auto" → 5 = Auto
- User changed UI to "Custom", reran dump, unit returned 4 → 4 = Custom
- DMCC reference rev 26.1.0.27 claims 4=Auto, 5=Custom — directly contradicted by test.

**How to apply:** Use 4 to set Custom, 5 to set Auto in SET commands.
Ignore the DMCC reference for this parameter. DmccCommand.cs comment corrected.

## Confirmed mapping (fw 6.1.16_tc9)
| Integer | DMCC reference | Actual |
|---|---|---|
| 0 | GS1 | GS1 (assumed correct) |
| 1 | HIBCC | HIBCC (assumed correct) |
| 2 | UDI (HIBCC+GS1) | UDI (assumed correct) |
| 3 | UID | UID (assumed correct) |
| 4 | Auto | **Custom** |
| 5 | Custom | **Auto** |
| 6 | Cryptocode | Cryptocode (assumed correct) |
