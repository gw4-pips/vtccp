---
name: TRUCHECK.APPLICATION-STANDARD mapping
description: The DMCC reference integer-to-name mapping for APPLICATION-STANDARD is inverted vs actual firmware behaviour on DM475V fw 6.1.16_tc9.
---

## Rule
**4 = Custom, 5 = Auto** on DM475V fw 6.1.16_tc9. The DMCC reference has them swapped.

**Why:** Empirically confirmed 2026-08-05:
- Unit returned 5 while VCCS UI showed "Auto" → 5 = Auto
- User changed UI to "Custom", reran dump, unit returned 4 → 4 = Custom
- DMCC reference rev 26.1.0.27 claims 4=Auto, 5=Custom — directly contradicted.

**How to apply:** Use 4 to set Custom, 5 to set Auto. DmccCommand.cs corrected.
Bug report filed: vtccp/references/cognex-bug-reports/TRUCHECK-APPLICATION-STANDARD-mapping-error.md

## Confirmed mapping (fw 6.1.16_tc9)
| Integer | DMCC reference | Actual |
|---|---|---|
| 4 | Auto | **Custom** |
| 5 | Custom | **Auto** |
