---
name: TruCheck RFID validation policy
description: When a GS1 DataMatrix pass should suppress redundant RFID cross-validation and its report section.
---

# TruCheck RFID validation policy

Do not open an RFID cross-validation window or show an RFID validation section
when a GS1 DataMatrix already has a passing TruCheck application, overall-grade,
or correlated Data Format Check result.

**Why:** The operator considers the passing TruCheck GS1 validation authoritative
for this workflow. A second pass may be technically possible, but it adds no
operator-facing value and unnecessarily lengthens the scan cycle and report.

**How to apply:** Keep RFID validation active for failed GS1 DataMatrix results
and non-GS1 symbologies. When a pass suppresses validation, remove its result
section and header badge rather than labelling it as skipped or unavailable.