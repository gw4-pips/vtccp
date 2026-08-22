---
name: RFID cross-validation semantics
description: Meaning of RFID cross-validation relative to the TruCheck and VeriWedge GS1 parsers.
---

# RFID cross-validation semantics

“RFID cross-validation” is the GS1 parser comparison/fallback path: use the
VeriWedge GS1 parser when the TruCheck parser cannot validate the relevant RFID
data or its validation fails. It is not a generic RFID step to omit merely
because the barcode received a passing TruCheck grade.

**Why:** A TruCheck barcode pass and a parser’s ability to validate RFID-related
GS1 data are separate concerns. Conflating them removes the very fallback that
cross-validation is intended to provide.

**How to apply:** Any future parser-selection logic must model the TruCheck
parser result and the VeriWedge fallback explicitly. Use the correlated native
TruCheck DFC's explicit pass/fail outcome, not a barcode grade or application
status. Delay fallback invocation until that native outcome is available, and
only call a report result “Cross-Validation” when the fallback actually ran.

**Confirmation:** The user reconfirmed this definition on 2026-08-21 before
the fallback behavior was implemented.
