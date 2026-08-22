---
name: Webscan QR exports without AG
description: How to import valid Webscan QR HTML exports that omit the Average Grade quality row without inventing verifier data.
---

Webscan QR exports may omit the `Average Grade (AG)` quality row while still
publishing a literal report-level `Verification Grades` result. When that
overall-grade row and the required available quality evidence are present, AG is
optional: accept the export, preserve the reported overall grade, and leave AG
fields blank.

**Why:** Rejecting the whole export solely for the absent AG row loses real
verifier evidence. Filling AG from the overall grade would violate the rule
that absent verifier data stays absent.

**How to apply:** Keep the report-level `Verification Grades` row, UEC, Symbol
Contrast, and DECODE rows as structural requirements. Do not add or infer AG.
Independently preserve any Webscan-native Data Format Check outcome, including
FAIL, exactly as the source report states it.