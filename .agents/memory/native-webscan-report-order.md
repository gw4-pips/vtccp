---
name: Native Webscan report order
description: Required ordering policy for multi-symbol Webscan report presentation.
---

## Rule

When a Webscan event contains multiple native symbol reports, preserve the
verifier's source sequence everywhere it is presented: barcode summary, grade
table, RFID result rows, and each barcode image/DFC/parser panel. Do not
reorder by symbology family or restrict the report to a fixed symbol count.

**Why:** Operators compare the VCCS report directly with the Webscan output;
different sequence between panels makes it appear that evidence belongs to the
wrong symbol.

**How to apply:** Use the ordered native-report collection as the single
projection for every per-symbol report section. Only use a compatibility
projection for older records that do not carry native report boundaries.