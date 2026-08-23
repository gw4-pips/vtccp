---
name: Standalone linear GS1 verifier behavior
description: EAN/UPC scans use the normal GS1 parser/report policy; native DFC may show a non-14-digit GTIN and separate check digit.
---

EAN-13, UPC-A, EAN-8, and UPC-E are ordinary GS1 symbologies in the verifier
and should not receive a special parser-comparison policy. Preserve the
verifier's native GTIN and separately reported check digit exactly as supplied,
even when the GTIN is not 14 digits. Do not locally recompute the result.

**Why:** The user confirmed that both Webscan and DataMan TruCheck use the
normal two-parser behavior and that the GS1 Barcode Syntax Engine accepts the
GTIN-only payload.

**How to apply:** A standalone linear report may have one verifier symbol and
grade row because that is the evidence shape, but it must retain the normal
GS1 parser comparison behavior and native DFC handling.

For RFID comparison only, UPC-A is represented as GTIN-14 with two leading
zeroes and EAN-13 with one leading zero. Keep EAN-8, UPC-E, and 2/5-digit
add-ons unresolved until native verifier captures establish their behavior.

**Why:** Equal GS1 identities can have different native barcode and EPC
representations; treating the literal lengths as a mismatch gives a false
RFID failure.

**How to apply:** Normalize only the explicitly supported linear forms for
comparison; never replace the native verifier values shown in the report.