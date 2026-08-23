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

The archived GS1 Barcode Engine contains the authoritative zero-compression
inverse cases for UPC-E expansion in `src/c-lib/ean.c`: a sixth compressed digit
of 0–2 expands as `Nabf0000cdeC`, 3 as `Nabc00000deC`, 4 as
`Nabcd00000eC`, and 5–9 as `Nabcde0000fC`, where `N` is the number system and
`C` is the unchanged check digit. Its implementation is explicitly written
for number-system 0, so number-system 1 still needs standards/native-capture
verification before production use.

**Why:** UPC-E is zero-suppression, not a simple left-padding operation; the
expansion branch depends on the sixth compressed digit.

**How to apply:** Accept the 8-digit HRI as `NabcdefC`, expand to UPC-A,
validate that the expanded UPC-A check digit equals `C`, then normalize UPC-A
to GTIN-14 for RFID comparison. EAN-8 to EAN-13 is five leading zeroes.