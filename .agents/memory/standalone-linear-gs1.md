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