---
name: DMST embedded barcode evidence
description: Safe rules for selecting barcode images embedded in DMST HTML reports.
---

**Rule:** Never use Base64 payload text as image context. Treat DMST's `Image` / `General Characteristics` table as affirmative capture evidence, while retaining explicit branding and ambiguity rejection.

**Why:** Arbitrary bytes encoded as Base64 can coincidentally contain words such as `logo` or `QR`, making a real barcode capture look like branding or a different symbology.

**How to apply:** When refining HTML image selection, inspect only the image metadata and surrounding markup after embedded payloads are removed. Preserve the source HTML as the sole provenance for the selected image.