---
name: GS1 RFID addendum boundary
description: Product and document-boundary rule for combining canonical GS1 reports with RFID evidence.
---

Canonical GS1 pages must be byte-for-byte independent of RFID presence. When paired RFID evidence exists, render the complete established VCCS VeriWedge report separately and append its PDF pages after the canonical GS1 document; do not recreate a shortened RFID table inside the GS1 layout.

**Why:** RFID is VCCS supplemental cross-validation, not part of the official GS1/ISO barcode assessment. Keeping independently rendered documents prevents RFID data or CSS from changing the assessment pages and preserves the stronger, self-identifying VeriWedge presentation.

**How to apply:** Capture reader/acquisition provenance at scan time, render unavailable values honestly, and merge only after both documents have been rendered in the selected A4 or Letter profile. Add final page numbering after the merge.