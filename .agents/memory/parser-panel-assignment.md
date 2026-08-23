---
name: Parser panel assignment
description: Multi-symbol parser panels must be assigned by the exact native report identity, not the record-level primary fields.
---

## Rule

In a multi-symbol Webscan event, the dual TruCheck/VeriWedge parser panel belongs to the native report whose symbology and decoded identity are explicitly selected for parser comparison. It must not be inferred from a record-level decoded value that may represent a different symbol.

**Why:** A three-symbol export can contain two same-family Data Matrix reports with different GTINs; record-level fields can point at the RFID-matching report while the user expects the designated primary parser evidence to remain with the first/native report.

**How to apply:** Keep shell generation, native DFC rendering, and parser-panel assignment separate. Verify the rendered PDF with each symbol’s identity beside its parser block.