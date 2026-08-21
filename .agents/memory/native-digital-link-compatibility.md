---
name: Native Digital Link compatibility
description: How VCCS reports known native GS1 Digital Link parser limitations without conflating them with RFID outcomes.
---

## Rule

When the independent VeriWedge GS1 Digital Link parser passes while a native verifier DFC fails, keep the native result and RFID cross-validation as separate outcomes. For a known unsupported native version, present the native failure as `FAIL*` and attach an explanation that the named firmware or software does not support GS1 Digital Link parsing.

The documented boundaries are DM475V/DMTC firmware numeric version 6.1.16 (the latest released verifier-line artifact is `6.1.16_sr4`) and earlier, plus Webscan TruCheck software 3.03.74 and earlier. Pre-release firmware suffixes with the same numeric 6.1.16 boundary are treated consistently until Cognex publishes a later released verifier-line version.

**Why:** A native Digital Link parser limitation can otherwise look like either invalid GS1 data or an EPC-versus-barcode RFID mismatch. Those are distinct facts and must remain reviewable independently.

**How to apply:** Only add the compatibility annotation when the native DFC actually reports Fail and the VCCS Digital Link validation reports Valid. Do not rewrite verifier-sourced values or turn an actual RFID mismatch into a pass.