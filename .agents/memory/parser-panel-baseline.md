---
name: Parser-panel baseline
description: User-confirmed visual baseline for the VCCS PDF parser panel.
---

# Parser-panel baseline

Treat the v1.5.44 VCCS PDF parser-panel layout as the canonical visual
baseline. The user rejected the later parser-panel revision and directed a
return to this layout.

**Why:** Parser-panel geometry and labels must be validated against an actual
generated PDF, not inferred from CSS or source-level column math.

**How to apply:** Before changing the parser panel, render a representative
report and compare it visually with the v1.5.44 baseline. Preserve the
baseline layout unless the user explicitly approves a new rendered result.

**Approved exceptions:** The VeriWedge parser header and Field-column text are
flush left inside their parser half. The user also approved giving the
VeriWedge parser a 57% share of the dual-panel DFC area, shifting the center
divider about 0.5in left to protect long BSE descriptions and element strings.

**Why:** The current BSE output needs more horizontal room than the native
verifier half; visually verify this exception in a rendered Letter PDF before
altering the split again.

**Structural rule:** For GS1-applicable scans, retain both TruCheck and
VeriWedge parser blocks regardless of whether TruCheck passes. The RFID result
label is independent: use Cross-Validation only when TruCheck cannot confirm
or fails and VeriWedge is needed as the fallback/comparison path.

**Why:** A visible dual-parser panel is the accepted report structure, not
evidence that the RFID result itself must be called Cross-Validation.

**How to apply:** Never gate the right-hand parser block on the RFID result
label or the native TruCheck pass state.