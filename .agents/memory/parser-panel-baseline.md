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