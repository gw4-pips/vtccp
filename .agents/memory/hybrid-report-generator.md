---
name: Hybrid Report generator
description: History and archive status of the hybrid HTML report feature; pivot to standalone PDF.
---

# Hybrid Report — Archived

The hybrid HTML report feature was abandoned on 2026-08-13 after multiple sessions
of structural drift caused by reconstructing someone else's HTML template.

## What was archived
- `vtccp/DeviceInterface/Reports/_archived/HybridReportGenerator.cs` — the full generator (v2.1)
- All wiring removed from `SessionViewModel` (field refs, Push-mode watcher startup, fire-and-forget block)
- `AppSettings.GenerateHybridReport` default set to `false`
- Static sample preserved at `dist/hybrid-report-sample-v5.html`

## Why abandoned
Structural fidelity to the Webscan HTML required running their renderer. Every session
the template drifted because it was reconstructed from memory rather than the live file.
The "verbatim template" fix (copying the file character-for-character) solved the immediate
problem but the maintenance burden and complexity were not justified.

## Replacement direction
Standalone PDF page appended to the Webscan PDF (leaving Webscan output untouched).
Key advantages: no Cognex friction, simpler build, focused VCCS-only content, survives
Webscan format changes. PDF merge = one PdfSharp call.

**Why:** Keeping Webscan's output untouched removes the "did VTCCP tamper with this?"
risk for Cognex and customer audits.

## Recommended PDF library
QuestPDF (MIT license, no native deps, excellent .NET layout API) for generation;
PdfSharp for the merge step (append VCCS page to Webscan PDF).
