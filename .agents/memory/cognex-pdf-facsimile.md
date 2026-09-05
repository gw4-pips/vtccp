---
name: Cognex PDF facsimile
description: Evidence and implementation direction for recreating the native Cognex PDF from DMST HTML
---

The native Cognex PDF and DMST HTML use the same report structure and content.
The examined native PDF identifies its producer as HiQPdf 12.0, uses A4 pages
(595 × 842 points), and shrinks the full report content onto the first page.
The trailing blank second page is consistent with the HTML rule
`body { page-break-after: always; }`.

Printing the same HTML through Adobe produced Letter pages (612 × 792 points)
at a larger effective scale, splitting Image/General Characteristics onto page
one and Quality Parameters/Data Format Check onto page two.

**Why:** DMST is not expected to export both HTML and its own PDF for one scan.
The HTML is sufficient source evidence to create a close facsimile of the native
Cognex PDF; no second verifier scan or native PDF export is required.

**How to apply:** Preserve the native HTML, then render a separately named,
clearly derived Cognex-facsimile PDF using controlled A4 portrait settings,
print backgrounds, a wide virtual browser viewport, and calibrated shrink-to-fit.
Neutralize the trailing forced page break to avoid a blank page. Validate layout
against native samples for QR, DataMatrix, linear, multi-symbol, and unusually
long reports before treating pagination as production-stable.