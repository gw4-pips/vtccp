# Possible Work Item: Cognex PDF Facsimile Investigation

## Status

Deferred investigation. This is not approved implementation work yet.

## Objective

Determine how Cognex DataMan Setup Tool converts its native TruCheck HTML report
into the Cognex PDF so RFID VeriWedge can generate a close, clearly identified
facsimile from the HTML produced for the same scan.

This would avoid requiring DataMan to export both HTML and PDF for each scan.
The native HTML remains the verifier evidence; VeriWedge would perform only the
derived PDF rendering.

## Current evidence

- The examined native Cognex PDF identifies **HiQPdf 12.0** as its producer.
- The native PDF uses **A4 portrait** pages (595 × 842 points).
- Its report structure, values, images, colors, and ordering match the DMST HTML.
- Cognex shrinks the report to fit the meaningful content on one A4 page.
- Adobe printing uses different paper, margins, and scaling, causing the same
  HTML to split across two populated Letter pages.
- The native blank trailing page is consistent with the HTML rule
  `body { page-break-after: always; }`.

## Proposed investigation order

1. **Observe local PDF generation with Windows Process Monitor**
   - Monitor DMST file, process, and registry activity while it creates a PDF.
   - Look for temporary HTML/PDF files, external converter processes,
     configuration reads, and loaded HiQPdf assemblies.
   - Determine whether conversion is in-process or delegated to another program.

2. **Inspect the DMST installation and configuration**
   - Locate the HiQPdf assemblies and related configuration.
   - Search for A4, margin, browser-width, zoom, fit, and pagination settings.
   - Do not modify or redistribute Cognex components.

3. **Capture device and loopback traffic during HTML and PDF output**
   - Compare otherwise equivalent scans configured for HTML versus PDF.
   - Determine whether PDF-specific requests, responses, or parameters cross
     the wire.
   - Expectation: the device supplies report HTML/data and DMST performs the
     HiQPdf conversion locally. A capture is intended to confirm or disprove
     that expectation.

4. **Calibrate the VeriWedge renderer**
   - Render preserved DMST HTML using controlled A4 portrait settings.
   - Enable print backgrounds and set a fixed virtual browser width.
   - Tune margins and shrink-to-fit scale against native Cognex examples.
   - Remove only the trailing forced page break from the rendering copy so the
     derived PDF does not inherit Cognex's blank final page.

## Validation set

Before treating pagination as production-stable, compare native and derived
reports for:

- QR Code
- DataMatrix
- Linear symbols
- Passing and failing verification results
- Long Data Format Check sections
- Multi-symbol reports

## Deliverable and provenance

The output must be labeled as a **VCCS-rendered facsimile derived from native
Cognex HTML**, not as an original Cognex-generated PDF. Native measurements and
report values remain unchanged; only the PDF rendering is produced by VCCS.

## Decision criteria

Proceed to implementation only if the investigation shows that:

1. The HTML contains all content needed for the target report.
2. A stable renderer configuration reproduces Cognex pagination and geometry
   closely across the validation set.
3. The resulting file can be distinguished clearly from the original native
   Cognex artifact.
