---
name: Hybrid Report Generator
description: Self-contained HTML report merging Webscan TruCheck barcode grades + VCCS FlexWedge RFID validation. Fire-and-forget from SessionViewModel.
---

# Hybrid Report Generator

## What it does
`DeviceInterface/Reports/HybridReportGenerator.cs` — static class, two public methods:
- `Generate(VerificationRecord)` → returns complete HTML string
- `SaveAsync(record, outputDir, filenameOverride?, ct)` → writes `yyyy-MM-dd_HH-mm-ss_hybrid_report.html`

## Where it's called
`VtccpApp/ViewModels/SessionViewModel.cs` — fire-and-forget Task.Run after `_sessionMgr.AddRecord(record)`. Never blocks the scan loop. Output dir = `_sessionOutputDir` (same as Excel file) unless `AppSettings.HybridReportOutputDirectory` overrides it.

## Settings (ConfigEngine/Models/AppSettings.cs)
- `GenerateHybridReport` (bool, default **true**)
- `HybridReportOutputDirectory` (string?, null = session dir)

## HTML structure — exact Webscan TruCheck match
Same inline CSS (`border-style:solid`, `padding:0.025in`, black-header `background-color:black;color:white`), same section order:
1. Report header (logo placeholder, title, date/time, device/fw/serial, **RFID status badge**)
2. Report Summary table (Data, Symbology, optional job context rows)
3. Verification Grades table (Standard/Grade/Aperture/Wavelength/Lighting/Formal Grade)
4. **RFID Validation — VCCS FlexWedge™** section (Tag Detected, EPC Hex, GTIN-14, GCP Valid, Serial, Validation Result [colour-coded], Mismatch Detail, Scan Window)
5. Symbol Image (base64 data URI, ROI crop preferred over full-frame)
6. ISO 15415 / ISO 15416 Quality Parameters table
7. VCCS footer + Print/Save as PDF button (hidden during print)

## PDF export
`@media print { .no-print: display:none; @page { margin: 0.65in; } }` — operator opens HTML in browser, Ctrl+P → Save as PDF.

**Why:** No Handlebars engine exists in the VTCCP codebase; plain C# StringBuilder avoids any dependency. Webscan-matching CSS is lifted verbatim from html_stylesheet.xslt analysis.

**How to apply:** If the HTML structure needs to change, edit `AppendXxx` private methods in `HybridReportGenerator.cs`. The `BaseCss` const at the bottom holds all styles — Webscan base block first, VCCS additions second.
