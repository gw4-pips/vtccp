---

**User [2026-08-18]:** [continue with approved PDF rework plan]

**Assistant [2026-08-18]:** Executed all 18 approved PDF rework changes in one commit (0d4df50). Seven files changed:

Data plumbing: DmstHtmlScraper scrapes "Verified:" from HTML <p> tags and captures the source filename; DmstHtmlReport carries HtmlVerifiedString + HtmlSourceFileName; DmstReportValidator.MergeAndValidate now sets WebscanSourcePath and HtmlVerifiedString on the merged VerificationRecord — fixing both long-standing nulls.

Generator (v1.5.6): section1Title checks DeviceName so DM475V gets "COGNEX DataMan TruCheck…"; REPORT_DATETIME uses raw Verified: string verbatim; HDR_DATETIME and HDR_SW tokens removed; AppSpecStackCell single-line; logo fallback VeriWedge.

Template: "RFID VeriWedge™ Pro" throughout. Header column-flex with header-copyright line (www.vccs.llc / ©). Header-meta flex 0.67. All 5 section borders 2px; meta separator 2px; meta-lbl bold/black/8pt. Symbology col 110pt, Application Standard col 100pt. "TruCheck Report Name" label. App v1.5.13.

---

**User [2026-08-18]:** Corrections to approved PDF plan — grade values must be scraped verbatim from HTML (not reformatted); additional layout feedback on the generated PDF.

**Assistant [2026-08-18]:** Implemented all changes at eeecf0e — report v1.5.7, app v1.5.14.

Verbatim grade scraping (6 files): DmstHtmlScraper Step 4e locates "Formal Grade" in the flat cell list via FindLastIndex, reads 6 verbatim strings (Standard, Grade display, Aperture, Wavelength, Lighting, Formal Grade). Stored on DmstHtmlReport.Html* properties, propagated through DmstReportValidator, added to VerificationRecord. Generator BuildGradeRows() prefers Html* fields over push-XML values; GradeDisplay() remains as fallback for push-XML-only mode.

Template layout (vccs-report-template.html): Section ① heading sec-sub-hdr → sec-hdr (dark navy 10pt, matches RFID section height). BVG heading sec-sub-hdr → sec-hdr with margin-top:-2px (overlaps bottom 2px border of sum-table, joins BVG to top section). Content gap 6pt → 4pt. header-meta flex 0.67 → 0.25 (~1 more inch left shift of title block). Symbology col 110→100pt, Application Standard 100→90pt. S8 separator 2px → 1px navy. App-spec cell: en dash, justify-content:center. Product name: VCCS <em>RFID VeriWedge™ Pro</em> everywhere (h1, RFID section heading, footer, copyright, title, logo fallback).
