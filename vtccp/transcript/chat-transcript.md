---

**User [2026-08-18]:** [continue with approved PDF rework plan]

**Assistant [2026-08-18]:** Executed all 18 approved PDF rework changes in one commit (0d4df50). Seven files changed:

Data plumbing: DmstHtmlScraper scrapes "Verified:" from HTML <p> tags and captures the source filename; DmstHtmlReport carries HtmlVerifiedString + HtmlSourceFileName; DmstReportValidator.MergeAndValidate now sets WebscanSourcePath and HtmlVerifiedString on the merged VerificationRecord — fixing both long-standing nulls.

Generator (v1.5.6): section1Title checks DeviceName so DM475V gets "COGNEX DataMan TruCheck…"; REPORT_DATETIME uses raw Verified: string verbatim; HDR_DATETIME and HDR_SW tokens removed; AppSpecStackCell single-line; logo fallback VeriWedge.

Template: "RFID VeriWedge™ Pro" throughout. Header column-flex with header-copyright line (www.vccs.llc / ©). Header-meta flex 0.67. All 5 section borders 2px; meta separator 2px; meta-lbl bold/black/8pt. Symbology col 110pt, Application Standard col 100pt. "TruCheck Report Name" label. App v1.5.13.
