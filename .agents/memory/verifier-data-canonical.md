---
name: Verifier data is canonical — violations audit
description: The rule that verifier output must never be re-derived, plus a registry of known violations found in the codebase.
---

## The rule

Any value the verifier (DM TC or Webscan TruCheck) reports must be taken verbatim from its output (DM TC HTML or push XML). VTCCP must never recalculate, reconstruct, or substitute a locally-computed value. If verifier data is absent, the field must remain absent — not filled algorithmically. This rule is in replit.md.

## Known violations (audited 2026-08-18)

### HIGH — directly shapes what appears in the report as verifier data

**H1. BuildDataFormatCheck (DmstReportValidator.cs ~402-451)**
- What: extracts AI(01)/AI(21) from decoded push data, recomputes check digit, synthesizes PASS/FAIL and DFC rows for 2D GS1 symbols.
- Now used only as fallback when DM TC HTML DFC is absent (HTML scraping is now primary).
- Violation: even as a fallback, locally computed DFC replaces absent verifier data. Should be null when HTML DFC is absent.

**H2. BuildLinearDataFormatCheck (DmstReportValidator.cs ~453-547)**
- What: reconstructs EAN/UPC GTIN rows, recalculates check digits from HTML-scraped linear decoded data, invents N/A outcomes for UPC-E/unknown lengths. Report-facing via VccsHtmlReportGenerator lines 386-430.
- Violation: check digits and outcomes not taken from verifier; calculated locally.

**H3. ECC200 codeword derivation (DmstResultParser.cs ~465-475)**
- What: when push XML omits DataCodewords/ErrorCorrectionBudget, derives them from matrix size via ECC200 lookup table.
- Violation: these are verifier quality fields; absent values should stay absent.

**H4. Overall pass/fail inference (DmstResultParser.cs ~344-358)**
- What: derives overall scan PASS/FAIL from grade letter bands (A/B=PASS, F=FAIL) when verifier doesn't explicitly provide VerificationOverallPass.
- Violation: local letter-band logic can disagree with firmware for C/D grades; should use only VerificationOverallPass, or remain unknown.

### MEDIUM — influences report logic, one step removed

**M5. NumericToLetterGrade (DmstResultParser.cs ~197-210)**
- What: converts numeric grade to A–F using locally hard-coded ISO thresholds when verifier provides only a number.

**M6. Symbology reclassification (DmstResultParser.cs ~255-280; VerificationXmlMap.cs ~212-260)**
- What: reclassifies DataMatrix + ]d2 as GS1 DataMatrix locally. Interpretation, not verbatim verifier output.

**M7. BarcodeDataFormatter display transforms**
- What: substitutes control chars (FNC1 → <GS>, etc.) in the decoded data field used by the report.
- Lower risk if this is truly display-only and raw data is preserved separately.

**M8. Multi-mode extraction heuristic (DmstHtmlScraper.cs ~574-643)**
- What: heuristically searches HTML cells for grade data; hard-codes ISO/IEC 15416 as standard.

### LOW

**L9. GS1Parser.cs (ExcelEngine)**
- What: independently parses GS1 AIs from decoded data for batch/lot metadata in Excel. Not a quality/grade field.

## HTML-only VCCS PDF decision

The VCCS barcode-verification PDF is fail-closed: it can render only after a real
local DMST HTML artifact is correlated by the verifier's literal `Verified:` identity.
HTTP content, reader transport values, timestamps, filenames, and locally generated
values cannot qualify or fill a missing barcode-verifier field. If the HTML lacks a
field, the PDF must state that it is unavailable; it must not substitute another source.

**Why:** A barcode-verification PDF must never present uncorrelated or reconstructed
values as TruCheck data. A false-looking-complete report is worse than no report.

**How to apply:** Preserve an explicit filesystem-artifact provenance state through
acceptance, Excel, and rendering. Render only the dedicated HTML-backed field for
barcode summary, grades, Data Format Check, report time, application standard, and
image. Keep RFID-reader values visibly separate.

## Status

The strict HTML-only VCCS PDF path is remediated. The broader generic-record
violations above remain known and need a separate, explicitly discussed scope before
they are changed.
