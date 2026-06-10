# D1 Report Design — Architecture Notes

**Created**: 2026-06-10

---

## 1. Report rendering architecture (confirmed)

See `parser-strategy.md` for parser decisions. For rendering:

- VTCCP serializes `VerificationRecord` → VTCCP XML schema
- `XslCompiledTransform` (.NET built-in, no NuGet) + `vtccp-report.en.xslt` → HTML
- CSS file handles visual style (logo, colors, fonts) — swappable per customer/branding
- HTML → PDF via `WebBrowser.Print()` (zero-dependency) or `PuppeteerSharp` (better pagination)
- Webscan XSLT files (`attached_assets/report_stylesheet.en_*.xslt`, `html_stylesheet_*.xslt`)
  serve as **content specification reference** — section ordering, field labels, conditional
  blocks, formal grade string format. NOT used verbatim (schema differs).

---

## 2. Webscan XSLT files as content reference (confirmed 2026-06-10)

The Webscan XSLT stylesheets ARE what control the layout of the Webscan HTML/PDF reports.
Architecture: `VerificationReport XML → html_stylesheet.xslt → HTML`
              `VerificationReport XML → report_stylesheet.xslt → XSL-FO → PDF`

They are the authoritative reference for:
- Section ordering and grouping
- Field labels (all four language variants: EN/DE/ES/JA)
- Conditional rendering logic (QR-only, DM-only, 1D-only blocks)
- Formal grade string construction (`1.0/16/660/45Q`)
- Calibration warning, optics disclaimer, application standard note trigger conditions
- Multi-scan summary structure (summary_stylesheet)

Files on disk:
- `attached_assets/html_stylesheet_1781048355486.xslt`
- `attached_assets/report_stylesheet.en_1781048355486.xslt`
- `attached_assets/summary_stylesheet.en_1781048355486.xslt`
(plus DE/ES/JA variants)

---

## 3. GS1 Standard Verification Report Templates — NICE-TO-HAVE OUTPUT FORMAT

**Source**: GS1 General Specifications GSCN WR 25-266, "Update to verification templates",
effective Dec 2025. Release 26.0 Draft (Jan 2026). Sections §5.12.7.3 (1D) and §5.12.7.4 (2D).
File on disk: `attached_assets/GSCN-25-266-Verification-2D_Test_Spec_Update_1781051065221.pdf`

**Feature**: VTCCP should have the ability to output verification reports conforming to the
GS1 standard template format — both the 1D (linear) and 2D templates. This is a globally
recognised format used by testing agencies, print quality laboratories, and compliance
programmes worldwide. **Priority: nice-to-have, not essential for initial release.**

### 3.1 Template structure

Both templates share a common header block and two-column technical analysis layout.

**Common header (both 1D and 2D):**

| Field | Source in VTCCP | Notes |
|---|---|---|
| Name / Address | New UI field | Operator/company identity — session-level |
| Issue date | Scan timestamp | Auto-populated |
| Product description | New UI field | Job-level, requires operator input |
| Type of barcode | `SymbologyName` | Already captured |
| Print method | New UI field | Operator input (flexo, offset, laser, thermal, etc.) |
| Number of barcodes on product | New UI field | Operator input |
| Verifier device | `DeviceModel` | Already captured |
| Verification software version | VTCCP assembly version | Auto-populated |
| Last verifier calibration date | `CalibrationDate` / `FieldCalibrated` | Already captured |

**1D template (§5.12.7.3) — ISO/IEC 15416 parameters:**

| GS1 parameters | Source | ISO/IEC parameters | Source |
|---|---|---|---|
| Barcode structure | Operator/structured check | Overall ISO/IEC grade | `OverallGrade` ✓ |
| X-dimension (magnification) | `NominalXDim` ✓ | Decode | `DecodeGrade` ✓ |
| Barcode height | New — not in VerificationRecord | Symbol contrast | `SCGrade` / `SCValue` ✓ |
| Quiet Zone left/right | `LQZGrade`/`RQZGrade` ✓ | Min reflectance | `MRGrade` / `MRValue` ✓ |
| Human readable | Operator / HRI output | Edge contrast | `ECGrade` ✓ |
| Barcode width | Not captured | Modulation | `ModGrade` ✓ |
| GS1 Company Prefix validity | `ApplicationPass` ✓ | Defects | `DefGrade` ✓ |
| Data structure (syntax) | `ApplicationStandard` / `ApplicationPass` ✓ | Decodability | `DecodabilityGrade` ✓ |
| Print growth (+/-%) | `PrintGrowthHPercent` ✓ | Print growth | `PrintGrowthHGrade` ✓ |

**2D template (§5.12.7.4) — ISO/IEC 15415 parameters:**

| GS1 parameters | Source | ISO/IEC parameters | Source |
|---|---|---|---|
| Matrix size | `MatrixSize` ✓ | Overall ISO/IEC grade | `OverallGrade` ✓ |
| X-dimension/cell size¹ | `NominalXDim` ✓ | Decode | `DecodeGrade` ✓ |
| Data structure (syntax)² | `ApplicationPass` ✓ | Symbol contrast | `SCGrade` / `SCValue` ✓ |
| GS1 Company Prefix validity | `ApplicationPass` ✓ | Modulation | `ModGrade` ✓ |
| Human readable | HRI output | Axial nonuniformity | `ANUGrade` ✓ |
| | | Grid nonuniformity | `GNUGrade` ✓ |
| | | Unused Error Correction | `UECGrade` ✓ |
| | | Print growth (H) | `PrintGrowthHGrade` ✓ |
| | | Print growth (V) | `PrintGrowthVGrade` ✓ |
| | | Fixed pattern damage | `FPDGrade` ✓ |
| | | Clock track / solid area (DM only)³ | `HCTGrade` / `VCTGrade` ✓ |
| | | Quiet Zones QZL1/QZL2 (DM only)³ | `LQZGrade` / `BQZGrade` ✓ |
| | | L1 and L2 (DM only)³ | `LLSGrade` / `BLSGrade` ✓ |
| | | Format information (QR only)⁴ | `QrFIBGrade` ✓ |
| | | Version information (QR only)⁴ | `QrVIBGrade` ✓ |

¹ X-dim/cell size = average of x and y dimension (footnote clarified in this GSCN)
² Data structure footnote updated in this GSCN to include GS1 Digital Link URI Syntax
³ Data Matrix only — ISO/IEC 16022
⁴ QR Code only — ISO/IEC 18004

**Operator-supplied free-text fields (both templates):**
- GS1 barcode placement compliance (yes/no + comment)
- Business critical comments
- Educational comments (analysis narrative)
- Localised notes and disclaimer blocks (standard boilerplate, configurable per agency)

### 3.2 VTCCP coverage assessment

Most ISO/IEC grade fields are already captured in `VerificationRecord`. Fields requiring
new UI input (not automatable from the verifier scan):

| Field | Why manual | Where to collect |
|---|---|---|
| Product description | Not a scan property | Job-level header in new session UI |
| Print method | Not reported by device | Job-level header |
| Number of barcodes on product | Not reported by device | Job-level header |
| Operator/company name + address | Identity of testing agency | Settings / profile UI |
| Barcode height (1D) | Not in push XML or HTML report | Probe DMCC or accept manual |
| Barcode width (1D) | Not in push XML or HTML report | Probe DMCC or accept manual |
| GS1 placement compliance | Physical/visual judgement | Operator checkbox |
| Business critical comments | Analyst narrative | Free-text per record |
| Educational comments | Analyst narrative | Free-text per record |

### 3.3 GSCN WR 25-266 change summary (Dec 2025)

This GSCN made two targeted changes to the General Specifications templates:
1. **Print growth in 1D template**: Clarified that Print Growth is only a *graded*
   parameter for 2D. In the 1D template it is retained as a process control parameter
   only (not a graded ISO/IEC parameter). VTCCP should follow this distinction.
2. **X-dimension/cell size footnote (2D)**: Footnote now explicitly states the value
   is the *average* of x and y cell dimensions — not x-only. NominalXDim in VTCCP
   should reflect this when populating the GS1 2D template field.

### 3.4 Note on the two uploaded files

Both uploaded PDFs (`*065221.pdf` and `*109484.pdf`) are byte-for-byte identical
(confirmed by MD5). The user intended to upload a second separate file (GS1US online
fillable version) but uploaded the same GSCN twice. The GS1US fillable form was not
captured — re-upload needed if reference is desired.

