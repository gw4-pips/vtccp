# VCCS PDF renderer layout check

**Review date:** 2026-08-22  
**Current source:** report `v1.5.41`, application `1.5.57`

## Result

The widened report layout passes the source-level and fixture checks below. A
fresh native Windows PDF render of the current `v1.5.41` source could not be
run in this Linux workspace because neither WebView2 nor the bundled
`wkhtmltopdf.exe` is available here. The native-renderer evidence included in
this review is therefore explicitly identified as the immediately preceding
Windows render (`v1.5.38`), not as a `v1.5.41` render.

## Review samples

- [Current production HTML fixture](vccs-review-v1.5.41.html) — generated
  directly through `VccsHtmlReportGenerator.Generate` using a representative
  multi-mode GS1 record with an embedded barcode, two parser columns, four
  native DFC rows, and RFID fields. SHA-256:
  `20b310aa8c5ae9715fc6b8f18408a5518076f13a9612774387737b84a2233de6`.
- [Windows/Edge rendered page](vccs-windows-v1.5.38-page-1.png) — page 1
  rasterized from the preserved Windows-produced PDF
  `../../../attached_assets/2026-08-21_19-55-45_vccs_rfid_20260821-195542_1787357003680.pdf`.
  The PDF metadata reports one Letter page and an Edge/Chromium creator:
  `Mozilla/5.0 (Windows NT 10.0; Win64; x64) ... Edg/151.0.0.0`.

## Checks performed

### Current `v1.5.41` source fixture

- Letter page geometry remains `8.5in × 11in`.
- Side padding is `0.4in`, leaving the widened content area.
- The barcode image/DFC split is fixed at `19.48% / 80.52%`.
- The dual parser table is `table-layout: fixed` with mirrored
  `17% / 25% / 7%` field/data/check columns and a `2%` center divider.
- The fixture contains both `DataMan TruCheck GS1 Parser` and `VeriWedge GS1
  Parser` headings, the embedded barcode image, the center divider, and the
  four section-border declarations.
- The fixture contains the current footer version `v1.5.41`.

### Preserved native Windows render

The preceding Windows/Edge PDF was inspected at `612 × 792 pt` and has one
page. Its extracted text stayed inside the page's content bounds:

| Measurement | Result |
| --- | ---: |
| Leftmost text | `36.0 pt` |
| Rightmost text | `576.0 pt` |
| Lowest text | `756.0 pt` |

Visual inspection of the preserved page confirmed that the wider content
area, barcode image, dual parser columns, center divider, table borders, and
footer remain visible and aligned without page overflow or clipping. This
confirms the widened layout's behavior in the production Edge path as of
`v1.5.38`; it does not certify the final fixed-column changes in `v1.5.41`.

## Repeating the remaining Windows-only check

On the Windows workstation, build the current source and generate one report
through the normal application path so `VccsPdfRenderer.RenderAsync` selects
WebView2 (or its production fallback). Confirm the resulting PDF is one Letter
page and visually inspect the same items listed above. Update this record
with the new PDF filename, renderer metadata, report version, and screenshot.