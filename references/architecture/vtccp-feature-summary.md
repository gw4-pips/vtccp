# VTCCP — Feature Summary

> v1.1 — 2026-06-10 · "Shipped" = code complete + device-confirmed or test-confirmed.

---

## 1. Core Session Management

| Feature | Status | Notes |
|---|---|---|
| Device configuration store | Shipped | `ConfigRepository` + `DeviceConfig`; supports multiple named devices |
| SDK connection (DataManSdkClient) | Shipped | Cognex SDK via reflection; SDK DLL not redistributed |
| HTTP event subscriber | Shipped | `GET /events?enable` on port 44444; DMST-independent result delivery |
| Push listener (TCP raw XML) | Shipped | `StartPushListenerAsync`; DmstListener on DmstListenPort |
| Poll mode (TRIGGER ON + result wait) | Shipped | `TriggerAndGetResultAsync`; fallback when push unavailable |
| Session state / context stamping | Shipped | `SessionState` stamped on every `VerificationRecord` |
| Device info query on connect | Shipped | `DEVICE.TYPE`, firmware, serial → `DeviceInfo` |
| Connection medium auto-detection | Shipped | GigE / USB-Ethernet / USB-COM inferred from IP prefix |

---

## 2. Verification Record Model

| Feature | Status | Notes |
|---|---|---|
| `VerificationRecord` core grade fields | Shipped | All ISO 15415 (2D) grade parameters |
| ISO 15416 (1D) grade fields | Shipped | EAN-13, Code 128, UPC — all 1D parameters |
| QR Code grade fields (8 pattern grades) | Shipped | ULP/URP/LLP/HCT/VCT/ALP/VIB/FIB; device-confirmed scan #2 |
| DM — Data Codewords / EC Budget | Shipped | ECC200 table lookup (q.symbols=null on fw 6.1.16_sr4) |
| QR — Data Codewords / EC Budget | Shipped | QR table lookup from MatrixSize |
| Modulation Values array | Shipped | `ModulationValuesData`; variable-size grid per symbol |
| Codeword Values array | Shipped | `CodewordValuesData`; ECC200 boundary marker; per-codeword grade |
| GS1 application parse | Shipped | `gs1-syntax-engine` v1.4.0; `ApplicationStandard/Pass/Reason` |
| OpticsSource tagging | Shipped | `LiveScan` / `LoadedImage` / `StitchedImage` |
| Connection metadata fields | Shipped | `DeviceModel`, `ConnectionAddress`, `ConnectionMedium` |
| `FormalGrade` (ISO notation) | Shipped | `1.0/16/660/45Q` format from DMST HTML scraper |
| `ImagePolarity` | Shipped | From DmstHtmlScraper `ParseHtml()` |
| `CalibrationWarning` flag | Shipped | Set when `FieldCalibrated=false` (all observed scans) |
| `ECLevel` (QR) | Not in push XML | In DMST HTML report — DmstHtmlScraper ParseHtml() extension pending |
| `DataMaskPattern` (QR) | Not in push XML | Same — pending |
| `ECI` value | Not in push XML | Same — pending |
| `EncodedCharacters` | Unresolvable | Both push paths dead on fw 6.1.16_sr4 — emit empty |
| ISO grading standard edition | Echo device | NOT hard-coded; echoed from DMST report string |

---

## 3. Excel Report Engine

| Feature | Status | Notes |
|---|---|---|
| TruCheck-compatible main sheet | Shipped | `TruCheckCompatibleSchema`; column layout matches TruCheck export |
| ISO 15415 (2D) column set | Shipped | `DataMatrix2DMapper` |
| ISO 15416 (1D) column set | Shipped | `ISO15416Mapper` |
| Universal columns (model / connection) | Shipped | Reader Model, Connection, Link Medium — all 3 mappers |
| Modulation Values sub-sheet | Shipped | `ModValuesSheetWriter`; isBlack coloring; variable grid |
| Codeword Values sub-sheet | Shipped | `CwValuesSheetWriter`; data/ECC boundary marker |
| Company logo in title row | Shipped | 160×54 px banner at column I; XLS + XLSX adapters |
| XLS adapter (NPOI) | Shipped | `XlsAdapter` |
| XLSX adapter (EPPlus) | Shipped | `XlsxAdapter` |
| Job template enforcement | Shipped | Template → `JobTemplate`; session validates fields |
| **Session / job parent rows** | Shipped | Each job writes a header parent row; individual scan records are indented child rows beneath it, grouped by session |
| **1D ISO 15416 — 10-scan child rows** | Planned | ISO 15416 requires sampling at 10 equally-spaced heights across the bar. Each height = one child row; the parent row holds the worst-case (minimum) grade per parameter. Not yet implemented — requires per-scan sub-image capture at each sample line. |
| **Column outline grouping (collapse/expand)** | Planned | EPPlus `GroupColumn()` / NPOI `SetColumnOutlineLevel()` allow the operator to collapse less-used column sections (e.g. raw numeric arrays, DPM-only columns) with Excel's native +/− outline controls. Design decision pending: which column groups to expose. |
| **Parse-result child rows** | Planned | GS1 AI parse results, ISO 15434 segment parse results, and supplement decode results will each expand as child rows beneath the main record row — consistent with DMST TruCheck's layout. Column grouping and row outline levels should mirror TruCheck for drop-in compatibility. |

---

## 4. Push Script Infrastructure

| Feature | Status | Notes |
|---|---|---|
| Push script authoring | Shipped | `DmstPushScript_v1.js` — current production: v1.34 |
| Script Viewer web artifact | Shipped | `artifacts/script-viewer` — side-by-side diff view |
| Version bump rule | Convention | Every code-change commit bumps `<Version>` in `VtccpApp.csproj` |
| v1.33 device-confirmed | Shipped | Probe campaign complete 2026-05-24 — 4 fields resolved |
| v1.34 production build | Shipped | Debug probes stripped; `r.barcodeAssignment` probe queued |

---

## 5. DMST HTML Scraper

| Feature | Status | Notes |
|---|---|---|
| `DmstHtmlScraper` + `ParseHtml()` | Shipped | Reads `{Documents}\{DeviceName}\CodeQuality\*.html` |
| `DmstHtmlReport` model | Shipped | All scraped fields typed |
| `DmstReportValidator` | Shipped | Report freshness check |
| DeviceSession wiring | Shipped | `TryMergeAsync` called in poll + push + load paths |
| `ImagePolarity` from HTML | Shipped | "Black on white" / "White on black" |
| `ECLevel`, `DataMaskPattern`, `ECI` from HTML | Pending | `ParseHtml()` extension — QR-only fields |
| `FormalGrade` from HTML | Shipped | `1.0/16/660/45Q` format confirmed |
| `EncodedCharacters`, `DataCW`, `ECBudget` from HTML | Shipped | Resolves push XML bugs #1, #5, #6 |

---

## 6. Live Feed

| Feature | Status | Notes |
|---|---|---|
| GetFreshFrameAsync (TRIGGER ON + IMAGE.SEND) | Shipped | 1.5–2.5 fps; no LIVEIMG.SEND (dead on fw 6.1.16_sr4) |
| LiveFeedWindow — image display | Shipped | `LiveFeedViewModel` + `LiveFeedWindow.xaml` |
| Crosshair overlay | Shipped | Two 1.5 px red Rectangles + 5 px centre dot; `SnapsToDevicePixels`; always visible |
| **ROI — user-drawn rubber-band rectangle** | Shipped | Left-click + drag anywhere in the FOV draws a free-form amber dashed rectangle. Right-click clears it with a fade-out hint. Normalised coordinates (0–1 per axis) stored in `LiveFeedViewModel.RoiNormalized` for downstream use (IMAGE.LOAD ROI, report overlay). Accidental single-pixel clicks are ignored. |
| Go Live / Freeze / Verify state machine | Shipped | Three-state: Idle → Live → Frozen |
| Grades gated on Frozen state | Shipped | `OnResultReceived` ignores results during Live state |

---

## 7. Symbol Stitching (STITCH-1 Phase 1)

| Feature | Status | Notes |
|---|---|---|
| `StitchingEngine` — skew correction | Shipped (untested) | Bar-top edge sampling at ¼/¾ column positions; atan2 angle. Adequate for Phase 1 on flat surfaces; see Phase 2 clock-track note below. |
| `StitchingEngine` — composite | Shipped (untested) | Horizontal concat at leftSeam/rightSeam; height normalisation |
| `StitchingEngine` — seam estimate | Shipped (untested) | 75% of left-image width as default |
| **Vertical alignment default — bottom-align** | **Design decision** | For 1D symbols the composite must be bottom-aligned, not centre-aligned. HRI text sits at the bottom in the overwhelming majority of label formats; even 1–2 px vertical misalignment makes the HRI look broken and destroys operator trust, even when the bar grades are correct. 1D's vertical redundancy means a small vertical offset is harmless for grading but fatal for credibility. `StitchingEngine.Composite` should pad from the top (align bottom edges), not centre-pad. Implementation change required before first real test. |
| `StitchingViewModel` — capture state machine | Shipped | Idle → CapturingLeft → LeftCaptured → CapturingRight → BothCaptured → Previewing → Verifying → Result |
| `StitchingWindow` — three-panel UI | Shipped | Left / Composite / Right panels; seam slider; Verify button |
| `DeviceSession.LoadImageAndVerifyAsync` | Shipped | IMAGE.LOAD + IMAGE.REPLAY; OpticsSource = "StitchedImage" |
| Phase 1 test images | **Pending** | C128 FX label images not yet received — parameters may need tuning |
| **Phase 2 — clock-track alignment (2D)** | Planned | For DataMatrix the alignment mechanism must be the clock tracks (top row, right column of the symbol matrix), not the crude bar-top-edge heuristic used in Phase 1. Algorithm intent: (1) detect the horizontal clock track row in each half-image via projection peak analysis; (2) align them to sub-pixel accuracy; (3) measure L-pattern (finder pattern: left column + bottom row) angle to confirm rotation is fully corrected — any residual L-pattern skew indicates the composite is still rotated and must be re-corrected. Clock-track cross-correlation in the overlap region is the Phase 2 seam-detection candidate. |
| **Phase 2 — cross-correlation seam detection** | Planned | Automatic seam placement using normalised cross-correlation of overlapping columns in the two corrected half-images; replaces manual slider for standard cases. |

---

## 8. OCR Engine

| Feature | Status | Notes |
|---|---|---|
| `DualEngineOcrRunner` | Shipped | Windows.Media.Ocr (primary) + Tesseract 5.2.0 (fallback) |
| `tessdata/eng.traineddata` | Shipped | Bundled at build output |
| OCR on verification image | Shipped | Runs in `AcceptRecordInnerAsync` |
| UI toggle (enable/disable OCR) | TODO | `_ocrEnabled` defaults true; no UI control yet |

---

## 9. UPC/EAN Supplemental Mode

| Feature | Status | Notes |
|---|---|---|
| Read/write `UPC-EAN.SUPPLEMENT` via DMCC | Shipped | 5 modes: Ignore / Required / Required-2 / Required-5 / Not-Required |
| Session UI — Device Configuration card | Shipped | 6 radio buttons (5 modes + header); Read / Apply buttons |
| **SUPP-1 — Proprietary supplement bar/space decoder + ISO grader** | **Planned** | Cognex does not expose supplement quality grades (2-digit / 5-digit add-ons for UPC/EAN). VTCCP plan: extract the supplement bar/space sequence from the verification JPEG using an edge-detection bar/space decoder operating on the image provided by the device; grade the physical quality parameters per ISO 15416 and the UPC/EAN symbology specification — Decodability principally, plus Edge Contrast, Modulation, and Defects where the image resolution permits. Results caveated explicitly as **"VTCCP Proprietary — not DMST TruCheck grade"** so the qualification boundary is unambiguous. Implemented and surfaced whenever a supplement is present and the device supplement mode is not Ignore. |
| **SUPP-2 — Middle Margin (inter-character gap) validation** | **Planned** | The space between the rightmost bar of the main EAN/UPC symbol and the leftmost bar of the supplement start pattern has **both a minimum AND a maximum** prescribed width per the GS1 General Specifications — unlike ordinary quiet zones, which carry only a minimum. Too narrow: the scanner reads the supplement bars as a continuation of the main symbol guard pattern. Too wide: the scanner fails to associate the supplement with the main symbol at all. Axicon (barcode quality instruments) uses the term **"Middle Margin"** for this space; the same term appears in their EPS (Electronic Point-of-Sale) symbol encoding software, where they are the specification authority because they build the encoding side as well as the verification side. VTCCP plan: measure Middle Margin width from the verification JPEG, compare both bounds against the applicable GS1/symbology spec, and report as a separate named parameter with pass/warn/fail outcome and the measured vs. specified range. |
| **SUPP-3 — Left / Right Quiet Zone grading (independent, image-based)** | **Planned** | ISO 15416:2016 Table 1 lists Quiet Zone as a **single combined grade parameter** — the standard defines it as a pass/fail against the symbology specification minimum width, covering both sides, with no requirement to report them separately. Cognex TruCheck has always reported a single QZ grade (user has requested independent L/R reporting over a period spanning two decades; no change has ever been made). VTCCP plan: extract bar/space widths from the verification JPEG for the quiet zone region on each side independently, compute measured width in modules (X-dimensions), compare each side against the symbology specification minimum, cross-validate our measurements against the TruCheck combined QZ value to calibrate the image-based measurement, and report **LQZGrade** and **RQZGrade** as separate proprietary parameters. Each grade indicates pass or fail for that side, so an operator can immediately see which edge is deficient. Labelled "VTCCP Proprietary" — beyond the single-parameter scope of ISO 15416 and beyond what any Cognex verifier currently reports. Applicable to all linear symbologies: EAN-13, EAN-8, UPC-A, UPC-E, Code 128, ITF, Code 39, Codabar, etc. |

---

## 10. Sensor & Frame Metadata

| Feature | Status | Notes |
|---|---|---|
| Session-level sensor spec (DM475V / DM395 / DM390) | Shipped | `DeviceSensorSpec`; 2448×2048 / 3.45µm confirmed from manuals |
| `IMAGE.SIZE` DMCC query on connect | Shipped | 4-level enum (Full / ¼ / 1/16 / 1/64) |
| Per-scan frame dimensions from JPEG SOF0 | D4 scope | `VerifFrameWidthPx/HeightPx` in VerificationRecord |

---

## 11. Data Parsers

VTCCP runs multiple post-decode parse passes on the raw decoded string.  Each parser is
independent; failures are non-fatal and surfaced as informational fields, never grade overrides.

### 11a. GS1 Syntax Engine (shipped)

| Feature | Status | Notes |
|---|---|---|
| GS1 Application Identifier parsing | Shipped | `gs1-syntax-engine` v1.4.0 — covers GS1 DataMatrix (`]d2`), GS1-128 (Code 128 with FNC1), GS1 QR (`]Q3`), and any bearer symbology carrying a GS1 string |
| GS1 Digital Link resolver | Shipped | DL URI (https://…) decomposed to AI set; same output as flat GS1 string |
| `]d2` GS1 DataMatrix detection | Shipped | FNC1 first position; `0x1D` → `\|` before XML parse |
| ApplicationStandard / Pass / Reason | Shipped | Three fields from push XML; informational only — never overrides ISO grade |
| GS1 format fail design rule | Confirmed | VTCCP uses `OverallGrade` (ISO) as operative grade; `ApplicationPass` = informational |

### 11b. AIM ID Decoder (shipped)

| Feature | Status | Notes |
|---|---|---|
| AIM ID prefix decode | Shipped | `]d1` (DM), `]Q1` (QR), `]E0` (EAN-13), `]C0` (Code 128), etc. |
| Symbology classification | Shipped | `ClassifySymbology` map; "QR" prefix added for QR Code |
| Modifier byte semantics | Confirmed | `]Q1` modifier=1 encodes ECI presence (not ECLevel) — confirmed scan #2 |

### 11c. ISO 15434 / MIL-STD-130 Parser (planned)

| Feature | Status | Notes |
|---|---|---|
| ISO 15434 transport layer parse | Planned | Multi-record envelope: `[)>RS06…` / `[)>RS12…` format segments |
| MIL-STD-130 DI decode | Planned | Data Identifier prefix decode for defence/government labels |
| **Blocked by** | C-series materials | ISO 15434 and MIL-STD-130 documents not yet acquired |

### 11d. Supplement Image Decoder (planned — SUPP-1)

| Feature | Status | Notes |
|---|---|---|
| Bar/space extraction from JPEG | Planned | Edge-detect the supplement bars/spaces from the device-provided JPEG |
| ISO 15416 physical quality grading | Planned | Decodability (principal parameter), Edge Contrast, Modulation, Defects |
| Caveat labelling | Planned | All supplement grades labelled "VTCCP Proprietary" — see §9 SUPP-1 |

---

## 12. ISO 29158 / DPM Support

| Feature | Status | Notes |
|---|---|---|
| DPM grading fields (DMV-8072V) | Planned | `CellContrast`, `CellModulation` partially inventoried |
| Full ISO 29158 schema | **Blocked** | Waiting for C1 (ISO/IEC 29158 document) |
| DPM not out-of-scope | Note | Core stated feature — do not omit from descriptions |

---

## 13. Planned / Future Features

| Feature | Notes |
|---|---|
| **D1 — HTML/PDF verification report** | Blocked by C3 (Webscan docs); CalibrationWarning + OpticsSource disclaimer + logo embed |
| **D2 — Reverse-report from Excel** | Blocked by D1 |
| **D3 — QR Code full parity** | ECLevel / DataMaskPattern / ECI via ParseHtml() extension; DebugRSymbology v1.34 |
| **D4 — Full IMAGE.LOAD implementation** | Unblocked; sidecar XML archival; full-frame archive optional |
| **SUPP-1 — Supplement bar/space decoder + ISO grader** | Proprietary; see §9 and §11d; unblocked pending test images |
| **SUPP-2 — Middle Margin validation** | Min AND max prescribed by GS1; Axicon EPS terminology; see §9 |
| **SUPP-3 — Independent L/R Quiet Zone grading** | ISO 15416 defines QZ as a single combined grade; Cognex has never split it; VTCCP to measure each side from JPEG and report LQZGrade + RQZGrade separately; calibrate against TC combined value; labelled VTCCP Proprietary; see §9 |
| **STITCH-1 Phase 2 — clock-track alignment** | See §7; 2D DataMatrix alignment via clock tracks + L-pattern test |
| **Excel 1D 10-scan child rows** | ISO 15416 per-height sub-records; see §3 |
| **Excel column outline grouping** | Collapse/expand column groups; design decision on group boundaries pending |
| **VCCS Command Pilot** | Blocked by DMST TC window screenshots from user |
| **Batch external symbol upload** | Future; IMAGE.LOAD × N loop |
| **Directional lighting detection** | Sobel gradient histogram on DPM JPEG; future/no timeline |
| **Image-based OpticsSource discrimination** | CU/MRD + r.image.exposureTime discriminator; QR IMAGE.LOAD path open |
| **Per-scan r.image metadata** | exposureTime, gain, FoV, RoI, illIntensity, mirrorAngle — D4 scope |
| **Config snapshot / restore** | Full DMCC parameter profile save/load |

---

## Reference Library (71+ files)

All A-series docs complete.  Key files:

| Path | Content |
|---|---|
| `references/architecture/optics-source-model.md` | OpticsSource tri-state; ISO disclaimer logic |
| `references/architecture/vtccp-vs-dmst-feature-matrix.md` | Side-by-side VTCCP ↔ DMST capability comparison |
| `references/architecture/coglink-reference-log.md` | Coglink facts from DM390 manual |
| `references/architecture/sensor-frame-metadata-plan.md` | Sensor spec + IMAGE.SIZE + per-scan dims plan |
| `references/architecture/firmware-confirmed-facts.md` | All device-confirmed probe results (scans #1–#15) |
| `references/architecture/wireshark-protocol-analysis.md` | Port 44444 mux; codes.xml format; push XML new fields |
| `references/architecture/backlog.md` | Full feature backlog with status |
| `references/samples/live-scans/` | All 15 scan catalogs |
