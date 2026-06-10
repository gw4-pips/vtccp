# VTCCP — Feature Summary

> Status as of 2026-06-10.  "Shipped" = code complete + device-confirmed or test-confirmed.

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
| Crosshair overlay | Shipped | Two 1px red Rectangles; always visible; no hit test |
| ROI guide overlay | Shipped | Dashed amber rectangle (~67% of frame); dismissed on click; resets on Go Live |
| Go Live / Freeze / Verify state machine | Shipped | Three-state: Idle → Live → Frozen |
| Grades gated on Frozen state | Shipped | `OnResultReceived` ignores results during Live state |

---

## 7. Symbol Stitching (STITCH-1 Phase 1)

| Feature | Status | Notes |
|---|---|---|
| `StitchingEngine` — skew correction | Shipped (untested) | Bar-top edge sampling at ¼/¾ column positions; atan2 angle |
| `StitchingEngine` — composite | Shipped (untested) | Horizontal concat at leftSeam/rightSeam; height normalization |
| `StitchingEngine` — seam estimate | Shipped (untested) | 75% of left-image width as default |
| `StitchingViewModel` — capture state machine | Shipped | Idle → CapturingLeft → LeftCaptured → CapturingRight → BothCaptured → Previewing → Verifying → Result |
| `StitchingWindow` — three-panel UI | Shipped | Left / Composite / Right panels; seam slider; Verify button |
| `DeviceSession.LoadImageAndVerifyAsync` | Shipped | IMAGE.LOAD + IMAGE.REPLAY; OpticsSource = "StitchedImage" |
| Phase 1 test images | **Pending** | C128 FX label images not yet received — parameters may need tuning |
| Phase 2 automatic seam detection | Planned | Cross-correlation or projection-based; no timeline |

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

---

## 10. Sensor & Frame Metadata

| Feature | Status | Notes |
|---|---|---|
| Session-level sensor spec (DM475V / DM395 / DM390) | Shipped | `DeviceSensorSpec`; 2448×2048 / 3.45µm confirmed from manuals |
| `IMAGE.SIZE` DMCC query on connect | Shipped | 4-level enum (Full / ¼ / 1/16 / 1/64) |
| Per-scan frame dimensions from JPEG SOF0 | D4 scope | `VerifFrameWidthPx/HeightPx` in VerificationRecord |

---

## 11. GS1 Digital Link

| Feature | Status | Notes |
|---|---|---|
| GS1 Digital Link resolver | Shipped | `gs1-syntax-engine` v1.4.0; DL URI → AI decomposition |
| `]d2` GS1 DataMatrix detection | Shipped | FNC1 first position; `0x1D` → `|` before XML parse |

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
