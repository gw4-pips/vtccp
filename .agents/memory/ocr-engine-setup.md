---
name: OcrEngine project setup
description: Dual-engine OCR scaffold (Windows.Media.Ocr + Tesseract 5.2.0). VtccpApp TFM bumped. Tessdata runtime requirement.
---

**Rule:** OcrEngine project targets `net8.0-windows10.0.18362.0` — this is required for WinRT Windows.Media.Ocr access. VtccpApp was bumped to the same TFM.

**Why:** `net8.0-windows` (unversioned) does not expose WinRT projections. Windows.Media.Ocr needs at minimum `net8.0-windows10.0.14393.0`; 18362.0 (Windows 1903) was chosen as a safe baseline matching Visual Studio 2019+ tooling.

**Tesseract runtime requirement:** `tessdata/eng.traineddata` (~12 MB) must be present at `{exe dir}/tessdata/eng.traineddata` at runtime. Download from:
  `https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata`
The Tesseract NuGet (5.2.0) bundles the engine DLLs — only the language data file needs manual placement. In the deployed Command Pilot installer, tessdata is bundled.

**Windows.Media.Ocr:** Zero install. Built into Windows 10/11. Nothing for the user to download.

**ExcelEngine dependency design:** `VerificationRecord` uses `OcrResultDto` (in ExcelEngine) instead of `OcrEngine.OcrResult` directly, to avoid a hard compile-time reference from ExcelEngine → OcrEngine. Command Pilot maps `OcrEngine.OcrResult` → `OcrResultDto` before storing on the record.

**Confidence tier model:** High (exact match) / Medium (edit dist ≤2) / Low (edit dist >2) / Single (one engine) / Unreadable (both fail). Agreed text = Windows engine output (higher baseline on clean label stock).
