---
name: Three-level image stack
description: Three distinct image sources on DataMan DM475V/DM395V. Level 2 ROI frame now implemented. OCR image source selection rule confirmed by user.
---

**Rule:** Never conflate the three levels. Each has different content, size, and capture mechanism.

| Level | Name | Source | Size | VTCCP status |
|---|---|---|---|---|
| 1 | Barcode crop | `r.trucheck.jpegImage` → `VerificationRecord.JpegImageBase64` | ~200–600 px | COMPLETE |
| 2 | ROI frame | `IMAGE.SEND` → `VerificationRecord.RoiJpegImageBase64` | IMAGE.SIZE-dependent | IMPLEMENTED — pending device confirmation |
| 3 | Full sensor frame | `DataManSystem.GetLastReadImage()` SDK | 2448×2048 DM475V/DM395V | Not started |

**Why:** DMST crops verification panel image to barcode symbol only (not full ROI). Confirmed by user 2026-05-26. HRI and surrounding label text (lot numbers, expiry dates, IUID strings) are OUTSIDE the barcode crop for all symbologies except UPC/EAN.

**OCR image source selection (user-confirmed 2026-05-26):**
- Default (all except UPC/EAN): Level 2 ROI frame (`RoiJpegImageBase64`)
- UPC/EAN exception: Level 1 barcode crop — HRI is canonically part of the UPC/EAN symbol presentation
- Fallback when ROI unavailable: Level 1 barcode crop

**Level 2 implementation details:**
- `DataManSdkClient.GetRoiImageAsync()`: Stage 1 = SDK `SendCommandWithExpectedBinaryResult("IMAGE.SEND")` via reflection; Stage 2 = raw TCP fallback
- `StripDmccHeader()` finds `\r\n\r\n` boundary in first 64 bytes to isolate JPEG payload
- `IsJpeg()` validates 0xFF 0xD8 SOI marker before accepting any response
- `DeviceSession.AttachRoiImageAsync()` wired into `TriggerAndGetResultAsync` and `ReplayAndGetResultAsync`
- For loaded-image replays: IMAGE.SEND returns the loaded image itself (not live camera ROI) — still useful as OCR source

**Open probe:** First device scan will reveal which path (SDK vs TCP) succeeds, and whether IMAGE.SEND returns the DMST-configured ROI or the full sensor frame.
