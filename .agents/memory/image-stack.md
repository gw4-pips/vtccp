---
name: Three-level image stack
description: Three distinct image sources on DataMan DM475V/DM395V. Clean-state snapshot confirmed DECODER.ROI and Go Live architecture.
---

**Rule:** Never conflate the three levels. Each has different content, size, and capture mechanism.

| Level | Name | Source | Size | VTCCP status |
|---|---|---|---|---|
| 1 | Barcode crop | `r.trucheck.jpegImage` → `VerificationRecord.JpegImageBase64` | ~200–600 px (firmware bounding box) | COMPLETE |
| 2 | Full camera frame | `IMAGE.SEND` via raw TCP port 23 | IMAGE.SIZE=1 → 1224×1024 | IMPLEMENTED in LiveFeedClient |
| 3 | Full sensor frame | `DataManSystem.GetLastReadImage()` SDK | 2448×2048 DM475V/DM395V | Not started |

**Why:** DMST crops verification panel image to barcode symbol only (not full ROI). Confirmed by user 2026-05-26. HRI and surrounding label text (lot numbers, expiry dates, IUID strings) are OUTSIDE the barcode crop for all symbologies except UPC/EAN.

**OCR image source selection (user-confirmed 2026-05-26):**
- Default (all except UPC/EAN): Level 2 ROI frame (`RoiJpegImageBase64`)
- UPC/EAN exception: Level 1 barcode crop — HRI is canonically part of the UPC/EAN symbol presentation
- Fallback when ROI unavailable: Level 1 barcode crop

---

## Clean-state device settings (2026-06-08 snapshot — authoritative)

| Setting | Value | Notes |
|---|---|---|
| `TRIGGER.TYPE` | 0 | Single/External — **NEVER change this** |
| `LIVEIMG.MODE` | 0 | Stay at 0; setting to 2 caused NVRAM corruption previously |
| `IMAGE.FORMAT` | 1 | JPEG |
| `IMAGE.SIZE` | 1 | **DMST TruCheck default** — quarter resolution → IMAGE.SEND returns 1224×1024 |
| `IMAGE.QUALITY` | 50 | JPEG quality |
| `DECODER.ROI` | 0 2448 0 2048 | **Full sensor** — L1 barcode crop is NOT from DECODER.ROI |

## Go Live architecture (confirmed from DMST Wireshark + clean-state snapshot)

TRIGGER.TYPE stays 0 throughout all states. Live feed is a **client-side polling loop**:
- Fire `TRIGGER ON` every 400 ms via raw TCP port 23 (with extended ACK)
- Wait ~150 ms for sensor readout
- Pull frame via `IMAGE.SEND` on same TCP connection
- Display the full 1224×1024 frame (L2) — do NOT replace with L1 barcode crop

On Verify:
- Stop the polling loop
- Wait ~400 ms for any in-flight poll to complete
- Fire ONE `TRIGGER ON` via raw TCP port 23
- Keep last full IMAGE.SEND frame visible — do NOT replace with L1 barcode crop from JpegImageBase64

On Cancel / Dispose: stop loop, no device command needed.

## DECODER.ROI (user-defined scan ROI — DMCC confirmed)

Command: `SET DECODER.ROI {left} {right} {top} {bottom}`
Platform: ALL (including DM475V), v3.0.0+
Constraints:
- left, right: **multiples of 8**; right ≥ left+64
- top, bottom: **multiples of 4**; bottom ≥ top+64
- Coordinates relative to upper-left of **displayed image**
- Must stay within CAMERA.FOV settings

At IMAGE.SIZE=1 (display 1224×1024), sensor scale factor = 2×.
CAMERA.XPAND-ROI is NOT for DM475V (FOVE/1DDataStitching, different platform list).
