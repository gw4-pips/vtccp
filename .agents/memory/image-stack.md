---
name: Three-level image stack
description: Image layer model for DataMan DM475V — L0/L1/L2/L2.ROI/L3 with naming corrections (2026-06-24)
---

**Rule:** Never conflate the levels. Each has different content, size, and capture mechanism.

| Level | Name | Source | Size | VTCCP status |
|---|---|---|---|---|
| L0 | DMST Native PNG | DMST filesystem | Full sensor 2448×2048, lossless | DMST-triggered only |
| L1 | Barcode crop | Push XML `JpegImageBase64` | Firmware bounding box ~200–600 px, JPEG | COMPLETE |
| L2 | Full camera scene | `IMAGE.SEND` DMCC | Full frame at IMAGE.SIZE (e.g. 1224×1024 at SIZE=1) | IMPLEMENTED |
| L2.ROI | Virtual ROI crop | L2 cropped to DECODER.ROI | Operator scan region — wider than L1, tighter than L2 | **Virtual — derived, not captured** |
| L3 | Full sensor frame | SDK `GetResultImage()` | 2448×2048 | D4 scope |

**Critical naming correction (2026-06-24):**
- `IMAGE.SEND` returns the **full camera scene**, NOT an ROI crop.
- `RoiJpegImageBase64` in VerificationRecord is a legacy misnomer — content is full scene.
- L1 (barcode crop) = firmware-generated tight bounding box around symbol only (DMST TC panel view).
- L2.ROI (Virtual ROI crop) = DECODER.ROI region derived by cropping L2 to `r.image.RoI` coordinates.
  ROI coords confirmed present in r.image 28-key inventory (scan #12, 2026-05-24).
  NOT a separately captured image — computed in software, not a firmware API output.

**PNG metadata opportunity (logged 2026-06-24, not yet implemented):**
PNG `tEXt` chunks can embed DECODER.ROI coords + scan metadata at L2 save time, making the
Virtual ROI Crop reconstructable from the PNG alone. BMP has no metadata capability.
Planned evaluation: IMAGE.FORMAT=2 (PNG) from IMAGE.SEND + tEXt metadata write.

**Why:**
DMST crops verification panel to barcode symbol only (L1). HRI and surrounding label text
are outside the barcode crop. IMAGE.SEND returns the full camera scene (L2); confusing it
with the ROI crop caused architectural errors.

**OCR image source selection (user-confirmed 2026-05-26):**
- Default (all except UPC/EAN): L2 full camera scene — includes HRI, lot, expiry
- UPC/EAN exception: L1 barcode crop — HRI is canonically part of the symbol
- Fallback when L2 unavailable: L1 barcode crop

---

## Clean-state device settings (2026-06-08 snapshot — authoritative)

| Setting | Value | Notes |
|---|---|---|
| `TRIGGER.TYPE` | 0 | Single/External — **NEVER change this** |
| `LIVEIMG.MODE` | 0 | Stay at 0; setting to 2 caused NVRAM corruption previously |
| `IMAGE.FORMAT` | 1 | JPEG (0=JPEG, 2=PNG from fw v5.6.3+) |
| `IMAGE.SIZE` | 1 | DMST default = 1/4 area = 1224×1024 |
| `IMAGE.QUALITY` | 50 | JPEG quality |
| `DECODER.ROI` | 0 2448 0 2048 | Sensor pixel space. At IMAGE.SIZE=1 → scale coords by 0.5 for IMAGE.SEND pixel space. |

## Go Live architecture (confirmed from DMST Wireshark + clean-state snapshot)

TRIGGER.TYPE stays 0 throughout. Live feed = client-side polling loop:
- Fire `TRIGGER ON` every 400 ms via raw TCP port 23 (with extended ACK)
- Wait ~150 ms for sensor readout
- Pull frame via `IMAGE.SEND` on same TCP connection
- Display full 1224×1024 frame (L2) — do NOT replace with L1 barcode crop

On Verify: stop loop, wait 400 ms, fire ONE `TRIGGER ON`, keep last full frame.
On Cancel / Dispose: stop loop, no device command needed.

## DECODER.ROI (user-defined scan ROI — DMCC confirmed)

Command: `SET DECODER.ROI {left} {right} {top} {bottom}`
Platform: ALL (including DM475V), v3.0.0+
Constraints:
- left, right: multiples of 8; right ≥ left+64
- top, bottom: multiples of 4; bottom ≥ top+64
- Coordinates relative to upper-left of displayed image
- Must stay within CAMERA.FOV settings

At IMAGE.SIZE=1 (1224×1024), sensor scale factor = 2×.
CAMERA.XPAND-ROI is NOT for DM475V (FOVE/1DDataStitching, different platform list).
