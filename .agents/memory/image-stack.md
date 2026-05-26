---
name: Three-level image stack
description: Three distinct image sources on DataMan DM475V/DM395V — barcode crop, ROI frame, full sensor frame. DMST crops to barcode only (user-confirmed 2026-05-26).
---

**Rule:** Never conflate the three levels. Each has different content, size, and capture mechanism.

| Level | Name | Source | Size | VTCCP status |
|---|---|---|---|---|
| 1 | Barcode crop | `r.trucheck.jpegImage` → `VerificationRecord.JpegImageBase64` | ~200–600 px | COMPLETE — already captured |
| 2 | ROI frame | `IMAGE.SEND` DMCC (DmccCommand.ImageSend) | IMAGE.SIZE-dependent | Probe needed (D4) |
| 3 | Full sensor frame | `DataManSystem.GetLastReadImage()` SDK | 2448×2048 DM475V/DM395V | Not started (D4) |

**Why:** DMST crops verification panel image to barcode symbol only (not full ROI). Confirmed by user 2026-05-26. Level 1 is therefore a tight symbol crop, not the label ROI. Label text (lot numbers, expiry dates, IUID strings) outside the immediate symbol area requires Level 2 or Level 3.

**Open probe:** Does `IMAGE.SEND` return ROI or full frame? Test via `SendCommandWithExpectedBinaryResult()` after a live scan; inspect JPEG pixel dimensions. If IMAGE.SEND = full frame, Level 2 and Level 3 collapse.

**OCR targeting:** Level 1 covers HRI adjacent to the symbol. Level 2 needed for surrounding label text. Level 3 for full layout.
