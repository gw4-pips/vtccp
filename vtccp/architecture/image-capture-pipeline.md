# VTCCP Image Capture Pipeline

**Version 1.0 — 2026-06-20**

---

## Overview

Every scan produces up to four distinct image artefacts, each from a different source
and with different resolution, encoding, and capture semantics. These must not be
conflated: they are different representations of the same label area at different
levels of zoom, timing, and format.

```
                          ┌──────────────┐
                          │  DM475V SDK  │
                          │  (Cognex SDK)│
                          └──────┬───────┘
                                 │ L3: full sensor frame
                                 │ 2448×2048 JPEG, ~300–800 KB
                                 │ GetResultImage() — SDK only
                                 │
              ┌──────────────────▼──────────────────┐
              │         Device image buffer         │
              │   (firmware ROI crop applied here)  │
              └───────┬──────────────┬──────────────┘
                      │              │
        L2: ROI frame │              │ L1: barcode crop
        IMAGE.SEND    │              │ push XML r.trucheck.jpegImage
        ~label area   │              │ barcode only, ~200–600 px
        incl. HRI     │              │ JPEG, base64, in push result
                      │              │
              ┌───────▼──────┐  ┌───▼──────────────────────┐
              │ RoiJpeg      │  │  VerificationRecord       │
              │ ImageBase64  │  │  .JpegImageBase64         │
              │ (DeviceSession│  │  (parsed from push XML)   │
              │ .GetRoiImage) │  └──────────────────────────┘
              └──────────────┘

L0: DMST native PNG save  ←  separate path, DMST-dependent
    Full-resolution PNG, lossless
    Filename: {timestamp}.png  (paired with {timestamp}.html)
    Captured only when DMST is open AND DMST triggered the scan
```

---

## Image Levels — Full Specification

### L0 — DMST Native PNG Save

| Property | Value |
|---|---|
| Source | DMST TruCheck application (filesystem write) |
| Format | PNG, lossless |
| Resolution | Full sensor resolution (2448×2048 on DM475V / DM395V) |
| Trigger dependency | DMST must be open **AND** must have triggered the scan — CP-triggered scans do NOT produce a saved PNG file (confirmed 2026-05) |
| Path | `{Documents}\{DeviceName}\CodeQuality\{timestamp}.png` |
| Paired with | `{timestamp}.html` (the DMST verification report) |
| VTCCP access | Filesystem watcher — rename only, content never modified |
| Status | Supported as a future file-rename hook (D4 scope) |
| Why PNG matters | Lossless — required for test plate development, sharing with verifier manufacturers, and IMAGE.LOAD re-grading without JPEG-induced grade drift |

**Grade drift risk (confirmed)**: JPEG DCT introduces ringing at module edges (SC%/Modulation),
blocking at 8×8 tile boundaries, and DC coefficient drift. Re-grading from an L0 PNG via
IMAGE.LOAD is the correct round-trip. Re-grading from an L1/L2 JPEG introduces ~0.1–0.2
grade unit drift at borderline grades.

---

### L1 — Barcode Crop (Push XML)

| Property | Value |
|---|---|
| Source | Push XML `r.trucheck.jpegImage` field |
| Field in VerificationRecord | `JpegImageBase64` |
| Format | JPEG, base64-encoded, embedded in the push XML `<JpegImageBase64>` element |
| Resolution | Firmware ROI crop to the barcode boundary only — typically ~200–600 px on longest axis (exact dimensions vary by scan; first raw push output with this field populated will confirm) |
| Timing | Arrives in the push result immediately after the scan, before HTML report |
| DMST dependency | None — present regardless of whether DMST is open |
| Trigger dependency | None — present on CP-triggered and DMST-triggered scans |
| Used for | Excel report embed, IMAGE.LOAD re-verification, OCR source (Block J schema columns) |
| Status | **Captured** — fully parsed and stored in VerificationRecord |

**Note**: This is the same barcode-boundary crop shown in the DMST verification panel.
DMST does NOT show the full-sensor frame in its TC panel — it shows the L1 crop.

---

### L2 — ROI Frame (IMAGE.SEND)

| Property | Value |
|---|---|
| Source | `IMAGE.SEND` DMCC command, issued after scan via `DeviceSession.GetRoiImageAsync()` |
| Field in VerificationRecord | `RoiJpegImageBase64` |
| Format | JPEG, raw bytes (`GetRoiImageAsync` returns `byte[]`) |
| Resolution | Operator-configured ROI frame — wider than L1 barcode crop; includes surrounding label area: HRI text, lot numbers, expiry, adjacent symbols |
| Scale | Controlled by `IMAGE.SIZE` setting: Full (2448×2048) / 1:4 / 1:16 / 1:64 |
| Timing | Fetched on demand **after** push result received — separate DMCC round-trip |
| DMST dependency | None |
| Trigger dependency | None — buffer persists until next scan |
| Method | `_client.GetRoiImageAsync()` → `AttachRoiImageAsync()` → `record with { RoiJpegImageBase64 = ... }` |
| Status | **Implemented** — `GetRoiImageAsync()` and `AttachRoiImageAsync()` in DeviceSession; invocation gated on session config |
| Throughput ceiling | ~1.5–2.5 fps sustained (confirmed: LIVEIMG.SEND is dead on all ports/modes) |

**IMAGE.SIZE setting note**: `IMAGE.SIZE` controls IMAGE.SEND output resolution ONLY. It
does NOT affect the push XML JPEG crop (L1). It does NOT affect DMST PNG save (L0).
Queried on connect via DMCC GET; stored in `VerificationRecord.ImageSizeSetting`.

**LIVEIMG.SEND is dead**: `LIVEIMG.SEND` returns no data on any port or mode on the
DM475V. The `svg_image.img` live-view stream is AES-encrypted and not accessible to
third parties. IMAGE.SEND (L2) is the only viable third-party image retrieval path.

---

### L3 — Full Sensor Frame (SDK)

| Property | Value |
|---|---|
| Source | Cognex DataMan SDK `GetResultImage()` |
| Format | JPEG, ~300–800 KB per frame |
| Resolution | Full sensor: 2448×2048 (DM475V / DM395V / DM395); 2048×1536 (DM394 / DM390) |
| Pixel pitch | 3.45 µm (all DM475V / DM395V / DM395 / DM394 / DM390) |
| DMST dependency | None |
| Status | D4 scope — not yet implemented. SDK connection required (separate from DMCC path) |
| Use case | Full-resolution archival; additive to L1/L2; not required for v1 |

---

## Capture Path by Mode

| Mode | L0 PNG | L1 Barcode crop | L2 ROI (IMAGE.SEND) | L3 SDK full frame |
|---|---|---|---|---|
| Push (DMCC/HTTP) | ✗ (CP-triggered) | ✓ (in push XML) | Optional on demand | D4 |
| Manual / AutoPoll | ✗ (CP-triggered) | ✓ (in push XML) | Optional on demand | D4 |
| DMST-triggered + scraper | ✓ (DMST saves) | ✓ (in push XML) | Optional on demand | D4 |

---

## OpticsSource Discriminator

`OpticsSource` in `VerificationRecord` distinguishes how the graded image was acquired:

| Value | Source | Discriminator fields |
|---|---|---|
| `"LiveScan"` | Fresh live camera frame | `ContrastUniformity` > 0, `MRD` > 0; `r.image.autoExposure = true` |
| `"LoadedImage"` | Pre-loaded JPEG via IMAGE.LOAD + IMAGE.REPLAY | `CU = -1`, `MRD = -1`; `r.image.autoExposure = false`, `exposureTime = 0` |
| `"StitchedImage"` | IMAGE.LOAD via DeviceSession.LoadAndReplayAsync | Same as LoadedImage; explicitly tagged by VTCCP |

**QR live scan anomaly (fw 6.1.16_sr4 confirmed)**: Live QR scans on this firmware
return `CU=-1 / MRD=-1`, which normally signals LoadedImage. VTCCP must NOT rely solely
on CU/MRD for QR — the `r.image.autoExposure` / `exposureTime` secondary discriminator
is required for QR live-scan detection. DM live scans always return CU>0 / MRD>0.

---

## IMAGE.LOAD Re-Grading

`DeviceSession.LoadAndReplayAsync(imagePath, ct)`:
1. Reads the local image file (JPEG/PNG/BMP)
2. Issues `IMAGE.LOAD` with the raw bytes over DMCC
3. Issues `IMAGE.REPLAY` to trigger a full TruCheck grading pass on the loaded image
4. Waits for and returns the push result (poll mode) or receives it via listener
5. Forces `OpticsSource = "LoadedImage"` on the returned record

**Best source for IMAGE.LOAD**: L0 PNG (lossless, full resolution). L1/L2 JPEG will
introduce JPEG compression artefacts into the re-grade result.

---

## OCR Source (Block J)

OCR runs in `AcceptRecordInnerAsync` against the L1 barcode-crop JPEG (`JpegImageBase64`).
The wider L2 ROI frame (which includes HRI text) is architecturally better suited for
OCR of human-readable text alongside the barcode, but L2 capture is optional and not
yet wired into the OCR path. L2 OCR is a future enhancement (Manual mode only, D4 scope).

---

*Cross-references: `optics-source-model.md`, `firmware-confirmed-facts.md` §1 (r.image),
`wireshark-protocol-analysis.md` §4, `sensor-frame-metadata-plan.md`.*
