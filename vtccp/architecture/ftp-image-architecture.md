# VTCCP FTP-IMAGE Architecture

**Version 1.0 — 2026-06-24**  
**Status**: Hypothesis confirmed by Glenn Reuss (Cognex chief eng). ONE PROBE NEEDED to confirm full-frame content.

---

## Purpose

FTP-IMAGE is the DMCC-controlled mechanism by which the DM475V pushes a full-frame
image to an external FTP or SFTP server after each scan. It is almost certainly the
same firmware image-save mechanism that underlies DMST's "Log All Decoded Images"
function in the Data Logging tab — but delivered via FTP instead of local filesystem.

If confirmed full-frame: FTP-IMAGE eliminates all three problems with the current
DMST image archival path:
1. **Non-sticky settings** — DMST "Log All Decoded Images" checkbox reverts on restart
2. **DMST trigger dependency** — DMST filesystem save does NOT fire for CP-triggered scans
3. **SDK dependency** — FTP-IMAGE requires no Cognex SDK call (`GetResultImage()`)

---

## Image Content

### What we have confirmed
- `JpegImageBase64` in push XML codes.xml = **firmware-generated ROI crop** of the barcode symbol
  - Always JPEG
  - Always the barcode region, not the full frame
  - Typically 200–600 px on a side
  - No format or size configuration knobs
  - This is Level 1 (barcode crop) in the VTCCP image stack

### What FTP-IMAGE is hypothesized to deliver
- **Full sensor frame** at the resolution and format set by `IMAGE.SIZE` + `IMAGE.FORMAT`
- This is Level 3 (full frame, 2448×2048 at IMAGE.SIZE=0) in the VTCCP image stack
- Glenn Reuss (Cognex chief eng, 2026-06-24): "probably whatever image size and type is set by the DMCC commands or in the DMST UI"

### Probe to confirm
1. Set `FTP-IMAGE.ENABLE ON` on the device
2. Configure `FTP-IMAGE.IP-ADDRESS` to FileZilla instance on lab machine
3. Run one scan (any trigger source)
4. Check image file: if dimensions ≈ 2448×2048 → **full-frame confirmed**; if ≈ 200–600px → ROI crop (low value, duplicates codes.xml)

---

## Image Format and Size Control

These are **Camera-category DMCC keys** shared with `IMAGE.SEND`:

### IMAGE.FORMAT
| Value | Format | Notes |
|---|---|---|
| 0 | JPEG | Default |
| 2 | PNG | Lossless; confirmed on DM475V fw 5.6.3+ |

**Recommendation**: PNG for archival (lossless, verifier re-grading via IMAGE.LOAD will not introduce JPEG artifacts). JPEG for space-efficient daily production use. Operator-selectable via CP settings.

### IMAGE.SIZE
| Value | Resolution | Dimensions (DM475V) | File size est. |
|---|---|---|---|
| 0 | Full | 2448 × 2048 | ~5 MB (PNG) / ~300–800 KB (JPEG) |
| 1 | 1/4 | 1224 × 1024 | ~1.2 MB (PNG) / ~80–200 KB (JPEG) |
| 2 | 1/16 | 612 × 512 | ~0.3 MB (PNG) |
| 3 | 1/64 | 306 × 256 | ~80 KB (PNG) |

**Important**: `IMAGE.SIZE` affects `IMAGE.SEND` AND (hypothesis) `FTP-IMAGE`. Setting
`IMAGE.SIZE=0` for full-frame FTP archival means `IMAGE.SEND` also returns full-frame.
This is desirable — consistent image resolution across all acquisition paths.

---

## DMCC Key Reference — Complete FTP-IMAGE Inventory

| Key | SET/GET | Description |
|---|---|---|
| `FTP-IMAGE.ENABLE` | SET/GET | `ON` or `OFF` — enable/disable per-scan FTP push |
| `FTP-IMAGE.IP-ADDRESS` | SET/GET | IP address of FTP/SFTP server |
| `FTP-IMAGE.PORT` | SET/GET | Server port (default 21 for FTP, 22 for SFTP) |
| `FTP-IMAGE.USER-NAME` | SET/GET | Login username |
| `FTP-IMAGE.PASSWORD` | SET/GET | Login password |
| `FTP-IMAGE.SERVER-TYPE` | SET/GET | `0`=FTP, `2`=SFTP. SFTP confirmed: DM380/390/580/590/8700. DM475V: **TBD** |
| `FTP-IMAGE.FILE-NAME` | SET/GET | Static filename for uploaded file |
| `FTP-IMAGE.CUSTOM-FILE-NAME` | SET/GET | `ON`=device generates unique name per scan; `OFF`=FTP server names it |
| `FTP-IMAGE.FILE-NAME-USE-SCRIPT` | SET/GET | `ON`=use `FTP-IMAGE.SERVER-PATH-SCRIPT` for name generation |
| `FTP-IMAGE.SERVER-PATH` | SET/GET | Directory path on the server (e.g. `/images/pips-lbl/`) |
| `FTP-IMAGE.SERVER-PATH-SCRIPT` | SET/GET | Script for dynamic path generation (SET loads script from host; GET returns current script) |
| `FTP-IMAGE.SERVER-PATH-SCRIPT-ERROR` | GET | Error message from last path-script execution |
| `FTP-IMAGE.MAX-APPEND` | SET/GET | Maximum auto-increment suffix value for filename uniqueness |
| `FTP-IMAGE.MAX-APPEND-START-VALUE` | SET/GET | Initial value for auto-increment suffix |
| `FTP-IMAGE.IDLE-LIMIT` | SET/GET | Maximum number of idle connections after data transfer |
| `FTP-IMAGE.IDLE-TIME` | SET/GET | Maximum duration of keepalive between transfers |
| `FTP-IMAGE.SERVER-FINGERPRINT` | SET/GET | SFTP host key fingerprint for server identity verification |

**Platforms**: all Ethernet readers. DM8000 wireless readers with Ethernet base station also supported.

---

## Replacement Mapping: DMST Data Logging → FTP-IMAGE

The DMST "Data Logging → Decoded Images" tab fields (screenshot 2026-06-24) map to:

| DMST Field | FTP-IMAGE Equivalent |
|---|---|
| Path (`C:\Users\Administrator\Documents\DM Reports & Decoded Images`) | `FTP-IMAGE.IP-ADDRESS` + `FTP-IMAGE.SERVER-PATH` |
| Filename Prefix (`VSNB-DM50-2604-00-MASTER-90`) | `FTP-IMAGE.FILE-NAME` or `FTP-IMAGE.FILE-NAME-USE-SCRIPT` |
| Image resolution (Full/Quarter/etc.) | `IMAGE.SIZE` DMCC key |
| Image format (PNG/BMP) | `IMAGE.FORMAT` DMCC key |
| Include Overlay Graphics (SVG) | **No equivalent** — SVG overlay is DMST rendering only |

**Key advantages of FTP over DMST filesystem save:**
1. **Sticky** — all FTP-IMAGE keys persist in NVRAM; survive DMST restarts and power cycles
2. **DMST-independent** — fires for ANY trigger source including CP software trigger
3. **Remotely configurable** — CP can set all keys via raw TCP DMCC on connect
4. **Network delivery** — image lands directly on a server share, no local-path dependency

---

## CP Integration Plan

### Phase 1 — Probe (immediate, one scan)
Confirm full-frame content via FileZilla test. See "Probe to confirm" above.

### Phase 2 — On Connect (if confirmed full-frame)
CP reads current FTP-IMAGE state:
```
GET FTP-IMAGE.ENABLE
GET FTP-IMAGE.IP-ADDRESS
GET FTP-IMAGE.PORT
GET IMAGE.SIZE
GET IMAGE.FORMAT
```
If operator has configured FTP-IMAGE in DMST, CP respects those settings.
If not configured, CP can optionally configure its own FTP-IMAGE target.

### Phase 3 — CP-Managed FTP Server (future)
CP hosts a lightweight FTP listener (Windows `FtpWebRequest` or similar). Device is
configured to push images to `127.0.0.1:21` or a local CP-managed socket. Images
arrive directly in CP memory — no external FileZilla required. More robust for
single-machine lab setups.

### Phase 4 — Image Format Policy (operator-configurable)
CP Settings → Image Archival:
- `IMAGE.FORMAT`: JPEG / PNG (default PNG for archival)
- `IMAGE.SIZE`: Full (default) / Quarter / 1/16
- Applied via DMCC on session start; restored to prior values on session end

---

## Relationship to Existing Image Stack

| Level | Name | Source | Status |
|---|---|---|---|
| L1 | Barcode crop | `JpegImageBase64` in codes.xml push XML | **COMPLETE** — always JPEG, ROI crop |
| L2 | ROI crop | `IMAGE.SEND` + software crop to `r.image.RoI` rectangle | **SCAFFOLDED** |
| **L3** | **Full frame** | **FTP-IMAGE (hypothesis)** | **PROBE NEEDED** |
| L3-alt | Full frame (SDK) | `DataManSystem.GetResultImage()` | D4 scope, SDK dependency |

If FTP-IMAGE probe confirms full-frame: L3-alt via SDK becomes unnecessary for
archival purposes. SDK `GetResultImage()` remains available for real-time display use cases.

---

*Cross-references: `WORKING-NOTES.md` §FTP-IMAGE, `firmware-confirmed-facts.md` §11, `image-capture-pipeline.md`, `wireshark-protocol-analysis.md`*
