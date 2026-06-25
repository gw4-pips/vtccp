# VTCCP — DMST-Independent Operation: Implementation Plan

**Version 1 — 2026-06-25**

---

## Goal

VTCCP operates the DM475V (and future DM395V) for full TruCheck verification
without requiring DMST to be running. DMST may remain open on the same machine
for operator recovery access — VTCCP does not interfere with it.

---

## What already exists (do not rebuild)

| Component | File | Status |
|---|---|---|
| HTTP events subscriber spec | `http-subscriber-spec.md` | ✓ Fully designed |
| FTP-IMAGE architecture spec | `ftp-image-architecture.md` | ✓ Designed; one probe pending |
| Push XML parser | `DmstResultParser.cs` | ✓ Complete |
| HTML report parser | `DmstHtmlScraper.ParseHtml()` | ✓ Complete |
| DMCC raw TCP (port 23) | `DeviceSession.SendRawDmccAsync` | ✓ Working |
| Result model + Excel writer | `VerificationRecord`, `ExcelWriter` | ✓ Complete |

---

## DMST coexistence

The device supports multiple simultaneous HTTP event subscribers on port 44444.
VTCCP's `GET /events?enable` connection coexists with DMST's identical connection —
confirmed from Wireshark: DMST's HTTP subscription and VTCCP's DMCC sessions
run in parallel without conflict.

**DMST stays open.** VTCCP does not need to close or control it.

---

## Build sequence

### Phase 1 — HTTP Events Channel (result delivery without DMST)

**What it replaces**: `DmstHtmlScraper` filesystem reader + DMCC push listener

**New component**: `HttpEventsChannel`
- Opens one persistent TCP connection to device port 44444
- Sends `GET /events?enable HTTP/1.1` with `X-Peer` + `Date` headers
- Receives and demultiplexes inbound HTTP PUTs:
  - `PUT /codes.xml` → feed to `DmstResultParser.Parse()`
  - `PUT /pcm_report.html` → feed to `DmstHtmlScraper.ParseHtml()`
  - `PUT /svg_image.img` → raise `LiveImageReceived` event (SVG bytes)
  - `PUT /status.xml`, `PUT /vs.cfg` → discard
- Merges codes.xml + pcm_report.html → raises `ResultReceived`

**Dependency**: `pcm_report.html` arrives BEFORE `codes.xml` — cache HTML until
codes.xml arrives, then merge. Already the pattern in `DmstHtmlScraper`.

**DMCC stays active** in parallel for GET/SET parameter operations.

**Trigger source for Phase 1**: DMCC `TRIGGER ON` (port 23) — unchanged.
Results flow via events channel regardless of trigger source.

---

### Phase 2 — HTTP Command Channel (DMST-independent trigger)

**What it replaces**: DMCC `TRIGGER ON` over port 23 (optional — DMCC trigger still works)

**New component**: `HttpCommandChannel`
- Opens one persistent TCP connection to device port 44444
- Sends custom HTTP verbs — all confirmed from Wireshark:

```
RESUME /        → 200 OK     (session open)
ISALIVE /       → 204        (keepalive)

GET /monitormode?enable=true  → 204    (enter Sleep / Go Live)
GET /monitormode?enable=false → 204    (exit Sleep)
TRIGGER /on     → 204                  (fire verification scan)
TRIGGER /off    → 204                  (release)
GET /monitormode?enable=true  → 204    (return to Sleep)
```

**Cancel**: `GET /monitormode?enable=false` only — no trigger.

**Note**: `GET /device_info.xml` returns 401 on fw 6.1.16_sr4 — do not attempt.
`GET /vs.cfg` and `GET /parameters.xml` return AES-encrypted blobs — fetch but discard.

**Phase 2 dependency**: Phase 1 must be working first (results must arrive before
disconnecting from the existing trigger path).

---

### Phase 3 — FTP-IMAGE server (full-frame image archival)

**What it provides**: Full-sensor-frame image (2448×2048 at IMAGE.SIZE=0) per scan,
delivered by the device via FTP push after each verification.

**Architecture**:
- VTCCP hosts a minimal embedded FTP server (`VtccpFtpListener`) on a configurable port
- On connect to device, VTCCP writes (via DMCC):
  ```
  SET FTP-IMAGE.ENABLE ON
  SET FTP-IMAGE.IP-ADDRESS {vtccp_host_ip}
  SET FTP-IMAGE.SERVER-PORT {port}
  SET FTP-IMAGE.SERVER-LOGIN {user}
  SET FTP-IMAGE.SERVER-PASSWORD {pass}
  SET FTP-IMAGE.FILE-NAME {job}_{scan_index}
  ```
- Device pushes one image file per scan to VTCCP's FTP server
- VTCCP correlates by scan index or timestamp → attaches to `VerificationRecord`

**Open probe (do before building)**: Enable FTP-IMAGE pointing at FileZilla, run
one scan, check dimensions. If 2448×2048 → full-frame confirmed → build.
If ROI crop → no value over codes.xml JpegImageBase64 → skip.

**FTP server options** (in order of preference):
1. `FluentFTP` or `WinSCP .NET assembly` — proven .NET FTP server libs
2. Minimal hand-rolled passive-mode FTP (STOR command only — device only writes)

**Phase 3 dependency**: Phase 1 complete. FTP-IMAGE probe confirmed.

---

### Phase 4 — Live SVG image display (Go Live monitor view)

**What it provides**: Real-time annotated scan image from `PUT /svg_image.img`
on the events channel — same image shown in DMST's verification panel.

- Phase 1 already receives the SVG bytes and raises `LiveImageReceived`
- Phase 4 is purely UI: render SVG in a WPF `WebView2` or `SvgViewbox` control
- No new protocol work required — data already flowing from Phase 1

**Phase 4 dependency**: Phase 1 complete + UI design decision on SVG renderer.

---

## What remains DMCC-only (does not change)

- All GET/SET device configuration (aperture, lighting, grading standard, etc.)
- `IMAGE.SEND` if live view JPEG is ever needed (separate from FTP-IMAGE)
- Calibration commands
- `TRIGGER.TYPE` read/restore on connect/disconnect

---

## Open questions before starting Phase 1

1. **Multiple HTTP subscribers confirmed?** — DMST open + VTCCP both subscribed.
   One quick Wireshark test: open VTCCP's events connection while DMST is already
   connected. Does VTCCP receive PUT events? (Almost certainly yes — but confirm.)

2. **FTP-IMAGE full-frame probe** — one scan with FileZilla. Unblocks Phase 3 go/no-go.

3. **Verification enable/disable Wireshark capture** — may reveal additional command
   channel verbs needed for Phase 2 completeness.
