# DMST Live View Protocol Analysis
**Cognex DataMan DM475V — Firmware 6.1.16_sr4**
**Capture date**: 2026-06-09  |  **Analyst**: VCCS VTCCP Project  |  **Method**: Wireshark TCP stream analysis

---

## Summary

We wanted to understand how Cognex DataMan Software Tool (DMST) drives the live camera view
panel at approximately 3–5 fps without triggering a full TruCheck barcode verification scan
on each frame. A Wireshark capture of traffic between the DMST host PC (10.10.10.19) and the
DM475V verifier (10.10.10.7) during live view operation reveals the complete mechanism.

**Bottom line**: DMST polls a proprietary HTTP image endpoint (`GET /svg_image.img`) on the
device's existing port 44444 connection. The device returns a ~106 KB image body on each
successful response. The image data is **AES-encrypted** using a key embedded in the DMST
binary — not transmitted over the wire. This makes the endpoint inaccessible to third-party
clients without either reverse-engineering the DMST binary or obtaining the key from Cognex.

---

## Capture Statistics

| Metric | Value |
|---|---|
| Capture file | `475V_trigger_test_2026-06-09_2219.pcapng` |
| Total packets | 6,088 |
| Capture duration | ~25.7 seconds |
| Device IP | 10.10.10.7 |
| PC IP | 10.10.10.19 |
| Protocol port | 44444 (TCP, same port as DMCC and HTTP result push) |

---

## Live View Protocol Detail

### Request

DMST sends a plain HTTP/1.1 GET request to the device on the existing port 44444 connection,
approximately once every **289 ms** (3.5 fps observed; up to 4.8 fps in other sessions):

```
GET /svg_image.img HTTP/1.1
Date: Wed, 10 Jun 2026 02:18:04 GMT
X-Peer: 37916227
```

- **90 requests** issued over a 25.7-second live view session
- The `Date:` header is updated on every request (prevents caching)
- `X-Peer` appears to be a session or peer identifier assigned at connection time
- No authentication header, no DMCC command, no trigger issued

### Response — "Not Ready" (camera busy or warming up)

```
HTTP/1.1 500 Internal Server Error
Content-Length: 0
Cache-Control: no-cache
Pragma: no-cache
Connection: Keep-Alive
Server: DM475/6.1.16 (DeviceID=50)
```

- **33 of 82 responses** were empty 500s (40%) — device signals "no new frame yet"
- Content-Length: 0, no body
- DMST discards these and re-polls immediately

### Response — Image Frame Delivered

```
HTTP/1.1 500 Internal Server Error
Content-Length: 106144
Content-Encoding: aes
Cache-Control: no-cache
Pragma: no-cache
Connection: Keep-Alive
Server: DM475/6.1.16 (DeviceID=50)

[106,144 bytes of AES-encrypted binary data]
```

- **49 of 82 responses** contained image data
- Frame size: min 106,140 / max 106,572 / avg **106,281 bytes** (~104 KB) — remarkably consistent
- `Content-Encoding: aes` — body is AES-encrypted; not JPEG or any standard image format
- The device returns **HTTP 500** for both "ready" and "not ready" — 500 is used as the
  normal success code for this endpoint (not a real error)
- `Connection: Keep-Alive` — the same TCP connection is reused for all polls

### Key observations

1. **No barcode decode or TruCheck grading occurs** — this is a pure camera frame delivery
   path, completely separate from the verification pipeline.

2. **The device reuses port 44444** — the same port carries three simultaneous protocols:
   raw DMCC command sessions, HTTP result push subscriptions (`GET /events?enable`), and
   this HTTP image polling endpoint. The device disambiguates by the request format.

3. **AES-encrypted body, consistent 106 KB size** — at 106 KB per frame this is likely a
   downsampled or cropped JPEG encrypted before transmission. The consistent size suggests
   a fixed-dimension output regardless of scene content (i.e., the JPEG quality or dimensions
   are fixed, not the entropy). The AES key is not present anywhere in the capture.

4. **The `vs.cfg` config sync file** (pushed by the device to DMST as `PUT /vs.cfg`) also
   carries `Content-Encoding: aes` — the same key likely covers both. Neither is readable
   without the key.

---

## Protocol Map — Port 44444 Multiplexing

All of the following protocols share a single TCP port on the DM475V:

| Request format | Direction | Purpose |
|---|---|---|
| Raw DMCC XML (`\|\|>GET ...`) | PC → device | DMCC command sessions (config, trigger, IMAGE.SEND) |
| `GET /events?enable HTTP/1.1` | PC → device | HTTP result push subscription |
| `PUT /status.xml` | device → PC | Periodic telemetry (~1/sec) |
| `PUT /codes.xml` | device → PC | Verification result delivery |
| `PUT /pcm_report.html` | device → PC | Full HTML grade report |
| `PUT /vs.cfg` | device → PC | AES-encrypted config sync |
| **`GET /svg_image.img`** | **PC → device** | **Live view frame polling** |

---

## Implications for Third-Party Clients

| Question | Answer |
|---|---|
| Can a third party poll `/svg_image.img`? | Yes — the HTTP GET is trivial to replicate |
| Is the AES decryption key available? | **No** — not on the wire; embedded in DMST binary |
| Is there an unencrypted image path? | Unknown — `IMAGE.SEND` (DMCC, port 44444 or 23) delivers unencrypted JPEG after a trigger |
| Is `LIVEIMG.SEND` (DMCC reference) used? | **No** — not observed in DMST capture; tested manually, confirmed dead |
| What fps can a third party achieve? | 1.5–2.5 fps via `TRIGGER ON` + `IMAGE.SEND` (full TruCheck scan per frame); DMST achieves 3.5 fps via encrypted polling with no decode cost |

---

## Alternative: DMCC IMAGE.SEND Path (currently implemented in VTCCP)

Without the AES key, the only viable live view path for a third-party application is:

1. Open a raw TCP connection to port 44444 (or 23)
2. Issue `TRIGGER ON` (software trigger) — starts a scan
3. Wait for first byte of DMCC result XML (scan-done signal, ~200–450 ms depending on grade)
4. Open a second connection and issue `IMAGE.SEND` — returns the last-captured JPEG (unencrypted, ROI crop)
5. Display the frame, then repeat

This yields **1.5–2.5 fps** (grade F to grade A respectively). The frame is the same
barcode-crop JPEG shown in the DMST verification panel — firmware-processed, correct
aspect ratio, unencrypted. A full TruCheck grade is computed on every frame (unavoidable
with this path), which is why it cannot match DMST's 3.5 fps encrypted preview.

---

## Questions for Cognex Engineering

If the intent is to match DMST live view performance in a third-party application, the
following would need to come from Cognex:

1. **Is the AES key for `/svg_image.img` available to authorized SDK users or OEM partners?**
2. **Is there a documented, unencrypted image streaming endpoint** (e.g. a DMCC `LIVEIMG.SEND`
   path) that operates without a full TruCheck scan per frame?
3. ~~**Does `LIVEIMG.MODE=2` + `LIVEIMG.SEND` deliver unencrypted frames?**~~ **CONFIRMED
   DEAD — 2026-06-10.** Both `LIVEIMG.MODE=2` and `LIVEIMG.MODE=3` + `LIVEIMG.SEND 0 1 85`
   tested directly via raw TCP on port 23. Device accepts commands silently, returns
   **0 bytes** on both modes. Not functional on fw 6.1.16_sr4 / DM475V hardware.

---

*Document generated from Wireshark pcapng analysis of live DM475V traffic.
Raw capture archived at `vtccp/architecture/gui-reference/wireshark-dmst-full-capture.txt`.*
