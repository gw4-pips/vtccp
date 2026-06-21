# VTCCP HTTP Event Subscriber — Architecture Specification

**Version 1.0 — 2026-06-20**

---

## Purpose

The HTTP Event Subscriber (`HttpEventSubscriber`) is VTCCP's path to **complete DMST
independence** for result delivery. It eliminates the need for DMST to be open or
involved in any part of the scan result workflow while simultaneously delivering the
HTML supplemental fields that push XML cannot provide.

---

## Protocol Background

The Cognex DataMan device multiplexes two protocols on **port 44444**:

| Connection type | How the device identifies it | Protocol |
|---|---|---|
| Raw DMCC | First bytes are DMCC XML or text | DMCC session |
| HTTP subscription | First line is `GET /events?enable HTTP/1.1` | HTTP push stream |

A single TCP connection carries either DMCC traffic OR the HTTP event stream — not both.
The DMCC port (23) and DMCC-over-44444 are unaffected by an active HTTP subscription.

**Confirmed** from Wireshark capture (2026-05-25): packet 46 TCP header confirms
`Dst Port: 44444`; DMST's GET /events?enable subscription and VTCCP's DMCC sessions
coexist without conflict.

---

## Result Delivery Flow

```
  Device                                    VTCCP
    │                                          │
    │◄── GET /events?enable HTTP/1.1 ──────────│  TCP connect + subscribe
    │──► 204 No Content ───────────────────────►│
    │                                          │
    │  ... scan occurs (any trigger source) ...│
    │                                          │
    │──► PUT /pcm_report.html ─────────────────►│  HTML report (~131–202 KB)
    │    (arrives BEFORE codes.xml)             │  → DmstHtmlScraper.ParseHtml()
    │                                           │  → _pendingHtml cache
    │──► PUT /codes.xml ───────────────────────►│  Full result data (~9.4–202 KB)
    │    (verify: 202 KB; monitor: 9.4 KB)      │  → DmstResultParser.Parse()
    │                                           │  + MergeAndValidate(_pendingHtml)
    │                                           │  → ResultReceived raised
    │──► PUT /status.xml ─────────────────────►│  Telemetry (~4.6 KB, ~1/sec)
    │    (continuous, between scans)            │  → ignored / no handler
    │                                           │
    │──► PUT /vs.cfg ─────────────────────────►│  Config sync (AES-encrypted)
    │    (on config change)                     │  → ignored / not decodable
```

---

## Endpoints — Detail

### `GET /events?enable`
- **Direction**: VTCCP → Device
- **Response**: `204 No Content`
- **Effect**: Enrolls this TCP connection as a subscriber. Device will push all
  subsequent scan results and status updates to this connection.
- **Keep-alive**: Connection must remain open. The device pushes asynchronously
  whenever a scan occurs, regardless of trigger source.

### `PUT /pcm_report.html`
- **Direction**: Device → VTCCP
- **Size**: 131–202 KB (varies by symbology and result complexity)
- **Timing**: Arrives **before** `PUT /codes.xml` for each scan
- **Content**: Complete HTML verification report — the same document that DMST saves
  to disk and that `DmstHtmlScraper` reads from the filesystem
- **Handler**: `DmstHtmlScraper.ParseHtml()` → stored in `_pendingHtml` cache, merged
  into the scan record when `codes.xml` arrives
- **Fields delivered** (not in push XML / codes.xml):
  - `ECLevel` (QR: "M", "Q", "H", "L") — QR only; DM has no selectable EC level
  - `DataMaskPattern` (QR: 0–7) — QR only; DM has no data masking
  - `ECI` value (e.g., "000003" for Latin-1) — QR and others
  - `ImagePolarity` ("Black on white" / "White on black") — all symbologies
  - `DataCodewords` (integer) — exact value, vs. C# table lookup fallback
  - `ErrorCorrectionBudget` (integer) — exact value, vs. C# table lookup fallback
  - `EncodedCharacters` (integer) — exact value; **unresolvable from push XML on fw 6.1.16_sr4**

### `PUT /codes.xml`
- **Direction**: Device → VTCCP
- **Size**: ~9.4 KB (monitor scans) / ~202 KB (full TruCheck verification scans)
- **`origin` discriminator**:
  - `"monitor"` — background monitoring scan; minimal data, no TruCheck grades
  - `"common"` — full triggered TruCheck verification; complete grade data
- **Content**: Contains `<trucheck_verificaiton_result>` block (note firmware spelling)
  with the complete General Characteristics table, `<FormalGrade>`, `<OpticalVariant>`,
  all grade fields — same data as DMST push XML
- **Handler**: `DmstResultParser.Parse()` + `MergeAndValidate()` + merge from
  `_pendingHtml` → raises `ResultReceived`
- **Firmware misspellings** (do not correct — these are the literal XML element names):
  - `<trucheck_verificaiton_result>` (not "verification")
  - `<VerificaitonTime>` (not "Verification")
  - `<CanidateEvaluationTime>` (not "Candidate")

### `PUT /status.xml`
- **Direction**: Device → VTCCP
- **Size**: ~4.6 KB
- **Rate**: ~1/sec (continuous telemetry between scans)
- **Handler**: None — VTCCP ignores this stream (no useful scan data)

### `PUT /vs.cfg`
- **Direction**: Device → VTCCP
- **Content**: AES-encrypted configuration blob — not decodable by third parties
- **Handler**: None — VTCCP ignores this

---

## C# Implementation

### Class
`vtccp/DeviceInterface/HttpEventSubscriber.cs`

### Entry points in `DeviceSession`

```csharp
// Start
await session.StartHttpSubscriberAsync(sessionContext, ct);

// Stop
await session.StopHttpSubscriberAsync();
```

### Wiring in `SessionViewModel`

The HTTP subscriber is started for **both** session modes:
- **Push mode**: subscriber is the primary result path; `StartPushListenerAsync` not used
- **Manual / AutoPoll mode**: subscriber runs in parallel; enriches records with HTML
  fields that the poll path cannot provide

### `DmstHtmlScraper` coexistence

When the HTTP subscriber is running, `DmstHtmlScraper` (the filesystem watcher) is
**harmless alongside it**. Both can run concurrently:
- Scraper watches `{Documents}\{DeviceName}\CodeQuality\*.html`
- Subscriber receives `PUT /pcm_report.html` inline

In practice, both deliver the same HTML content. For CP-triggered scans DMST does not
save to disk, so the subscriber is the only source of HTML fields. For DMST-triggered
scans, both paths deliver the file — the first to arrive wins.

---

## Trigger Independence

The HTTP subscriber receives scan results **regardless of what triggered the scan**:
- CP software trigger (DMCC `TRIGGER ON`)
- DMST user trigger (button click / hotkey)
- External hardware trigger
- AutoPoll / motion detection

This is the key DMST-independence advantage over the filesystem scraper path, which
only delivers files for DMST-triggered scans.

---

## Fields Not Delivered by HTTP Subscriber

Even with the HTTP subscriber active, these fields remain inaccessible from the device:

| Field | Status |
|---|---|
| `ECLevel` (DM) | N/A — DM ECC200 has no selectable error correction level |
| `DataMaskPattern` (DM) | N/A — DM has no data masking |
| Full-resolution image (L3) | Requires Cognex DataMan SDK (`GetResultImage()`) — D4 scope |
| Live view stream | `LIVEIMG.SEND` dead; `svg_image.img` AES-encrypted |

---

## Port Conflict Avoidance

VTCCP maintains two separate TCP connections to the device:
1. **DMCC session** (port 23 or port 44444, raw text/XML) — for `TRIGGER ON`, `GET`, `SET`, `IMAGE.SEND`
2. **HTTP subscription** (port 44444, HTTP GET) — for `PUT /codes.xml`, `PUT /pcm_report.html`

These are independent sockets. The device's port-44444 dispatcher routes by connection
intent: DMCC XML connections and HTTP subscription connections do not interfere.

---

## Comparison: Three Result-Delivery Paths

| Path | DMST required? | HTML fields? | Trigger-independent? | Status |
|---|---|---|---|---|
| DMCC push listener (`StartPushListenerAsync`) | No | No (push XML only) | Yes | Implemented |
| Filesystem scraper (`DmstHtmlScraper`) | Yes (must be open) | Yes | No (DMST-triggered only) | Implemented |
| HTTP subscriber (`StartHttpSubscriberAsync`) | **No** | **Yes** | **Yes** | **Implemented** |

The HTTP subscriber supersedes the other two paths for result delivery when available.
The filesystem scraper remains relevant only for L0 PNG access and DMST-triggered
session continuity.

---

*Cross-references: `wireshark-protocol-analysis.md` (raw packet analysis, §6.3),
`firmware-confirmed-facts.md` §10 (Wireshark findings), `image-capture-pipeline.md`.*
