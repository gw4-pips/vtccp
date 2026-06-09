# VTCCP Command Pilot — Tools Map
**What each tool does, what data it owns, and what only it can provide**  
**Last updated**: 2026-06-09

---

## Overview

Command Pilot connects to a Cognex DataMan DM475V (or 395V) via Ethernet and orchestrates
five distinct tools to collect a complete verification record from a single scan. No single
tool can provide everything — understanding which tool owns which data is essential for both
debugging and feature planning.

```
┌─────────────────────────────────────────────────────────────────────────┐
│                         VTCCP Command Pilot                             │
│                                                                         │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌───────────┐  │
│  │  Cognex SDK  │  │ Raw TCP DMCC │  │ HTTP Pub/Sub │  │   DMST    │  │
│  │ (port 44444) │  │  (port  23)  │  │ (port 44444) │  │  Scraper  │  │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘  └─────┬─────┘  │
│         │                 │                  │                │         │
│    connect/info      trigger / live     result XML +      HTML report   │
│    session mgmt         view              HTML push          fields      │
│                                                                         │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │                    GS1 Syntax Engine (local)                     │   │
│  │             AI parsing · HRI · Digital Link ↔ AI                │   │
│  └─────────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## Tool 1 — Cognex DataMan SDK (`DataManSdkClient`, port 44444)

**What it is**: The official Cognex proprietary SDK (`Cognex.DataMan.SDK.PC.dll`),
wrapped in `DataManSdkClient.cs`. Connects to the device on port 44444 using the SDK's
own binary session protocol.

**What it does for VTCCP**:
- Establishes and maintains the primary authenticated TCP session to the device
- Queries device metadata at connect time: `FirmwareVersion`, `DeviceInfo.Type` (model),
  serial number, device name, calibration date
- Provides the SDK `DataManSystem` object which keeps the port-44444 session alive

**What it CANNOT do** (important limitations discovered through testing):
- **Cannot trigger**: `SendCommand("TRIGGER")` / `SendCommand("TRIGGER 1")` throw
  `InvalidParameterException` — the SDK validates command names internally and rejects
  TRIGGER before it ever reaches the wire. Software triggers must go via raw TCP (Tool 2).
- **Cannot run arbitrary DMCC**: Many commands not on the SDK's internal whitelist throw
  `InvalidCommandException` and never reach the device. Use raw TCP (Tool 2) for anything
  the SDK rejects.
- **`XmlResultArrived` event is dead** for external triggers — the event was designed for
  SDK-initiated triggers. For VTCCP's external-trigger workflow (DMST controls the trigger),
  results never arrive via this event. Result delivery is via HTTP subscriber (Tool 3).
- **`SetResultTypes()` must never be called** — it sends `DATA.IMAGE-TYPE` /
  `DATA.RESULT-TYPE` DMCC commands and persists them via `COM.DMCC-SAVE`, which strips
  image data from the result channel and blanks DMST's image panel until device reboot.

**Source**: `vtccp/DeviceInterface/Dmcc/DataManSdkClient.cs`

---

## Tool 2 — Raw TCP DMCC (`DmccCommand`, `LiveFeedClient`, port 23)

**What it is**: Direct TCP socket connection to the device's Telnet/DMCC interface on
**port 23**. No Cognex library involved — raw byte I/O.

**★ MANDATORY PREFIX**: Every command sent on port 23 MUST begin with `||>`:
```
||>COMMAND\r\n   ← correct
COMMAND\r\n      ← silently ignored — device returns nothing, no error
```
This has caused multiple lost debugging sessions. Never omit `||>`.

**What it does for VTCCP**:
- **Software trigger**: `||>TRIGGER ON\r\n` — fires a scan without touching the physical trigger
- **IMAGE.SEND** (live view): `||>IMAGE.SEND\r\n` → returns a JPEG of the current camera view
- **GET/SET any DMCC key**: UPC/EAN supplemental mode, LIVEIMG.MODE, connection response mode, etc.
- **Session response mode**: default on a fresh port-23 connection is silent (mode 0 — no ACK).
  Must send `||>SET COM.DMCC-RESPONSE 2\r\n` before any command that needs a response.

**ACK format** (fw 6.1.16_sr4):
```
||:::2[0]\r\n    ← success (status code is always the rightmost [N])
||:::2[101]\r\n  ← invalid command
||:::2[102]\r\n  ← invalid parameter
||:::2[104]\r\n  ← parameter rejected (reader state)
```

**What it CANNOT provide**: Scan result data — the device pushes scan results to the HTTP
channel (Tool 3), not back on the port-23 connection.

**Source**: `vtccp/DeviceInterface/Dmcc/DmccCommand.cs`, `LiveFeedClient.cs`

---

## Tool 3 — HTTP Event Subscriber (`HttpEventSubscriber`, `DmstListener`, port 44444)

**What it is**: An HTTP long-poll subscription to the device's result-push channel.
Opening `GET /events?enable HTTP/1.1` on port 44444 registers VTCCP as a subscriber.
After each scan, the device pushes results via HTTP PUT to all subscribers.

**What the device pushes** (in order per scan):
| Endpoint | Size | Timing | Content |
|---|---|---|---|
| `PUT /pcm_report.html` | 131–202 KB | Before codes.xml | Full HTML verification report (same file DMST saves to disk) |
| `PUT /codes.xml` | 9.4 KB (monitor) / 202 KB (verify) | After HTML | Full result XML including trucheck block |
| `PUT /status.xml` | ~4.6 KB | ~1/sec always | Device telemetry |
| `PUT /vs.cfg` | varies | periodic | AES-encrypted config sync (unreadable) |

**`origin` discriminator** in codes.xml: `"monitor"` = background scan (ignore); `"common"` = full TruCheck verification result (process).

**What it provides** (from `codes.xml`):
- Complete grading result: `<trucheck_verificaiton_result>` block (firmware typo: "verificaiton")
- Formal grade string: `<FormalGrade>1.0/16/660/45Q</FormalGrade>`
- `<OpticalVariant>DM475V</OpticalVariant>` — exact model string
- All ISO quality parameters (numeric grades, percentages)
- General Characteristics: **correct** EncodedCharacters, DataCodewords, ErrorCorrectionBudget
  (values here are correct where the push script XML gives wrong/empty values)
- Decoded barcode data, symbology, scan timestamps
- JPEG image (base64-encoded) — the same ROI crop shown in DMST's verification panel
- **GS1 Application Standard validation block** (when GS1 parser is enabled on device):
  - `<ApplicationStandard>` — e.g. "GS1", "Custom"
  - `<ApplicationPass>` — e.g. "Pass", "Fail (Data Format)", "Fail (Quality)"
  - `<ApplicationPassReason>` — e.g. "Data Format", "Quality"
  - Confirmed in codes.xml (Wireshark capture) and in HTML report
  - Design rule: `ApplicationPass` is **informational only** — `OverallGrade` (ISO) is always
    the operative grade; ApplicationPass never overrides it. Wired: `DmstResultParser` →
    `VerificationRecord.ApplicationStandard/Pass/PassReason` → Excel.

**This is the PRIMARY result delivery path.** All scan records flow through here.

**Source**: `vtccp/DeviceInterface/Dmst/HttpEventSubscriber.cs`, `DmstListener.cs`,
`DmstResultParser.cs`

---

## Tool 4 — DMST HTML Scraper (`DmstHtmlScraper`)

**What it is**: A parser that reads the `pcm_report.html` file that DMST saves to disk at
`{Documents}\{DeviceName}\CodeQuality\{timestamp}.html`. The HTTP subscriber (Tool 3)
also receives the identical HTML body via `PUT /pcm_report.html`, so the scraper's data
source is available with or without DMST running.

**Why it exists**: Several fields that appear in DMST's UI and HTML report are
**not accessible via push XML on fw 6.1.16_sr4**. The probe campaign confirmed this
definitively (v1.33 scans #14/#15). The HTML report is the only source for these fields.

**Fields ONLY available from the HTML report** (not in push XML):
| Field | Applies to | Status |
|---|---|---|
| `ImagePolarity` | All symbologies | ✓ implemented — "Black on white" / "White on black" |
| `ECLevel` (Error Correction Level) | **QR only** | ParseHtml() extension pending — "M","Q","H","L" confirmed in HTML |
| `DataMaskPattern` | **QR only** | ParseHtml() extension pending — value "2" confirmed in HTML |
| `ECI` value | QR + others | ParseHtml() extension pending — "000003" confirmed in HTML |

**Fields where HTML gives the CORRECT value vs push XML's wrong value**:
| Field | Push XML (wrong) | HTML (correct) |
|---|---|---|
| EncodedCharacters | 33 (DM) / 39 (QR) | 38 (DM) / 36 (QR) |
| DataCodewords | empty (q.symbols=null) | 32 (DM 16×36) |
| ErrorCorrectionBudget | empty (q.symbols=null) | 24 (DM 16×36) |

**Source**: `vtccp/DeviceInterface/Dmst/DmstHtmlScraper.cs`, `DmstHtmlReport.cs`

---

## Tool 5 — GS1 Barcode Syntax Engine (`GS1Encoder`, local library)

**What it is**: The GS1 reference implementation library (`gs1encoders.dll`, native C with
C# P/Invoke wrapper). Version 1.4.0, vendored at `vtccp/lib/gs1-syntax-engine/`.
Reference standard: GS1 Digital Link Standard URI Syntax v1.6.0 (Mar 2025), on file.

**What it does for VTCCP**:
- **AI element string validation**: checks that `(01)`, `(10)`, `(11)` etc. values conform
  to GS1 data syntax rules (format, length, check digits, mutual exclusions)
- **HRI generation**: produces `(AI) value` Human-Readable Interpretation lines from
  any GS1 data input
- **GS1 Digital Link URI parsing**: when `DecodedData` starts with `https://` or `http://`,
  set `encoder.DataStr = DecodedData` → library extracts the AI element string and populates
  `AIdataStr`, `HRI[]` identically to traditional AI input
- **DL URI generation**: `GetDLuri(stem)` converts AI data to a GS1 Digital Link URI
- `PermitZeroSuppressedGTINinDLuris`, `PermitUnknownAIs` — handle legacy and non-standard inputs
- `DLignoredQueryParams` — reports non-AI URI query params that were present but ignored

**DataMan CT comparison**: CT uses the GS1 Syntax Engine only for Digital Link and its own
internal parser for traditional AI strings. VTCCP should use the library for both — no
legacy AI parser to protect.

**Current usage in VTCCP**: `BarcodeDataFormatter` uses it for `<F1>` GS1 formatting.
Digital Link detection/routing not yet implemented (future work — detect `https://` prefix
in `DecodedData`, route to `encoder.DataStr`).

**Source**: `vtccp/lib/gs1-syntax-engine/src/GS1Encoder.cs`  
**Reference**: `vtccp/architecture/gs1-digital-link.md`

---

## Data Source Summary — What Lives Where

| Data field | Tool 1 SDK | Tool 2 Raw TCP | Tool 3 codes.xml | Tool 4 HTML | Tool 5 GS1 Lib |
|---|:---:|:---:|:---:|:---:|:---:|
| Device model / firmware / serial | ✓ | | | | |
| Calibration date | ✓ | | | | |
| Software trigger | | ✓ | | | |
| Live view JPEG (IMAGE.SEND) | | ✓ | | | |
| DMCC GET/SET (supp mode, etc.) | | ✓ | | | |
| All ISO grade parameters | | | ✓ | | |
| Formal grade string | | | ✓ | | |
| Decoded barcode data (raw) | | | ✓ | | |
| Verification JPEG (ROI crop) | | | ✓ | | |
| Scan timestamps | | | ✓ | | |
| GS1 ApplicationStandard/Pass/Reason | | | ✓ | also in HTML | |
| **EncodedCharacters (correct)** | | | **HTML only** | ✓ | |
| **DataCodewords (correct)** | | | **HTML only** | ✓ | |
| **ErrorCorrectionBudget (correct)** | | | **HTML only** | ✓ | |
| **ImagePolarity** | | | | ✓ | |
| **ECLevel (QR only)** | | | | ✓ pending | |
| **DataMaskPattern (QR only)** | | | | ✓ pending | |
| **ECI value** | | | | ✓ pending | |
| AI validation + HRI | | | | | ✓ |
| Digital Link ↔ AI conversion | | | | | ✓ |

**Bold rows = fields with no alternative source.** These are permanently unresolvable
from push XML on fw 6.1.16_sr4 — the HTML report (Tool 4) is the only path.

---

## Session Flow — How the Tools Fire Together

```
Connect
  → Tool 1 (SDK): ConnectAsync → device info, session established
  → Tool 2 (raw TCP): SET COM.DMCC-RESPONSE 2, GET TRIGGER.TYPE, supplemental mode

Scan triggered (external — operator presses DMST trigger or physical button)
  → Tool 3 (HTTP): PUT /pcm_report.html arrives first → DmstHtmlScraper queued
  → Tool 3 (HTTP): PUT /codes.xml arrives → DmstResultParser → VerificationRecord
  → Tool 4 (scraper): TryMergeAsync fills HTML-only fields into the same record
  → Tool 5 (GS1 lib): BarcodeDataFormatter formats decoded data + AI parsing

Live View (Phase I — in progress)
  → Tool 2 (raw TCP): TRIGGER ON → drain scan result XML → IMAGE.SEND → JPEG

Disconnect
  → Tool 1 (SDK): session closed; TRIGGER.TYPE restored
```

---

## Key Invariants

- **TRIGGER.TYPE = 0 always.** Never change it. VTCCP fires software triggers via `TRIGGER ON` 
  at TRIGGER.TYPE=0 without modifying this setting.
- **`SetResultTypes()` never called.** Would corrupt NVRAM and blank DMST image panel.
- **`COM.DMCC-SAVE` never called** for anything except changes the user explicitly intends to persist.
- **`||>` prefix on every port-23 command.** Bare commands are silently ignored.
- **LIVEIMG.MODE = 0.** Setting to 2 previously caused NVRAM corruption on this firmware.
