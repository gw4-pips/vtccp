---
name: DM475V HTTP result delivery
description: How TC verification results reach VTCCP — the only working path on fw 6.1.16_sr4
---

## The only path for TC verification results

This document is DataMan/DMST-specific. Do not apply its device HTTP stream,
port, or subscriber assumptions to Webscan TruChecks; Webscan devices are
USB-connected and have a separate result path.

Port 44444 multiplexes DMCC (raw TCP) and HTTP on the same port. The device distinguishes by opening bytes.

`XmlResultArrived` on the SDK connection ONLY fires when the device is in Presentation mode (TRIGGER.TYPE=1, autonomous self-scanning). For any externally-triggered scan (DMST Verify button, physical reader button, raw TCP TRIGGER), results go to the HTTP channel ONLY — the SDK connection is deaf.

## HTTP subscription

```
GET /events?enable HTTP/1.1\r\n
X-Peer: 0\r\n
\r\n
```
Device replies `204 No Content`. Keep connection open. Device then pushes:
- `PUT /status.xml` (~1/sec, heartbeat — ignore)
- `PUT /vs.cfg` (AES-encrypted — ignore)
- `PUT /pcm_report.html` — HTML report, arrives BEFORE codes.xml for same scan
- `PUT /codes.xml` — push XML in base64 inside `<general><full_string encoding="base64">`
- `PUT /svg_image.img` (~5MB SVG — ignore)

Filter on `origin="common"` in codes.xml root element. `origin="monitor"` = DMST live-view frame, no TruCheck data.

## HTML-report generation prerequisite

DMST **must be running and connected to the specific DataMan verifier** for that verifier to generate its TruCheck HTML report. This is a device/report-generation prerequisite, not merely a client-side convenience.

- The TruCheck window may remain open or be closed after DMST connects; neither state prevents report generation.
- Connecting DMST to a V-series reader opens TruCheck in a blank startup state. Operators may close that window if desired.
- Without DMST's connection, do not expect a correlated HTML report. VTCCP must continue to fail closed for TruCheck-only PDF data rather than substituting XML or inferred values.

**Why:** User-confirmed operational finding. The two attached VCCS reports show that, once a DMST-connected verifier generated correlated HTML, GS1 Data Format fields appeared only when that feature was selected; when it was not selected, the HTML omitted them and the PDF correctly reported them as unavailable.

## Implementation

`HttpEventSubscriber.cs` is fully implemented. `DeviceSession.StartHttpSubscriberAsync()` starts it. Called after `ConnectAsync()` in `SessionViewModel.cs` with `ResultReceived` wired to `OnPushRecord`.

**Why:** Confirmed 2026-05-28. Physical button press, DMST Verify button, and raw TCP TRIGGER all produce results only on the HTTP channel. SDK XmlResultArrived never fires for these.

## Concurrency with DMST and DM TC — CANONICAL (refined by user)

DMST (browser UI / device config) and DM TC (DataMan TruCheck application) do NOT block VTCCP. All three can run simultaneously on the same PC.

- Port 44444 is the **device's** HTTP server — it accepts multiple concurrent clients.
- VTCCP's `GET /events?enable` HTTP subscription is device-side; the device pushes to it regardless of what PC applications are open.
- A scan triggered from within DMST or DM TC generates a device event that lands in VTCCP's HTTP subscriber just like any other scan.
- Operators CAN leave DMST and DM TC open while VTCCP is running and recording; no need to close either.
- **Exception:** if TruCheck is in **LIVE mode**, VTCCP cannot trigger a scan. LIVE mode must not have been activated, or the user must cancel/exit LIVE mode before VTCCP attempts a trigger.

**Why:** The user confirmed this empirically — "with VTCCP open a scan triggered from within either app lands in VTCCP." This is expected: port 44444 is not held exclusively by any one PC app.

## CP software trigger (Path B) — UNRESOLVED

Raw TCP `TRIGGER\r\n` confirmed NOT causing a device scan (HTTP subscriber would have caught it). SDK throws InvalidParameterException for both `TRIGGER` and `TRIGGER 1`. Root cause unknown — may require correct DMCC parameter form or HTTP-channel trigger.
