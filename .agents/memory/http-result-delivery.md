---
name: DM475V HTTP result delivery
description: How TC verification results reach VTCCP — the only working path on fw 6.1.16_sr4
---

## The only path for TC verification results

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

## Implementation

`HttpEventSubscriber.cs` is fully implemented. `DeviceSession.StartHttpSubscriberAsync()` starts it. Called after `ConnectAsync()` in `SessionViewModel.cs` with `ResultReceived` wired to `OnPushRecord`.

**Why:** Confirmed 2026-05-28. Physical button press, DMST Verify button, and raw TCP TRIGGER all produce results only on the HTTP channel. SDK XmlResultArrived never fires for these.

## Concurrency with DMST and DM TC — CANONICAL (confirmed 2026-08-18)

DMST (browser UI / device config) and DM TC (DataMan TruCheck application) do NOT block VTCCP. All three can run simultaneously on the same PC.

- Port 44444 is the **device's** HTTP server — it accepts multiple concurrent clients.
- VTCCP's `GET /events?enable` HTTP subscription is device-side; the device pushes to it regardless of what PC applications are open.
- A scan triggered from within DMST or DM TC generates a device event that lands in VTCCP's HTTP subscriber just like any other scan.
- Operators CAN use DMST for live image viewing or DM TC for one-off checks while VTCCP is running and recording. No need to close either.

**Why:** The user confirmed this empirically — "with VTCCP open a scan triggered from within either app lands in VTCCP." This is expected: port 44444 is not held exclusively by any one PC app.

## CP software trigger (Path B) — UNRESOLVED

Raw TCP `TRIGGER\r\n` confirmed NOT causing a device scan (HTTP subscriber would have caught it). SDK throws InvalidParameterException for both `TRIGGER` and `TRIGGER 1`. Root cause unknown — may require correct DMCC parameter form or HTTP-channel trigger.
