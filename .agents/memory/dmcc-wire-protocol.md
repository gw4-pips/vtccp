---
name: DMCC raw TCP wire protocol
description: Port, command format, ACK format, and trigger sequence for raw DMCC on DM475V fw 6.1.16_sr4
---

## Port — CRITICAL

- **Port 23** — raw DMCC text interface (Telnet/DMCC). This is where `||>COMMAND\r\n` works. **CONFIRMED WORKING 2026-05-28.**
- **Port 44444** — DataMan SDK/HTTP port. Only the Cognex SDK's own binary session protocol works here. Bare TCP with `||>COMMAND\r\n` is silently ignored — device does not recognise it as a DMCC session at all, returns zero bytes for every command.

**Why:** Port 44444 multiplexes SDK binary sessions and HTTP event subscription (GET /events?enable). It does NOT accept raw DMCC text commands from new TCP connections. Port 23 is the classic DMCC Telnet interface.

## Command format

```
||>COMMAND\r\n
```

- `||>` is the wire header (DmccCommand.WireHeader)
- Commands are plain ASCII text
- TRIGGER ON (not TRIGGER, not TRIGGER 1)

## Session setup sequence — CONFIRMED WORKING 2026-05-28

1. Connect TCP to **port 23**
2. Wait briefly for banner (200 ms) — DM475V at 10.10.10.7 sends **NO banner** on port 23
3. Send `||>SET COM.DMCC-RESPONSE 2\r\n` — switch to Extended mode (no ACK for this line; device was in Silent mode, mode takes effect immediately)
4. Send `||>TRIGGER ON\r\n`
5. Read ACK: `||:::2[0]\r\n` = success (11 bytes)

## ACK format — port 23 (fw 6.1.16_sr4)

Port-23 connections return ACKs in this form:
```
||:::2[0]\r\n
```
NOT the classic `||[0]\r\n`. The `:::2` is a session/mode prefix specific to this firmware/connection type. The status code is always the **rightmost** `[N]` on the line.

**DmccResponse.Parse updated 2026-05-28**: uses `LastIndexOf(']')` + `LastIndexOf('[', rb)` so it handles both `||[0]` and `||:::2[0]` correctly. Match condition is `StartsWith("||")` — not `StartsWith("||[")`.

## TRIGGER.TYPE — no manipulation needed

TRIGGER ON works at TRIGGER.TYPE=0 (Single/external) — do NOT change TRIGGER.TYPE before firing. The device accepts software TRIGGER ON in its normal idle state. Earlier code that set TRIGGER.TYPE 1 before triggering was unnecessary and wrong.

## Result delivery

The scan result does NOT come back on the port-23 DMCC connection. It arrives asynchronously via the HTTP subscriber (GET /events?enable on port 44444 → PUT /codes.xml). The DMCC connection only needs to fire the trigger and read the ACK, then can be closed.

## Response modes

- Mode 0 (Silent, default): no ACK for any command — zero bytes returned
- Mode 2 (Extended): every command returns `||...[status]\r\n`
  - `[0]` = OK
  - `[101]` = invalid command
  - `[102]` = invalid parameter
  - `[104]` = parameter rejected due to reader state

## Constants (DmccCommand.cs)

- `WireHeader` = `"||>"`
- `SetDmccResponseExtended` = `"SET COM.DMCC-RESPONSE 2"`
- `TriggerOn` = `"TRIGGER ON"`
- `TriggerOff` = `"TRIGGER OFF"`
