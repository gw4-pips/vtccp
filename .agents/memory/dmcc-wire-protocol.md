---
name: DMCC raw TCP wire protocol
description: Port, command format, ACK format, and trigger sequence for raw DMCC on DM475V fw 6.1.16_sr4
---

## Port — CRITICAL

- **Port 23** — raw DMCC text interface (Telnet/DMCC). This is where `||>COMMAND\r\n` works. **CONFIRMED WORKING 2026-05-28.**
- **Port 44444** — DataMan SDK/HTTP port. Only the Cognex SDK's own binary session protocol works here. Bare TCP with `||>COMMAND\r\n` is silently ignored — device does not recognise it as a DMCC session at all, returns zero bytes for every command.

**Why:** Port 44444 multiplexes SDK binary sessions and HTTP event subscription (GET /events?enable). It does NOT accept raw DMCC text commands from new TCP connections. Port 23 is the classic DMCC Telnet interface.

## ★★★ MANDATORY PREFIX — NEVER OMIT ★★★

**EVERY raw DMCC command on port 23 MUST begin with `||>` — no exceptions.**

```
||>COMMAND\r\n       ← CORRECT
COMMAND\r\n          ← SILENTLY IGNORED — device returns zero bytes, no error
IMAGE.SEND\r\n       ← SILENTLY IGNORED — this is how the IMAGE.SEND bug went undetected
```

The device does NOT reject bare commands with an error code — it simply does nothing.
This has caused multiple debugging sessions of false leads. The `||>` prefix is mandatory
for every single command, every time, on every connection. `DmccCommand.WireHeader = "||>"`.

**How to apply**: Before writing any DMCC send call, confirm the string starts with
`DmccCommand.WireHeader` or the literal `||>`. Never send a bare command name.

## ★★★ SET EXTENDED MODE FIRST — NEVER SKIP ★★★

**After connecting to port 23, ALWAYS send `||>SET COM.DMCC-RESPONSE 2\r\n` before any GET command.**

- Default mode is Silent (0) — device returns ZERO BYTES for every command including GETs
- Extended mode (2) — every command returns `||:::N[status]VALUE\r\n`
- No ACK is returned for the SET COM.DMCC-RESPONSE 2 line itself; mode takes effect immediately
- Failure to set Extended mode looks identical to wrong prefix — both return nothing

This cost multiple debugging sessions. Both mistakes (missing prefix, missing mode switch) produce
the exact same symptom: all params return (no response).

## Command format

```
||>COMMAND\r\n
```

- `||>` is the wire header (DmccCommand.WireHeader)
- Commands are plain ASCII text
- TRIGGER ON (not TRIGGER, not TRIGGER 1)

## Session setup sequence — CONFIRMED WORKING 2026-05-28 / 2026-08-17

1. Connect TCP to **port 23**
2. Wait briefly for banner (200 ms) — DM475V at 10.10.10.7/10.10.10.4 sends **NO banner** on port 23
3. Send `||>SET COM.DMCC-RESPONSE 2\r\n` — switch to Extended mode (no ACK for this line; mode takes effect immediately)
4. Send commands with `||>` prefix
5. Read ACK: `||:::2[0]\r\n` = success (11 bytes)

## ACK format — port 23 (fw 6.1.16_sr4 / 6.1.16_tc9)

Port-23 connections return ACKs in this form:
```
||:::N[status]VALUE\r\n
```
Examples:
- `||:::5[0]ON` — single-value response (value is on the SAME line after `[0]`)
- `||:::2[0]\r\n||:::12[0]6.1.16_tc9` — two-line response for some params
- `||:::5[0]467\r\nscript content...` — multi-line for COM.SCRIPT, DEVICE.LOG etc.

**Parse rule**: strip just `||:::\d+\[\d+\]` prefix — NOT `[^\r\n]*` (that eats the value).
`DmccResponse.Parse` uses `LastIndexOf(']')` + `LastIndexOf('[', rb)` for status code extraction.

## Full parameter dump tool

**Use `vtccp/tools/Get-DmSettings.ps1`** — 352 parameters, ~5 min, outputs timestamped .txt file.

```powershell
# Download and run:
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/gw4-pips/vtccp/main/vtccp/tools/Get-DmSettings.ps1" -OutFile "C:\Get-DmSettings.ps1"
C:\Get-DmSettings.ps1
```

Key fixes in the script (as of 2026-08-17):
- Sends `||>SET COM.DMCC-RESPONSE 2` before the parameter sweep
- Uses blocking `Read` with `ReadTimeout=500` — NOT `DataAvailable` (DataAvailable fires before data arrives on LAN, causing all values to read as (no response))
- Strips `\|\|:::\d+\[\d+\]` prefix (not the greedy `[^\r\n]*` which eats the value)

DO NOT use Windows Telnet (`telnet 10.10.10.4 23`) — it sends IAC Telnet negotiation bytes the device ignores; commands typed at the Telnet prompt are silently dropped.

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
