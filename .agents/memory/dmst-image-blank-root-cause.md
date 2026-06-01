---
name: DMST TC panel image blank — confirmed root cause and fix
description: Why DMST's TC panel image disappears post-scan when VTCCP is connected, and the confirmed fix.
---

## Root cause — CONFIRMED 2026-05-31

**`LIVEIMG.MODE = 0`** is the actual root cause. Not `DATA.IMAGE-TYPE`.

The Cognex DataMan SDK's `Connect()` sets `LIVEIMG.MODE` to `0`
("no image with result") and persists it via its internal config-save mechanism.
`LIVEIMG.MODE = 2` means "send image with each result" — the correct value for
DMST's TC panel to display the post-scan barcode image.

Confirmed by Telnet GET while VTCCP connected:
- `GET LIVEIMG.MODE` → `0`  (wrong — should be 2)
- `GET DATA.IMAGE-TYPE` → `0`  (correct — was never the cause)
- `GET DATA.RESULT-TYPE` → `513`  (factory default — never the cause)

`COM.DMCC-RESET` does NOT restore `LIVEIMG.MODE`. It only resets DMCC
session parameters. `LIVEIMG.MODE` is a CONFIG parameter — it requires
`SET LIVEIMG.MODE 2` + `CONFIG.SAVE` to fix.

## Fix (current approach — SDK-side, port 44444)

Implemented inside `Task.Run` in `DataManSdkClient.ConnectAsync`, right after
`_system.Connect()` returns. Uses the SDK's own already-open port 44444 connection:

```csharp
_system.SendCommand("SET LIVEIMG.MODE 2");
_system.SendCommand("CONFIG.SAVE");
```

SDK may throw `InvalidCommandException` if the command is not on its internal whitelist
— caught and logged. If blocked, a different approach is needed.

## CRITICAL: DMST blocks port 23

DMST holds a **persistent port 23 connection** while it is open and connected.
Any VTCCP attempt to open a second port 23 TCP connection blocks for the full
timeout and then fails silently. This is why all previous port-23 restore attempts
(fixes 2–4) had zero effect — the commands never reached the device.

Evidence: `VTCCP-Reset.ps1` explicitly says "close DMST completely so it does
not hold port 23." The OCE visible in VS Output at line 80 of every debug capture
is the 5-second `outerCts` firing when port 23 refuses a second connection.

**Port 23 based approaches will NOT work while DMST is open.**

## False leads ruled out

- `DATA.IMAGE-TYPE` — confirmed = 0 (correct) while VTCCP connected; never the cause.
- `DATA.RESULT-TYPE = 513` — factory default; not the cause.
- Port 23 restore (fixes 2–4) — blocked by DMST holding port 23; commands never reached device.
- VS Output `[VTCCP-SDK]` lines absent — because port 23 restore fails fast (OCE from outerCts)
  and VS drops the catch debug line during the heavy startup DLL loading phase.

## What COM.DMCC-RESET actually resets

Only DMCC session parameters: DATA.IMAGE-TYPE, DATA.RESULT-TYPE,
DATA.RESULT-ENCODING, DATA.RESULT-ALWAYSSEND, COM.DMCC-RESPONSE,
COM.DMCC-CHECKSUM, COM.DMCC-HEADER. LIVEIMG.MODE is NOT reset by this.
