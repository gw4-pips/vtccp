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

## Fix (implemented in DataManSdkClient.SendDmccRestoreAsync)

Four-command sequence sent on port 23 (pre-connect, post-connect, post-disconnect):

1. `COM.DMCC-RESET`  — resets DMCC session params (DATA.IMAGE-TYPE etc.)
2. `COM.DMCC-SAVE`   — persists DMCC defaults to NVRAM
3. `SET LIVEIMG.MODE 2`  — **the actual fix** — restores image delivery
4. `CONFIG.SAVE`     — persists LIVEIMG.MODE=2 to flash

**Why:** Without `CONFIG.SAVE`, the SDK's persisted `LIVEIMG.MODE=0` reloads
on the next device connection and blanks the image again.

## False leads ruled out

- `DATA.IMAGE-TYPE` — confirmed = 0 (correct) while VTCCP connected; never the cause.
- `DATA.RESULT-TYPE = 513` — factory default confirmed via CONFIG.DEFAULT. Not the cause.
- `COM.DMCC-RESET + COM.DMCC-SAVE` alone — insufficient; doesn't touch LIVEIMG.MODE.

## What COM.DMCC-RESET actually resets

Only DMCC session parameters: DATA.IMAGE-TYPE, DATA.RESULT-TYPE,
DATA.RESULT-ENCODING, DATA.RESULT-ALWAYSSEND, COM.DMCC-RESPONSE,
COM.DMCC-CHECKSUM, COM.DMCC-HEADER. Nothing else.
