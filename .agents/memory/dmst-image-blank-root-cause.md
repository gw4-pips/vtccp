---
name: DMST TC panel image blank — confirmed root cause and fix
description: Why DMST's TC panel image disappears post-scan when VTCCP was connected, the confirmed fix, and all false leads.
---

## Root cause — CONFIRMED 2026-05-31

**NVRAM corruption** from a prior SDK `COM.DMCC-SAVE` call (via `SetResultTypes()`).
The exact parameter was not identified. `CONFIG.DEFAULT + CONFIG.SAVE + REBOOT` cleared it.

This was a **one-time fix**. The image has been working since.

## CRITICAL — correct known-good values (device-confirmed 2026-05-31)

| Parameter | Known-good value | Notes |
|---|---|---|
| `LIVEIMG.MODE` | **0** | NOT 2 — TC panel image works with 0 |
| `DATA.IMAGE-TYPE` | 0 | Correct SDK default |
| `DATA.RESULT-TYPE` | 513 | DM475V factory default |
| `TRIGGER.TYPE` | 0 | External single trigger |
| `IMAGE.SIZE` | 1 | Confirmed from snapshot |

## All prior theories were WRONG

### Wrong theory 1: LIVEIMG.MODE = 0 is the cause
- **FALSE.** The known-good snapshot (2026-05-31 224936) shows LIVEIMG.MODE=0 with the
  TC panel image fully working. LIVEIMG.MODE=0 is the CORRECT value.
- All code that set LIVEIMG.MODE=2 has been removed from DataManSdkClient.cs.
- `SendDmccRestoreAsync` was removed entirely (no callers, wrong logic).

### Wrong theory 2: DMST holds port 23 exclusively
- **FALSE.** VTCCP's TRIGGER ON fires via a raw TCP connection to port 23, and this
  works correctly while DMST is open. Both VTCCP (trigger) and DMST can use port 23
  simultaneously. No conflict was observed.
- OCE visible at VS Output line 80 is the SDK's own internal exception from
  `_system.Connect()` — it is benign and always present. NOT from port 23 failures.

### Wrong theory 3: SDK Connect() sets LIVEIMG.MODE to 0
- **FALSE on fw 6.1.16_sr4.** No such effect was observed in the known-good state.
  The NVRAM corruption came from `SetResultTypes()` triggering `COM.DMCC-SAVE`, not
  from the SDK's Connect() overwriting LIVEIMG.MODE.

## Fix implemented

1. Removed `SetResultTypes()` call from DataManSdkClient (prevents future corruption).
2. Removed all LIVEIMG.MODE restore code from `ConnectAsync` and `DisconnectAsync`.
3. Removed `SendDmccRestoreAsync` (dead code with wrong logic).
4. Applied `CONFIG.DEFAULT + CONFIG.SAVE + REBOOT` on the device (one-time fix).

## Recovery if image disappears again

1. Telnet to port 23
2. `CONFIG.DEFAULT`
3. `CONFIG.SAVE`
4. `REBOOT`
5. Restore custom config (TRIGGER.TYPE, TruCheck settings, etc.)

Do NOT set LIVEIMG.MODE to 2.

**Why:** `SetResultTypes()` was the only call that triggered `COM.DMCC-SAVE`. With it
removed, the SDK no longer writes to NVRAM on connect. Future connections will not
corrupt the device config.
