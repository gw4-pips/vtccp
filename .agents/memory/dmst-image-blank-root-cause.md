---
name: DMST TC panel image blank — root cause and fix
description: Why DMST's TC panel image disappears post-scan when VTCCP is connected, and the confirmed fix.
---

## Root cause

The Cognex DataMan SDK's `Connect()` internally writes `DATA.IMAGE-TYPE` (and other
communication params) to values that suppress image delivery from the device's result
channel, then persists them via an internal `COM.DMCC-SAVE`. Those NVRAM values
survive for the duration of the VTCCP session and cause every post-scan result
delivered to DMST to arrive without image data — blanking the TC panel.

Key confirmed facts:
- `DATA.RESULT-TYPE=513` and `DATA.RESULT-ALWAYSSEND=513` are FACTORY DEFAULTS
  (confirmed via CONFIG.DEFAULT + GET). They are NOT the cause.
- `DATA.IMAGE-TYPE` is the key the SDK corrupts. Its factory default is unknown
  but COM.DMCC-RESET restores it correctly.
- The problem is ONLY present when VTCCP is connected. DMST alone: always fine.
- The live camera feed before a scan is unaffected (LIVEIMG is separate).
  Only the post-scan result image in the TC panel disappears.

## Fix (implemented in DataManSdkClient.ConnectAsync)

`SendDmccRestoreAsync` (COM.DMCC-RESET + COM.DMCC-SAVE) already existed but was
called only PRE-connect — immediately overwritten by the SDK's own Connect().

Fix: added a POST-connect call to `SendDmccRestoreAsync` that runs AFTER
`_system.Connect()` completes, undoing the SDK's NVRAM damage before any scan.
Safe because VTCCP uses HttpEventSubscriber for results, not the SDK result channel.

Three-phase restore sequence in DataManSdkClient:
  pre-connect  restore  →  _system.Connect()  →  post-connect restore

**Why:** pre-connect cleans up damage from the previous VTCCP session (crash recovery).
Post-connect cleans up damage from the current SDK Connect() call.

## What NOT to do

- Do NOT call CONFIG.DEFAULT to fix this — DATA.RESULT-TYPE=513 is factory default.
- Do NOT set DATA.RESULT-TYPE or DATA.RESULT-ALWAYSSEND — they are already correct.
- COM.DMCC-RESET alone (without COM.DMCC-SAVE) only fixes the current session;
  bad NVRAM value survives and DMST reloads it on next connection.
