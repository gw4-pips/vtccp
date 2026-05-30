---
name: DMST blank image — COM.DMCC-SAVE required after COM.DMCC-RESET
description: COM.DMCC-RESET alone does not fix DMST blank image panel; must follow with COM.DMCC-SAVE to persist defaults to NVRAM
---

## The rule

Any time VTCCP sends `COM.DMCC-RESET` on port 23, it **must** immediately follow with `COM.DMCC-SAVE` on the same connection (with a ~120ms pause between them).

**Why:** COM.DMCC-RESET resets DMCC communication parameters (DATA.IMAGE-TYPE, DATA.RESULT-TYPE, etc.) to firmware defaults for the current session. It does NOT write to NVRAM. The Cognex SDK's own `Connect()` internally calls an equivalent of `SetResultTypes()` + `COM.DMCC-SAVE`, writing a bad DATA.IMAGE-TYPE value to NVRAM. DMST loads from NVRAM on every reconnect. So COM.DMCC-RESET without COM.DMCC-SAVE leaves the bad NVRAM value in place → DMST image panel stays blank indefinitely, surviving device restarts and power cycles.

**How to apply:** Use `SendDmccRestoreAsync()` in `DataManSdkClient.cs`, which sends both commands in sequence. Called at:
- Connect time (before SDK Connect) — fixes any pre-existing bad NVRAM state immediately on Start Session
- Disconnect time (after SDK Disconnect) — undoes whatever the SDK wrote during the session

## Confirmed facts

- COM.DMCC-RESET WAS reaching the device and being ACKed on port 23 (debug output showed `||:::2[0]\r\n` response) — the command executed but NVRAM was not updated
- The bad NVRAM state survived: TC verifications via DMST showed no image even without VTCCP running, and across restarts — confirming it was a persisted (NVRAM) problem, not session-level
- COM.DMCC-SAVE as a solo command via raw TCP to port 23 has not been independently confirmed on this device/firmware — the assumption that it persists to NVRAM follows from the DMCC Reference spec; hardware test will confirm
