---
name: DM475V Trigger Type Ground Truth
description: Confirmed TRIGGER.TYPE enum, device idle state, SDK behaviour, and VTCCP design decisions for fw 6.1.16_sr4.
---

# DM475V Trigger Type — Device-Confirmed Ground Truth

## TRIGGER.TYPE enum — fw 6.1.16_sr4 (from DMCC Reference ZIP)

Source: `idp10154189968.htm` inside `DataMan_Command_Reference_6.1.16_sr4.zip`

| Value | Name |
|---|---|
| 0 | Single (external) |
| 1 | Presentation (internal) |
| 2 | Manual (button) |
| 3 | Burst (external) |
| 4 | Self (internal) |
| 5 | Continuous (external) |

There is NO "Single (software/DMCC)" entry. The old code comment `0=Continuous, 1=Single, 2=External` was completely wrong.

## Device idle state

- Device idles at TRIGGER.TYPE=**0** (Single external) — confirmed in every exception log
- DMST never changes TRIGGER.TYPE during Go Live or Verify — it stays at 0 the whole time
- DMST's continuous scanning is a **programmatic loop** firing software TRIGGER commands, not a firmware mode change
- The loop fires regardless of object presence (even scans an empty base plate)
- **Presentation mode (TRIGGER.TYPE=1) is never a factor in this environment** — established as canonical

## SDK behaviour for GET TRIGGER.TYPE

- GET TRIGGER.TYPE works fine via SDK (`_client.SendAsync`) — returns `'0'` or `'1'`
- The prior claim that "SDK PayLoad is empty on 6.1.16_sr4 for GET TRIGGER.TYPE" was WRONG
- No raw TCP bypass is needed for this command

## What VTCCP was doing (prior to 2026-05-28)

- `ConnectAsync`: read TRIGGER.TYPE (stored as `_originalTriggerType`), then immediately SET to 1
- `DisconnectAsync` / `RebootAndDisconnectAsync`: restored to `_originalTriggerType`
- The SET was an AI-assumed "Single software mode" — no documentation or user instruction supported it
- TRIGGER.TYPE=1 is Presentation (internal), not any kind of "Single software"

## Current state (as of 2026-05-28 — commit 43d58dc)

- `GET TRIGGER.TYPE` read is **kept** — logged as diagnostic, stored in `_originalTriggerType`
- `SET TRIGGER.TYPE 1` is **commented out** — probe to determine if it caused post-scan looping
- Both restore blocks (DisconnectAsync + RebootAndDisconnectAsync) are **commented out** in sync
- Raw TCP TRIGGER command fires a scan regardless of TRIGGER.TYPE — confirmed working at TRIGGER.TYPE=0

## Why raw TCP is still needed for TRIGGER itself

- `_system.SendCommand("TRIGGER 1")` → always throws `InvalidParameterException` (fw/SDK version mismatch)
- `_system.SendCommand("TRIGGER")` → also throws `InvalidParameterException`
- Raw TCP `||>TRIGGER\r\n` (second TCP connection) → works; result delivered via SDK's `XmlResultArrived` event (not the raw socket)
- This is unrelated to TRIGGER.TYPE — it is a firmware 6.1.16_sr4 / SDK v25 protocol mismatch

## Open probe (2026-05-28)

Does removing `SET TRIGGER.TYPE 1` change the undesirable post-scan looping behaviour?
Result expected from user test within ~45 min of 2026-05-28 session.
