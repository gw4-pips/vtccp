---
name: Trigger Reset & DMST Recovery Plan
description: Plan to restore TRIGGER.TYPE to live-feed mode on VTCCP exit; restore point; Manual vs Push cleanup paths.
---

# Trigger Reset & DMST Recovery — Working Plan

## Restore Point

**Commit**: `f474fa7` (HEAD at time of plan creation, 2026-05-29)
**Tag**: "AS-IS restore point — trigger investigation not yet begun"
All work below this point is deferred until the user restarts this investigation.

---

## The Problem (as stated by user, 2026-05-29)

After VTCCP runs a CP Trigger Scan (Push mode), the DMST TC window does not
recover to "live" mode after the scan completes. The user describes the live mode
as "120X/min live feed." DMST TC stays stuck on the result screen.

---

## What is CONFIRMED

- TRIGGER.TYPE=0 (Single/external) is the device's confirmed idle state during
  normal DMST TC operation (from trigger-type-ground-truth.md)
- VTCCP currently does NOT change TRIGGER.TYPE on connect or disconnect
  (SET and both restore blocks are commented out)
- Raw TCP TRIGGER ON fires a single scan regardless of TRIGGER.TYPE
- After a scan, the device returns to its configured TRIGGER.TYPE behaviour
- DMST's "continuous" scanning is a programmatic software-trigger loop, not a
  firmware mode change (Presentation mode is NOT in use here)

## What is NOT CONFIRMED / OPEN PROBES

1. **What TRIGGER.TYPE value corresponds to "120X/min live feed mode"?**
   - CONFIRMED that it is NOT TRIGGER.TYPE=1 (Presentation) — that theory was wrong
   - Current best hypothesis: it is still TRIGGER.TYPE=0 (Single), with DMST
     continuously issuing software triggers to produce the live feed effect
   - NEEDS WIRESHARK: capture DMST Go Live → Verify → Go Live on port 23 to see
     exactly what SET TRIGGER.TYPE values (if any) DMST sends at each transition
   - Do NOT assume any specific value until this capture is done

2. **Why does DMST TC not recover after a VTCCP Push scan?**
   - Hypothesis A: VTCCP's HTTP subscriber connection (GET /events?enable on port 44444)
     is still open while DMST TC tries to re-subscribe, causing a conflict
   - Hypothesis B: The brief port-23 TRIGGER ON connection from VTCCP leaves some
     device state that prevents DMST's trigger loop from restarting
   - Hypothesis C: DMST detects a competing subscriber and changes its own behaviour
   - NONE of these are confirmed. Need observed behaviour to select between them.

3. **What does "not recovering" look like exactly?**
   - Does the TC window freeze on the result? Go gray? Show an error?
   - Is it recoverable by clicking something in DMST, or requires VTCCP to stop?
   - User to describe or screenshot the exact stuck state

---

## Investigation Sequence (ONE VARIABLE AT A TIME)

### Step 0 — Baseline (before any code changes)
Run DMST TC Verify scan with VTCCP NOT running. Confirm TC window recovers normally.
This establishes the known-good baseline.

### Step 1 — Push mode trigger, then observe
Run VTCCP in Push mode. Fire ONE CP trigger scan. Observe exactly what DMST TC does
afterward. Is it the trigger that causes the problem, or the connection itself?

### Step 2 — HTTP subscriber vs no subscriber
If VTCCP is running in Push mode but the HTTP subscriber is NOT started (disconnect
HttpEventSubscriber from the Push session path temporarily), does DMST TC recover?
This isolates whether the HTTP subscriber connection is the cause.

### Step 3 — Wireshark capture
Capture port 23 + port 44444 during: VTCCP Push scan → DMST TC stuck → DMST recovery
attempt (or failure). Look for: SET TRIGGER.TYPE values, connection conflicts, error codes.

### Step 4 — Trigger type restore
Only AFTER confirming what the live-feed TRIGGER.TYPE value is: implement the restore.

---

## Code Locations (for when implementation is approved)

| File | Location | Pending change |
|---|---|---|
| `DeviceSession.DisconnectAsync()` | Lines 172–188 | Uncomment restore block |
| `DeviceSession.RebootAndDisconnectAsync()` | Lines 210–220 | Uncomment restore block |
| `DmccCommand.cs` | Near GetTriggerType | Add SetTriggerTypeLiveMode constant (value TBD from probe) |
| `SessionViewModel` | Push mode stop/exit handler | Brief port-23 connection to SET TRIGGER.TYPE on session stop |
| `MainWindow.xaml.cs` (or equivalent) | Window Closing event | Run full cleanup (Stop → disconnect → trigger restore) before allowing close |

---

## Manual vs Push Mode — Observed Behaviour (TO BE FILLED IN)

The following table is BLANK intentionally. Fill in only from direct observation,
not from code reading or inference.

| Scenario | DMST state | VTCCP mode | What actually happens |
|---|---|---|---|
| VTCCP starts, Connect pressed | DMST closed | Manual | ? |
| VTCCP starts, Connect pressed | DMST open | Manual | ? |
| VTCCP starts, Connect pressed | DMST open, TC active | Manual | ? |
| CP Trigger Scan fired | DMST closed | Manual | ? |
| CP Trigger Scan fired | DMST open | Push | ? |
| VTCCP Stop pressed | DMST closed | Manual | ? |
| VTCCP Stop pressed | DMST open | Push | ? — TC window stuck? |
| VTCCP X closed (no Stop) | DMST open | Push | ? |

---

## Quit Button (Planned, NOT YET BUILT)

A dedicated Quit button with an active-session warning dialog is confirmed as planned.
Prerequisites: trigger reset must be stable and tested first. The Quit button and the
X window-close Closing event must both run identical cleanup. Build only after trigger
investigation is complete and user approves implementation.
