---
name: VTCCP Working Rules
description: Canonical rules for this project — engagement style, build discipline, investigation methodology.
---

# VTCCP Working Rules — Canonical

## ★ RULE 1: DO NOT BUILD WITHOUT ASKING

Never write or push code unless the user has explicitly said to implement something.
Planning, notes, and analysis are the default output. Code only on explicit instruction.
This rule overrides everything else. It applies even when a fix seems obvious.

**Why:** The user has been burned by slap-dash implementations that assumed too much
and created new bugs. Trust must be rebuilt through careful, confirmed, deliberate steps.

## ★ RULE 2: ASSUME LESS, PROVE MORE

Do not infer device behaviour from DMCC documentation alone. The device's actual
behaviour in this environment (fw 6.1.16_sr4, DMST TruCheck 2025, DM475V hardware)
has repeatedly diverged from documentation assumptions.

**Methodology:**
- State what is CONFIRMED (observed in Wireshark, Debug logs, or user-reported outcomes)
- State what is ASSUMED (from docs, inference, or prior sessions)
- Label all assumptions explicitly and mark them for device verification before acting on them
- One variable at a time. One probe at a time. Confirm result before moving to the next.

**Why:** Prior sessions produced SET TRIGGER.TYPE=1 (Presentation) based on an assumption
that turned out to be wrong. The Presentation mode theory caused post-scan looping.
SDK XmlResultArrived was assumed to work — it doesn't for external triggers.
These errors cost real debugging time.

## ★ RULE 3: ONE VARIABLE AT A TIME

When investigating device behaviour, isolate one variable per test. Do not change
TRIGGER.TYPE AND the TCP port AND the session mode simultaneously. Change one thing,
observe the result, record it, then move to the next.

**Why:** With multiple changes in flight, it is impossible to attribute observed behaviour
to a specific cause. The user has explicitly requested this discipline.

## RULE 4: SET ASIDE PARKED ISSUES CLEANLY

When the user says to set aside an issue, note it as PARKED with a one-line description,
stop analyzing it, and do not revisit it until the user reopens it.

**Currently parked (as of 2026-05-29):**
- GS1 ]d1 vs ]d2 `<F1>` formatter issue — user will re-demonstrate the failure;
  do not draw conclusions about build state or timing until they do.

## RULE 5: AutoPoll MODE IS IGNORED

`ScanMode.AutoPoll` exists in the codebase. It is not a current priority and the user
has no recollection of requesting it. Treat it as non-existent for all planning and
discussion purposes until the user explicitly reactivates it.

## RULE 6: DOCUMENT OBSERVED FACTS, NOT ASSUMPTIONS

For Manual vs Push mode behaviour: document only what has been directly observed —
what DMST's TC window actually does in each combination of states (DMST open, DMST
closed, DMST active, DMST grayed out, VTCCP Manual, VTCCP Push). Do not fill in blanks
with inferences from code reading alone.

## RULE 7: QUIT BUTTON IS PLANNED (NOT YET BUILT)

A Quit button with an active-session warning dialog is confirmed as a planned UI feature.
It must not be built until the trigger reset / cleanup work is stable. It is downstream.
The X window-close path (Closing event) must do the same cleanup as the Quit button.
