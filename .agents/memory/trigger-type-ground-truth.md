---
name: DM475V Trigger Type Ground Truth
description: Confirmed trigger configuration of DM475-63530E-PIPS-Verif-Lab in its normal idle/static state, as seen in DMST immediately on connect.
---

# DM475V Trigger Type — Device-Confirmed Ground Truth

**Confirmed by user screenshot, 2026-05-27**

## Normal idle/static state (immediately upon DMST connect, before any scan)

- **Trigger type**: Single (external)
  - Type = Single: one acquisition per trigger signal
  - Source = External: hardware signal from PIPS system (NOT software/DMCC command)
- **Image panel**: gray (no scan yet)
- **ROI**: defined (visible in TC)
- **LED action on reader**: none — device is waiting for hardware trigger
- **TC display**: gray (no result)

## Motion detection

- **Motion detection is NOT checked**
- **Motion detection NEVER comes into play on this device/setup**
- Do not factor motion detection into any trigger logic, restore logic, or state analysis

## Previous incorrect theory

- "Presentation mode" theory was WRONG — device is NOT in Presentation mode
- The static/gray state is Single+External: device waiting for a hardware PIPS trigger, which never fires because no product is present

## Implications for VTCCP trigger restore

- When VTCCP sets TRIGGER.TYPE 1 (Single/software), it changes the trigger SOURCE from External to DMCC/software while keeping the type as Single
- The restore must return the device to Single+External, not just Single
- Exact DMCC string returned by GET TRIGGER.TYPE on this firmware is still unknown — may be "Single", "External", an integer, or a combined value
- Raw TCP GET TRIGGER.TYPE must be used at connect to capture the real value (SDK PayLoad is empty on 6.1.16_sr4)
- More user information is expected — wait before writing restore code
