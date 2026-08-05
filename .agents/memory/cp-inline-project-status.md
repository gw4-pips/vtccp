---
name: CP Inline project status
description: Strategic decision — CP Inline abandoned for the Serview 475V-LBL inline application; Command Pilot focus returns to desktop verification tool.
---

## Decision (5 August 2026)
CP (Command Pilot) was abandoned as the approach for the UDI DM-475V-LBL **inline production** application at the Serview customer site. Reason stated as "other reasons" — not a technical finding.

**Command Pilot** continues as a desktop barcode print-quality verifier tool (its original purpose).

## Implications
- CP Inline product-concept tasks (operator panel mockup, InlineIo stub, sales deck, etc.) may be moot for the Serview engagement but have not been explicitly cancelled — do not auto-close them without the user's direction.
- InlineIo C# stub (`vtccp/InlineIo/`) and CPIPM Project Outline v1.3 (`vtccp/references/`) were completed before the decision and remain in the repo.

## DM475-866D76 (DPM unit) state as of 5 August 2026
- Factory-defaulted via DataMan Setup Tool
- Firmware: 6.1.16_tc9
- Feature keys include TCVerification — grading engine intact
- Device profile needs to be re-applied from Command Pilot (Apply to Device) before use

## Live image at DM TC level
Both the DM-475V-LBL and DM-475V-DPM units failed to show live image in DataMan's own TC interface (not a CP issue). Root cause unknown, not investigated further.
