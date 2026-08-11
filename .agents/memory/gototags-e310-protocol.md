---
name: GoToTags E310 Protocol & Integration
description: SUPERSEDED. E310 evaluated and rejected. AsReader ASR-P35U is the canonical RFID reader.
---

## Status: SUPERSEDED

The GoToTags Desktop E310 UHF RFID reader has been **evaluated and rejected**.
The **AsReader ASR-P35U** is the canonical RFID hardware for VCCS RFID FlexWedge Pro
and all Command Pilot RFID integration.

## Why rejected
User evaluated in the field and rejected in favour of the ASR-P35U. No further detail
on record — assume the P35U was superior in range, reliability, or form factor.

## Reference
The VCCS RFID FlexWedge Pro (Python, in GS1 Resolver Replit project) is the
authoritative protocol reference for the ASR-P35U. Transfer brief at:
`vtccp/references/asr-p35u/TRANSFER-BRIEF.md`

## Archived protocol notes (E310 — do not use)
- Chipset: Impinj E310 (Indy R2000 family)
- FTDI USB VCP, 115200 8N1
- Frame: FF | DataLen | CmdCode | Data | CRC_Hi | CRC_Lo
- CRC-16/CCITT: init=0xFFFF, poly=0x1021, over bytes[1..frameLen-3], big-endian
- 0x21 Single Tag Inventory command
- Code: GoToTagsE310Protocol.cs, GoToTagsE310Reader.cs (both now obsolete)
