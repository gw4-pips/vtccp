---
name: GoToTags E310 Protocol & Integration
description: Wire protocol, CRC details, and VTCCP integration pattern for the GoToTags Desktop E310 UHF RFID reader (replaces MTI RU-824-100).
---

## Hardware
- SKU TDLP3LCFPP, $199.93 qty-1, store.gototags.com
- Chipset: Impinj E310 (Indy R2000 family) — industry-standard silicon
- FTDI USB VCP driver: CDM212364_Setup.zip (saved at vtccp/references/gototags-e310/)
- Appears as COM port after FTDI driver install — SerialPort architecture is correct
- 860–960 MHz, up to 1 metre range, FCC certified

## Protocol (confirmed from PDF rev 5-30-23)
- Baud: 115 200 8N1 (factory default, no change needed)
- Frame: `FF | DataLen (1B) | CmdCode (1B) | Data[DataLen] | CRC_Hi | CRC_Lo`
- CRC-16/CCITT: init=0xFFFF, poly=0x1021, computed over bytes[1..frameLen-3] (skips 0xFF header AND CRC bytes), stored big-endian
- Device auto-boots to APP layer at power-on; no boot-layer handshake needed

## Key commands (APP layer)
- `0x21` Single Tag Inventory: Data = Timeout(2B big-endian ms) + Option(1B)
  - Option=0x00 → no filter, no metadata; returns EPC + TagCRC only
  - Response Data: Status(2B) + Option(1B) + EPC(DataLen-5 bytes) + TagCRC(2B)
  - Status 0x0000 = tag found; any other = no tag or error
- `0x22` Synchronous Inventory (batch, with 0x29 Get Tag Buffer to retrieve)
- `0xAA48`/`0xAA49` Asynchronous Inventory (push mode, highest performance)

## VTCCP implementation
- GoToTagsE310Protocol.cs — frame builder + CRC + EPC extractor
- GoToTagsE310Reader.cs — IEpcReader via 0x21 loop (150ms slices, EPC de-dup)
- EpcReaderFactory.CreateGoToTagsE310() is the primary factory method
- MtiLlcsEpcReader marked superseded

## Hardware TODOs (verify on first plug-in)
- CRC round-trip: send one 0x21, confirm reader acknowledges cleanly
- No-tag status code: capture actual two-byte status when 0x21 times out empty

**Why:** MTI RU-824-100 is USB HID (not VCP) — SerialPort won't work for it. E310 uses FTDI VCP, making SerialPort the correct approach and the existing IEpcReader interface a natural fit.
