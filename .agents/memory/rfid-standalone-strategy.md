---
name: RFID Standalone Product Strategy
description: Strategic and architectural decisions for the RFID cross-validation feature as a standalone verifier-agnostic product
---

# RFID Standalone Product Strategy

## The rule
Raw EPC hex string is the integration contract. Any reader that outputs hex works.
This makes the feature brand- and model-agnostic by design.

**Why:** Hex is the lowest common denominator output of virtually every UHF RFID reader
regardless of brand, generation, or SDK. Designing to this contract means hardware is
swappable without touching the parser/validator/report stack.

## How to apply
- `IEpcReader` acquisition layer delivers hex via `MtiLlcsEpcReader` (Phase 0 concrete impl)
- Layer 2 (parser/validator) never references reader brand or SDK
- Full architecture in: `vtccp/architecture/rfid-standalone-product.md`

## Canonical RFID hardware — AsReader ASR-P35U

The **AsReader ASR-P35U** is the confirmed canonical RFID reader (2026-08-06).
- GoToTags E310: evaluated and rejected (see gototags-e310-protocol.md)
- MTI RU-824-100: earlier candidate, also superseded
- Protocol reference: VCCS RFID FlexWedge Pro (Python) in GS1 Resolver Replit project
- Transfer brief at: vtccp/references/asr-p35u/TRANSFER-BRIEF.md
- C# implementation target: AsReaderP35UEpcReader.cs implementing IEpcReader

## MTI RU-824-100 — archived (superseded)
- Reader = FTDI USB chip → virtual COM port (VCP driver) — NOT a keyboard wedge, NOT raw HID
- Protocol = MTI LLCS binary packet format over serial at 115200 8N1
- Integration = `System.IO.Ports.SerialPort` only; NO native DLL dependency
- SDK native DLLs (`Transfer.dll`, `ftd2xx.dll`, `rfid.dll`) are 32-bit .NET Framework era;
  do not use them; drive the COM port directly with documented LLCS packets
- Command reference manual PDF cloned locally:
  `vtccp/references/mti-sdk/RFID_Explorer/MTI RU-824 RFID Module Command Reference Manual v3.3.pdf`
- Triggerable: LLCS has explicit Inventory Start/Stop commands confirmed in source code
- **Why:** Avoids 32-bit native DLL dependency; keeps the entire stack .NET 8 pure-managed

## Key decisions
- Standalone product targets competitive verifier accounts (Axicon, OMRON, Webscan, REA)
  without requiring Command Pilot or DataMan — fastest monetization path
- PoC (RFID ME USB reader + raw hex) is structurally the MVP standalone product
- Partition value validated against GCP Length Table: flag encoding error + state correct
  value per GS1; never block the data comparison on a partition discrepancy
- GCP Length Table: file `vtccp/data/gcp-prefix-format-list.xml` (2026-05-03, 8.7MB, 200K entries)
  Auto-update: HEAD-fetch semi-private URL on CP startup; compare XML root `date` attr; notify user
- Handheld combo readers (Zebra RFD, Honeywell) are NOT competitors: no ISO grade,
  no GCP validation, no extended reporting — different product category entirely
- Separate USB barcode scanner (Tier 1 config) is aesthetically "kludgy" but accepted
  pragmatically for initial deployment: "build what we can sell"
- Goal for Tier 2: minimal-config competitive verifier adapter (OMRON SQLite DB,
  Axicon ScanDB CSV) so RFID decode marries to verifier scan with near-zero user setup
- Excel: child RFID tab (not additional columns on 167-col schema); Scan ID is join key
- Report: inject RFID HTML block post-XSLT; XSLT templates untouched
- EPC scheme dispatch on header byte; SGTIN-96 (0x30) and SGTIN-198 (0x36) are primary
