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
- `IEpcReader` acquisition layer (HID keyboard hook OR ComPort SerialPort) delivers hex
- Layer 2 (parser/validator) never references reader brand or SDK
- Full architecture in: `vtccp/architecture/rfid-standalone-product.md`

## Key decisions
- Standalone product targets competitive verifier accounts (Axicon, OMRON, Webscan, REA)
  without requiring Command Pilot or DataMan — fastest monetization path
- PoC (RFID ME USB reader + raw hex) is structurally the MVP standalone product
- Partition value validated against GCP Length Table: flag encoding error + state correct
  value per GS1; never block the data comparison on a partition discrepancy
- GCP Length Table now password-protected (Jan 2026); PIPS/VCCS supply as reference file
- Handheld combo readers (Zebra RFD, Honeywell) are NOT competitors: no ISO grade,
  no GCP validation, no extended reporting — different product category entirely
- Separate USB barcode scanner (Tier 1 config) is aesthetically "kludgy" but accepted
  pragmatically for initial deployment: "build what we can sell"
- Goal for Tier 2: minimal-config competitive verifier adapter (OMRON SQLite DB,
  Axicon ScanDB CSV) so RFID decode marries to verifier scan with near-zero user setup
- Excel: child RFID tab (not additional columns on 167-col schema); Scan ID is join key
- Report: inject RFID HTML block post-XSLT; XSLT templates untouched
- EPC scheme dispatch on header byte; SGTIN-96 (0x30) and SGTIN-198 (0x36) are primary
