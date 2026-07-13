---
name: Axicon 15500 and LVS 9510 Verifier Integration Research
description: Manufacturer identity, software interface, and planned VTCCP integration approach for the two additional barcode verifiers in the lab.
---

## Axicon 15500 (Axicon Auto ID Limited, Oxfordshire UK)
- Part of "15000 series linear & 2D verifier" product line
- Software: Axicon 15000 series v9.1 (March 2026); implements ISO/IEC 15415:2024 + 15416:2025
- USB connected
- Two automation hooks visible from website:
  1. **Automatic File Naming plugin** — writes result file to a configurable path after each scan
  2. **Camera Verifier Software Plugins** page — details TBD (could be COM automation or DLL hook)
- **Planned VTCCP approach**: AxiconFileAdapter.cs — folder watcher + CSV/XML parser
- Need from user: what format does the export file use (CSV / XML / text)?

## LVS 9510 (Omron Microscan — NOT Webscan)
- Manufactured by Microscan, acquired by Omron ~2018; now branded Omron Microscan LVS-9510
- Model family: LVS-95XX series
- 5.0 megapixel camera-based desktop verifier; USB plug-and-play
- Grades 1D and 2D barcodes to ISO/IEC standards; no equipment change per code type
- Software: "LVS-95XX software" (Windows proprietary)
- Manual available at: assets.omron.eu/downloads/latest/manual/en/lvs-9510_operation_manual_en.pdf
- Omron US page: automation.omron.com/en/us/products/family/VF9510
- **Planned VTCCP approach**: MicroscanLvsFileAdapter.cs — same folder-watch pattern as Axicon
- Both adapters share the same IVerifierResult interface that TruCheck already uses
- Need from user: does LVS-95XX software have an auto-save/export path setting? What format?

**Why durable:** User has both devices physically in lab and plans to integrate them into VTCCP. Integration gated on user confirming export format from each software.
