---
name: Webscan composite imports
description: Two-symbol Webscan HTML exports are one linear-plus-2D verification event with one shared RFID scan.
---

Treat a Webscan export containing exactly one linear report and one 2D report as one composite record, paired by symbol family rather than report order. Preserve each native report's grades, quality rows, Data Format Check, and image. Use one RFID inventory result to compare against both normalized GTINs; barcode agreement and RFID agreement are separate checks, and overall composite verification fails if either check fails. If the 2D report lacks a valid AI (01) GTIN, keep the native HTML/data evidence but report barcode agreement as FAIL / INCOMPLETE and hard-fail the composite.

**Why:** Webscan embeds independent report tables in one file, so flattening rows causes duplicate quality keys and mixing evidence. One physical item should not receive two RFID inventory windows.

**How to apply:** Keep the 2D symbol as the primary record for the existing RFID path, populate the existing linear multi-mode fields from the linear report, and keep the individual comparison outcomes visible.