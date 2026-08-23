---
title: GS1 multi-symbol report terminology
version: 1.0
revision: 2026-08-23
---

# GS1 and multi-symbol report terminology

## Scope

VTCCP uses **Multi-Symbol Report** for an export that contains two or more
independent native symbol reports. The current one-linear-plus-one-2D workflow
is a **Dual-Symbology Report**, rendered with the 2D report first because its
structured identity is the primary RFID comparison evidence.

This terminology describes report structure. It does not create a new
verification claim: each native report retains its own decoded value, image,
quality table, Data Format Check, and native result.

## Composite Component is a GS1 symbology term

**Composite Component** refers to the GS1 linear-plus-adjacent-2D component
concept, where the 2D component is carried with the linear symbol as part of a
GS1 Composite symbol. It is not a name for two separate Webscan HTML reports.
An independent GS1 DataMatrix report and an independent QR Code report remain
their own 2D symbologies. A future export containing several independent
symbols is a Multi-Symbol Report, not a Composite Component.

References:

* GS1 Glossary, “Composite Component”: https://gs1.org/standards/id-keys/gcp
* The local GS1 Barcode Syntax Engine wrapper and its `Symbology` enumeration:
  `lib/gs1-syntax-engine/src/GS1Encoder.cs`

The report importer may qualify a Multi-Symbol Report only when recognized
symbol identities agree and the RFID EPC supplies the configured identity.
Missing, ambiguous, unsupported, malformed, or contradictory evidence remains
unverified or rejected and is shown with its reason.