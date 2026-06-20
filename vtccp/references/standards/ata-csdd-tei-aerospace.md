# ATA CSDD — Text Element Identifiers (TEIs) — Aerospace Parts Traceability

**Standard**: Air Transport Association (ATA) Common Support Data Dictionary (CSDD)
**Domain**: Aerospace IUID (Item Unique Identification) and global part life-cycle traceability
**Transport syntax**: ISO/IEC 15434 — typically Format 12 (IUID Data Matrix) or Format 06
**Source**: ATA CSDD specification (supplied document, 2026-06-20)

---

## 1. What TEIs Are

Text Element Identifiers (TEIs) are **three-letter alphabetic data qualifiers** used in
aerospace parts marking to label data fields in machine-readable strings (primarily
2D Data Matrix barcodes). They serve the same purpose as:

| Standard | Qualifier type | Example |
|---|---|---|
| ATA CSDD TEIs (this doc) | 3-letter alpha | `MFR`, `PNO`, `SER` |
| ANSI MH10.8.2 Data Identifiers | Numeric short-form | `1P`, `4L`, `S` |
| GS1 Application Identifiers | 2–4 digit numeric | `01`, `10`, `21` |

TEIs are the aerospace-native qualifier system. They are NOT interchangeable with MH10.8.2
DIs or GS1 AIs, though they encode similar data concepts. A single barcode will use one
qualifier system consistently (TEI or DI, not mixed).

---

## 2. TEI Code Table

| TEI Code | Data Element | Description | Format / Example |
|---|---|---|---|
| `MFR` | Manufacturer / Enterprise ID | The identifier of the manufacturer — CAGE code, DUNS number, or ISO country/enterprise code | `MFR 12345` |
| `PNO` | Part Number | Current product or part identifier assigned by the design/responsible organization | `PNO 98765-101` |
| `SER` | Serial Number | Unique serial number for the specific unit (OEM-assigned) | `SER 0023` |
| `SEQ` | Sequence Number | Unique sequential tracking number assigned to the part | `SEQ 0001` |
| `UCN` | Unique Component ID Number | Used when an organization OTHER than the original manufacturer accomplishes serialization | `UCN ABC1234` |
| `LOT` | Lot Number | The lot or batch number to which the specific item belongs | `LOT 50921-A` |
| `ECI` | Export Control Flag/Indicator | Designates whether the item is subject to export control and/or restrictions (e.g., Wassenaar Arrangement) | `ECI [Value]` |
| `NST` | National Stock Number (NSN) | The official 13-digit federal item identification number | `NST [Value]` |
| `OEI` | Original Enterprise Identifier | Enterprise identifier of the original manufacturer (when different from current responsible org) | `OEI [Value]` |
| `OPN` | Original Part Number | Part number assigned by the original equipment manufacturer (OEM) — used when current responsible org differs | `OPN [Value]` |
| `BPN` | Base Part Number | Base part number without dash numbers or revision detail | `BPN [Value]` |
| `REV` | Revision Number | Current configuration revision level of the part | `REV B` |

---

## 3. Implementation — ISO/IEC 15434 Encoding

TEIs are encoded using the **ISO/IEC 15434** transport syntax. In aerospace IUID marking
the most common formats are:

- **Format 06**: `[)>\u001E06\u001D{TEI} {value}\u001D...` — used in commercial supply chain
- **Format 12**: IUID-specific Data Matrix format, incorporates TEIs directly

Within the encoded string:
- **`\u001D` (GS — Group Separator, ASCII 29)** separates data elements
- **`\u001E` (RS — Record Separator, ASCII 30)** separates records
- **`\u0004` (EOT — End of Transmission, ASCII 4)** terminates the message

A typical IUID-encoded Data Matrix payload:
```
[)>RS 06 GS MFR 12345 GS PNO 98765-101 GS SER 0023 RS EOT
```

In raw ASCII: `[)>\x1E06\x1DMFR 12345\x1DPNO 98765-101\x1DSER 0023\x1E\x04`

---

## 4. Relationship to DoD IUID (MIL-STD-130)

DoD Item Unique Identification (IUID) uses **both** MH10.8.2 Data Identifiers AND ATA TEIs
depending on the marking standard in effect:

| Context | Qualifier system | Key IDs |
|---|---|---|
| DoD MIL-STD-130N / IUID | MH10.8.2 DIs | `1P` (Part No), `S` (Serial), `4L` (Lot) |
| Aerospace commercial (ATA) | TEIs | `PNO`, `SER`, `LOT` |
| Both | ISO/IEC 15434 transport | Format 06 or Format 12 |

The DoD Guide to IUID (v3.0) and MIL-STD-130N (both in `references/mil-std/`) govern the
DoD path. The ATA CSDD governs the commercial aerospace path. VTCCP may encounter either.

---

## 5. Naming Collision Warning — `ECI`

⚠ **`ECI` appears in two completely unrelated contexts in VTCCP:**

| Context | Meaning |
|---|---|
| ATA TEI `ECI` | **Export Control Indicator** — a data field qualifier in a barcode payload |
| ISO/IEC 15424 / QR ECI | **Extended Channel Interpretation** — a barcode encoding mode flag |

These are entirely different concepts. An `ECI` TEI tag in a data string is payload
content; it has no relationship to the `ECI` protocol flag in the AIM symbology identifier
`]Q2` or in QR Code encoding mode selection.

---

## 6. VTCCP Parser Gap

The existing `ISO15434Parser.cs` handles **MH10.8.2 Data Identifiers** (numeric DIs:
`1P`, `4L`, `S`, `18V`, etc.). It does NOT currently handle ATA TEIs.

**Impact**: If VTCCP scans an aerospace part marked with TEI-encoded data (ISO 15434
Format 06 or 12 with `MFR`, `PNO`, `SER` etc.), the parser will not recognize the
qualifiers. The GS1 AI path is also separate and unaffected.

**Required when aerospace parts are in scope**: Add a TEI branch to `ISO15434Parser.cs`
alongside the existing MH10.8.2 DI branch. Detection heuristic: if the first element
after Format 06 header starts with three uppercase alpha characters followed by a space
(e.g., `MFR `, `PNO `, `SER `), treat as TEI-encoded.

---

## 7. MFR CAGE Code Context

The `MFR` TEI value is most commonly a **CAGE code** (Commercial and Government Entity
code) — a 5-character alphanumeric issued by the Defense Logistics Agency. Format:
- Positions 1 and 5: must be alpha (not I, O, Q, Z)
- Positions 2–4: alphanumeric

Example: `MFR JD9A7` where `JD9A7` is the manufacturer's CAGE code.

DUNS (Dun & Bradstreet) and ISO country+enterprise codes are also valid MFR values.

---

*Document version 1.0 — 2026-06-20. Source: ATA CSDD (supplied 2026-06-20).
Cross-references: MIL-STD-130N, DoD Guide IUID v3.0, ANSI MH10.8.2 (all in references/mil-std/).*
