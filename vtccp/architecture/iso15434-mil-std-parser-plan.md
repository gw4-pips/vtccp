# ISO 15434 / MIL-STD Parser — Status and Plan

> Created: 2026-06-09
> Related code: `vtccp/ExcelEngine/Utilities/ISO15434Parser.cs`,
>               `vtccp/ExcelEngine/Utilities/AutoBatchExtractor.cs`
> Reference docs: `vtccp/references/mil-std/` (intake folder — pending uploads)

---

## 1. What is already built

`ISO15434Parser.cs` is a **partial ISO 15434 parser** — it handles the envelope
syntax correctly but extracts only one field (batch/lot).

### Envelope mechanics — COMPLETE

| Feature | Status |
|---|---|
| `[)>` envelope detection | ✓ |
| RS (0x1E) format-indicator separator | ✓ |
| GS (0x1D) field delimiters | ✓ |
| EOT (0x04) end-of-message | ✓ |
| DataMan text-escaped forms (`<RS>`, `<GS>`, `<EOT>`) | ✓ |
| All format indicators accepted (06, 05, 12, etc.) | ✓ |
| DI prefix stripping | ✓ |

### Field extraction — PARTIAL

| DI | Meaning | Status |
|---|---|---|
| `4L` | Lot / Batch (primary) | ✓ extracted |
| `10L` | Lot alternate | ✓ extracted (same call as 4L, priority order) |
| `1L` | Lot tracking | ✓ extracted |
| `L` | Legacy lot | ✓ extracted |
| All other DIs | Part number, serial, cage code, contract, date, qty, etc. | **NOT extracted** |

`AutoBatchExtractor` calls `GS1Parser.ExtractBatchLot()` first, then falls
back to `ISO15434Parser.ExtractBatchLot()`. This is the only production
caller of the ISO 15434 parser — it is used exclusively for `BatchMode.AutoFromGS1`.

---

## 2. What is missing — gap analysis

The parser has no concept of the broader DI set. It cannot:

- Enumerate all DI fields present in a symbol
- Parse part numbers, serial numbers, cage codes, contract numbers, dates, quantities
- Validate DI value formats or lengths
- Produce a structured `ParsedDI` record analogous to `ParsedAI` in `GS1Parser`
- Route to a `ParseDetailRowWriter`-style display for MIL-STD records

This gap is **not a defect** — the parser was written for the one field VTCCP
needed at the time (batch extraction for `AutoBatchExtractor`). The envelope
mechanics are solid and can be extended.

---

## 3. Documents needed before building the full parser

### Required (gating)

| Standard | Purpose | Acquisition |
|---|---|---|
| **ANSI MH10.8.2** (current ed.) | The complete DI dictionary — every valid DI, its meaning, format, and allowed length | Purchase from MHIA / ANSI (not freely available) |
| **MIL-STD-130** (current rev N or later) | DoD UID item identification marking; defines mandatory DI set for Data Matrix UID labels | Free — ASSIST-QuickSearch (quicksearch.dla.mil) |
| **MIL-STD-129** (current rev P or later) | Military shipment/storage marking; defines PDF417 label DI set | Free — ASSIST-QuickSearch |

### Helpful (non-gating)

| Standard | Purpose |
|---|---|
| **ISO/IEC 15434:2006** (or latest ed.) | The normative transfer syntax spec; our envelope parser was written from its spec |
| **ISO/IEC 15418:2016** (ANSI MH10.8.2 / EAN/UCC companion) | Cross-reference for AIs vs DIs in shared barcodes |

**Intake folder**: `vtccp/references/mil-std/`
Upload MIL-STD-129, MIL-STD-130, and any ANSI MH10.8.2 materials here.

---

## 4. What the full parser will look like

Once documents are in hand, the implementation will follow the same pattern as
`GS1Parser` — a `ParsedDI` record, a `DI` dictionary keyed on DI string, and a
`ParseAllFields()` method that returns an ordered list of `ParsedDI` entries.

### `ParsedDI` (proposed)

```csharp
public sealed record ParsedDI(
    string   Di,          // e.g. "4L", "P", "1P", "25S", "S", "V"
    string   Label,       // human label, e.g. "Lot / Batch", "Part Number"
    string   Value,       // raw string value
    string   Standard,    // "MIL-STD-130", "MIL-STD-129", "ANSI MH10.8.2", "Unknown"
    bool     IsValid      // format/length check passed
);
```

### Known DI set for military applications (prior knowledge — must be cross-checked against docs)

This table is a pre-document best-effort; every row must be verified against
ANSI MH10.8.2 and the relevant MIL-STD before writing production code.

| DI | Meaning | MIL-STD-130 | MIL-STD-129 |
|---|---|---|---|
| `P` | Part Number | ✓ mandatory | ✓ |
| `1P` | Supplier Part Number | optional | optional |
| `S` | Serial Number | ✓ (UID path 1) | |
| `1T` | Traceability Code / Serial | ✓ (UID path 2) | |
| `4L` | Lot / Batch | ✓ (UID path 3) | ✓ |
| `Q` | Quantity | | ✓ |
| `V` | Supplier / Vendor Code (CAGE) | see 25S | ✓ |
| `25S` | Cage Code (DoD) | ✓ mandatory | ✓ |
| `6J` | Contract Number | | ✓ |
| `J` | Purchase Order Number | | ✓ |
| `K` | Purchase Order Line | | ✓ |
| `9D` | Date (YYYYMMDD) | | ✓ |
| `12D` | Manufacture Date | optional | optional |
| `15D` | Expiry Date | optional | optional |
| `14K` | National Stock Number (NSN) | | ✓ |

> ★ Do NOT use this table as the implementation source.
> It is a planning sketch only. Use ANSI MH10.8.2 as the authoritative DI dictionary.

### Format-indicator handling

| Format indicator | Standard | Action |
|---|---|---|
| `06` | ANSI MH10.8.2 (primary DI format) | full parse |
| `05` | EDI (ANSI X12 / EDIFACT) | extract envelope, emit raw body; no DI parse |
| `12` | ASC X12 | same as 05 |
| other | unknown | extract envelope, emit raw body |

---

## 5. Display design

When `ApplicationStandard == "ISO 15434"` (or similar string) in `VerificationRecord`,
`ParseDetailRowWriter` should invoke a `DI` renderer analogous to the existing `AI`
renderer — same row structure, DI label in column A, value in column B, format flag
in column C. The batch/lot row (DI `4L`) should render identically to GS1 AI `(10)`.

---

## 6. Implementation sequence (blocked pending documents)

1. Upload MIL-STD-130, MIL-STD-129, ANSI MH10.8.2 → `vtccp/references/mil-std/`
2. Read documents → derive authoritative DI table and format/length rules
3. Extend `ISO15434Parser` with `ParseAllFields()` returning `IReadOnlyList<ParsedDI>`
4. Add `DI` dictionary (label + format rule) populated from MH10.8.2
5. Wire `AutoBatchExtractor` to use the new `ParseAllFields()` and continue extracting
   `4L` as before (no regression)
6. Wire `DmstResultParser` to call `ParseAllFields()` when
   `ApplicationStandard` indicates ISO 15434
7. Store `IReadOnlyList<ParsedDI>` in `VerificationRecord.ParsedDIs`
8. Extend `ParseDetailRowWriter` to render DI rows
9. Regression test: existing batch-lot fixture must still pass unchanged
