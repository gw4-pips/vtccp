# RFID Cross-Validation — Standalone Product Architecture & Strategy

Rev 1.0 — 2026-07-12

## Strategic premise

The RFID cross-validation feature is verifier-agnostic by nature: a USB RFID reader
connects independently of whatever barcode verifier (or scanner) is in use. This means
the feature can be productized as a standalone offering targeting the global installed
base of competitive verifier accounts — Axicon, OMRON LVS, Webscan TruCheck, REA, and
others — without requiring Command Pilot or any DataMan integration.

The PoC work (RFID ME USB reader + raw hex parse + SGTIN decode + barcode comparison +
report output) is structurally the minimum viable standalone product, not just a feature
demo. Build it right and it ships.

## Market gap this fills

No commercially available product today:
- Captures RFID EPC alongside a barcode verification or scan event from any verifier
- Decodes EPC per current GS1 TDS (SGTIN-96/198 and full scheme set)
- Validates GCP encoding integrity against the GS1 GCP Length Table
- Cross-validates GTIN and Serial between RFID and barcode sources
- Produces a formatted verification-style report integrating both data channels

Handheld combo readers (Zebra RFD-series + TC mobile computer, Honeywell equivalents)
exist and can do basic barcode + RFID capture simultaneously with internal scripting
(Zebra DataWedge, Honeywell EasyDL). But they provide only basic string-match comparison
with no GCP validation, no segment-level EPC decode display, no extended GS1 reporting,
and — critically — no ISO grade data (they are scanners, not verifiers; they cannot
provide ISO 15416/15415 compliant quality grades for PIPS/VCCS compliance purposes).
The standalone CP product operates at a higher level than these existing solutions.

## Hardware configurations (three tiers)

### Tier 1 — No verifier (scan-only validation)
- USB barcode scanner (keyboard wedge, ~$50–100) + USB RFID reader (~$200)
- No ISO grade output; data match validation only (GTIN + serial from barcode vs. RFID)
- Acknowledged as aesthetically "kludgy" for a verification workflow but functional
- Lowest barrier to entry; useful for track-and-trace or pre-verification data QC
- Viable for initial deployments where verification is done separately

### Tier 2 — Competitive verifier + RFID (target configuration)
- Customer's existing verifier (any brand) + USB RFID reader
- CP hooks into verifier software output via minimal-config adapter (see below)
- Full output: ISO grade (from verifier) + RFID cross-validation + GCP encoding check
- This is the primary commercial target

### Tier 3 — Command Pilot full (DataMan + RFID)
- DataMan verifier via DMCC TCP + USB RFID reader
- Full Command Pilot feature set; RFID enrichment layer sits above DataMan adapter

## Competitive verifier interface — minimal-config adapters

Goal: zero or near-zero user configuration per verifier brand. Detection and hookup
should be automatic wherever possible.

| Verifier | Integration path | Config required |
|---|---|---|
| **OMRON LVS** | Query embedded SQLite/SQL database directly | Database file path (auto-detect from default install location) |
| **Axicon** | Parse CSV export via ScanDB tool, or HTML report | Watch folder path (configurable or default) |
| **Webscan TruCheck** | DMST HTML scrape (already implemented in CP) | IP address or USB |
| **REA** | Research needed | TBD |
| **DataMan** | DMCC TCP (already implemented in CP) | IP address |

OMRON's SQLite database is the cleanest path: CP queries it directly with a known
schema, no manual export step, no user action required per scan. Auto-detect the LVS
database at its default install path; user confirms on first run only.

Axicon's ScanDB CSV export requires a folder-watch approach: CP monitors a configured
output folder for new CSV files, parses them as they appear. RepGen HTML could also
be parsed as a fallback.

## RFID input architecture (two-layer)

### Layer 1 — Acquisition (thin, reader-specific, swappable)
Interface: `IEpcReader` — delivers a hex string per read event.
- `HidEpcReader` — global raw-input hook; captures CR-terminated hex string from
  HID keyboard-wedge reader; no driver; works with any plug-and-play UHF reader
- `ComPortEpcReader` — opens configurable COMx SerialPort; reads lines; supports
  trigger command if reader exposes one

Both deliver the same upward contract: a raw EPC hex string.
Raw hex as the interface contract makes the feature brand- and model-agnostic;
virtually any UHF RFID reader can output hex, regardless of brand or generation.

### Layer 2 — Parse / Validate / Enrich (reader-agnostic, pure C#)
- Strip optional CRC+PC prefix (auto-detect by length: 28 hex chars = has CRC+PC,
  24 hex chars = raw EPC; configurable override)
- Dispatch on header byte to scheme-specific decoder
- For SGTIN-96/198: extract GTIN-14 + Serial; validate partition value against GCP
  Length Table; flag partition encoding errors with correct value per GS1
- Compare decoded RFID values against barcode source (AI 01 / AI 21, or DL URI path)
- Produce RfidValidationResult for report injection and Excel capture

## EPC scheme dispatch table (TDS 2.3)

| Header | Scheme | Bits | GS1 key | Serial |
|---|---|---|---|---|
| 0x30 | SGTIN-96 | 96 | GTIN | Numeric ≤10 digits |
| 0x36 | SGTIN-198 | 198 | GTIN | Alphanumeric ≤20 chars |
| 0x31 | SSCC-96 | 96 | SSCC | — |
| 0x32 | SGLN-96 | 96 | GLN | Numeric extension |
| 0x39 | SGLN-195 | 195 | GLN | Alphanumeric extension |
| 0x33 | GRAI-96 | 96 | GRAI | Numeric |
| 0x37 | GRAI-170 | 170 | GRAI | Alphanumeric |
| 0x34 | GIAI-96 | 96 | GIAI | Numeric |
| 0x38 | GIAI-202 | 202 | GIAI | Alphanumeric |
| 0x2C | GDTI-96 | 96 | GDTI | Numeric |
| 0x3A | GDTI-174 | 174 | GDTI | Alphanumeric |
| 0x35 | GID-96 | 96 | General ID | Non-GS1 |
| 0x2D | GSRN-96 | 96 | GSRN | — |
| 0x2E | GSRNP-96 | 96 | GSRN (Provider) | — |
| 0x3C | CPI-96 | 96 | Component/Part | Numeric |
| 0x3D | CPI-var | Variable | Component/Part | Alphanumeric |
| 0x2F | USDOD-96 | 96 | US DoD | — |

64-bit schemes (SGTIN-64 etc.) are sunsetted — treat as unrecognized/log and skip.

## GCP Length Table

- **Status as of Jan 2026**: password-protected at gs1.org/standards/bc-epc-interop
- **For decode**: NOT required — partition value in SGTIN-96 tag gives GCP length
  directly via the fixed 7-row TDS partition table (embedded in code, never changes)
- **For GCP encoding validation**: required — cross-check partition-derived GCP length
  against expected GCP length for that company prefix; report discrepancy as tag
  encoding defect with correct value per GS1 (not a blocker for data comparison)
- **Supply path**: PIPS/VCCS can supply as a reference file; CP bundles with version
  check; OMRON/Axicon adapter lookup as fallback for unknown prefixes

## Excel output

- **New "RFID" child worksheet tab** in existing workbook (not additional columns on
  the 167-column TruCheckCompatibleSchema main sheet)
- Correlation key: **Scan ID** (sequential integer or short GUID generated by CP at
  trigger time, stamped on both main sheet row and RFID tab row)
- Timestamp on both for human readability; Scan ID is the authoritative join key
- RFID tab columns: Scan ID | Timestamp | Raw EPC Hex | Header | Filter | Partition |
  GCP | Item Ref | Indicator Digit | GTIN-14 | Serial | Barcode GTIN | Barcode Serial |
  GTIN Match | Serial Match | Partition Encoding Check | Notes

## Report output

- **TruCheck-style HTML report header**: add SGTIN summary rows (GTIN, serial, match
  result, partition check) to existing Characteristics table
- **Optional RFID results table**: generated as a C# HTML block injected after XSLT
  transform completes; only present when RFID reader is configured — existing XSLT
  templates unchanged
- **CP native reports**: same pattern — RFID section conditional on RFID being enabled

## GS1 Digital Link and Sunrise 2027 context

- GS1 Digital Link URI: `https://{host}/01/{GTIN-14}/21/{serial}?17={expiry}&10={lot}`
  Primary key (GTIN) mandatory in path; key qualifiers in path; attributes in query
  string; case-sensitive after hostname; GTIN must be 14 digits (zero-padded)
- QR Code + GS1 DL is the Sunrise 2027 retail POS 2D standard (not DataMatrix —
  DataMatrix is technically capable but not consumer-smartphone-scannable)
- By end of 2027 products in scope will increasingly carry: linear UPC/EAN + 2D QR
  (GS1 DL) + RFID tag simultaneously — CP's multi-surface validation story is aligned
  with this trajectory

## Relationship to Command Pilot and future roadmap

- Standalone RFID Validation: lightweight tray app, USB reader(s), minimal or no
  verifier integration initially; sells independently to competitive accounts
- Command Pilot full: DataMan + full verifier integration suite + RFID; sells to
  DataMan accounts and future multi-verifier accounts
- Same .NET 8 / WiX installer stack; same parser/validator/report core; feature flags
  determine which adapters are enabled at runtime
- Standalone is the natural upsell path into Command Pilot
- Future adapters: OMRON LVS (SQLite DB), Axicon (ScanDB CSV), REA (TBD) — each added
  per the verifier-agnostic adapter pattern (IVerificationSource → VerificationRecord)
