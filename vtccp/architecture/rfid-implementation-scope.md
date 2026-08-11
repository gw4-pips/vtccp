# RFID Cross-Validation — Phased Implementation Scope

Rev 1.2 — 2026-08-11

---

## Prerequisites — resolved

**RFID reader confirmed:** AsReader ASR-P35U UHF RFID Desktop Reader (unit KE00048)
- EPC Class 1 Gen2, 902–928 MHz US-band (REGION_US), TX power 13–27 dBm
- USB → VCP (Virtual COM Port); VID=0x339C / PID=0x271B; 115200 8N1
  No FTDI driver required — AsReader proprietary USB CDC driver installs via Windows Update
  or from the SDK zip. **Never open the COM port directly** — the SDK manages it internally.
- Communicates via AsReader SDK (`AsReaderP3xU.dll`, v1.3.0, 2026-02-13); namespace
  `AsReaderP3xU`; main class `AsReader`
- SDK DLL placement: `vtccp/lib/asreader-p3xu-sdk-1.3.0/AsReaderP3xU.dll`
  (not committed to repo — obtain from AsReader SDK zip; see `PLACE-DLL-HERE.md` in that dir)
- Protocol notes: `vtccp/references/asr-p35u/docs/PROTOCOL-NOTES.md`
- Confirmed firmware: Main FW 1.8.0 (updated 2026-08-10); RFID module FW RED4S_v2.2.2_K_SD
- Known FW 1.8.0 defect: `CallBackCommandData` never fires for `ReadMemory`; results arrive
  via `CallBackReadTagData` instead. Workaround implemented in `AsReaderP35UEpcReader.cs`
  (`_pendingTidCb` one-shot hook). Vendor notified 2026-08-08.
- Triggerable: `StartInventory(maxTags:1)` for single-tag mode; `StopInventory()` for
  continuous mode. Hardware trigger button also exposed via `CallBackTriggerHandler`.

**GCP Length Table:** `vtccp/data/gcp-prefix-format-list.xml`
- Provided file dated 2026-05-03, 200,108 entries, 8.7 MB; bundled as dev/fallback copy
- Update service base URL: `https://my2dir-resolver-bwa7agd0ctehbqf3.eastus2-01.azurewebsites.net`

| Resource | URL path | Auth |
|---|---|---|
| Current month's data | `/tools/gcp/interop/current.xml` | `X-GCP-Interop-Key` header required |
| Previous month's data | `/tools/gcp/interop/previous.xml` | `X-GCP-Interop-Key` header required |
| Changelog / metadata | `/tools/gcp/interop/data.json` | No key required |

- Auth key stored in environment variable `GCP_INTEROP_KEY` (Replit Secret); never in source
- Auto-update logic at CP startup:
  1. HEAD `current.xml` with `X-GCP-Interop-Key` header → read `Last-Modified` response header
  2. Compare against date recorded in stored copy's XML root `date` attribute
  3. If newer: notify user in status bar — "GCP prefix table update available (YYYY-MM-DD).
     Update now?" — one-click GET downloads full `current.xml` to
     `%APPDATA%\VTCCP\gcp-prefix-format-list.xml`; stored copy date updated in settings
  4. Runtime loads: AppData copy if present → bundled `vtccp/data/` copy as fallback
  5. If `GCP_INTEROP_KEY` is absent: skip update check silently; use bundled copy; no error shown
  6. HEAD check is async and non-blocking — CP opens normally; notification appears in status bar

---

## Phase 0 — POC: RFID integrated into Command Pilot as it exists

**Priority: highest. Build this first.**

This integrates the RFID cross-validation feature directly into the existing VTCCP/
Command Pilot codebase, using the DataMan verifier as the barcode source. Every
component built here is reused in all subsequent phases.

### 0.1 — IEpcReader acquisition layer  ✅ BUILT

Interface and Phase 0 concrete implementation are complete:

```
vtccp/DeviceInterface/Rfid/
  IEpcReader.cs              — interface: ConnectAsync(), DisconnectAsync(),
                               TriggerInventoryAsync(timeout) → IReadOnlyList<EpcReadResult>,
                               CancelAsync(), IsConnected; IAsyncDisposable
  AsReaderP35UEpcReader.cs   — ✅ ASR-P35U implementation (unit KE00048, FW 1.8.0, SDK 1.3.0):
                                 • ConnectWithVCP(comPort) via AsReaderP3xU.dll SDK
                                 • SetDelegate (all 6 callbacks in one call, before connect)
                                 • SetRegion(REGION_US) + SetTxPower(20 dBm default)
                                 • StartInventory(maxTags:1) for single-tag triggered reads
                                 • EPC extracted from CallBackReadTagData (uppercase hex)
                                 • PC word + RSSI captured; two's complement RSSI correction applied
                                 • FW 1.8.0 defect: ReadMemory result via cbTag not cbCommand;
                                   ReadTidAsync(_pendingTidCb one-shot hook) implements workaround
                                 • Thread-safe: _lock (public methods), _stateLock (SDK callbacks)
  EpcReaderFactory.cs        — instantiates per config (ASR-P35U is the Phase 0 impl;
                               interface open for future HID/COM-line variants)
```

**DLL prerequisite:** `AsReaderP3xU.dll` (SDK 1.3.0) must be placed at
`vtccp/lib/asreader-p3xu-sdk-1.3.0/AsReaderP3xU.dll` before building. The DLL is not
committed to the repo — obtain from the AsReader SDK zip. A `PLACE-DLL-HERE.md` placeholder
should be present in that directory.

SDK protocol reference: `vtccp/references/asr-p35u/docs/PROTOCOL-NOTES.md`

EPC extraction: `result.tagdata.epc` delivers uppercase hex with no spaces. PC word arrives
as hex string (e.g. `"3000"`). Strip CRC+PC by length (28 hex chars = has CRC+PC; 24 = net
EPC) or leave as configured. No reader brand/model reference in any downstream class.

### 0.2 — EPC parser and scheme dispatch

New pure-C# parsing library, no external dependencies:

```
vtccp/DeviceInterface/Rfid/
  EpcParser.cs           — entry point: ParseHex(string hex) → EpcParseResult
  EpcSchemeDispatcher.cs — dispatches on header byte to scheme decoder
  Schemes/
    Sgtin96Decoder.cs    — 96-bit: header+filter+partition+GCP+ItemRef+Serial(numeric)
    Sgtin198Decoder.cs   — 198-bit: same structure, alphanumeric serial (6-bit chars)
    Sscc96Decoder.cs     — 96-bit SSCC
    UnknownSchemeDecoder.cs — logs header byte, returns raw hex for report
  EpcParseResult.cs      — scheme, all decoded field values, raw hex, any decode errors
  PartitionTable.cs      — static 7-row TDS partition table (GCP bits, IR bits per partition)
```

Header byte dispatch table (all TDS 2.3 active schemes coded; decoders built
incrementally — SGTIN-96/198 first, others stubbed and logged):

| Header | Scheme | Phase |
|---|---|---|
| 0x30 | SGTIN-96 | Phase 0 |
| 0x36 | SGTIN-198 | Phase 0 |
| 0x31 | SSCC-96 | Phase 3 |
| 0x32/0x39 | SGLN-96/195 | Phase 3 |
| 0x33/0x37 | GRAI-96/170 | Phase 3 |
| 0x34/0x38 | GIAI-96/202 | Phase 3 |
| 0x2C/0x3A | GDTI-96/174 | Phase 3 |
| 0x35 | GID-96 | Phase 3 |
| 0x2D/0x2E | GSRN/GSRNP-96 | Phase 3 |
| 0x3C/0x3D | CPI-96/var | Phase 3 |
| 0x2F | USDOD-96 | Phase 3 |

### 0.3 — GCP partition validation

```
vtccp/DeviceInterface/Rfid/
  GcpLengthTable.cs      — loads GCP length table (XML/JSON) from bundled resource file
  GcpValidator.cs        — for a given GTIN prefix, looks up expected GCP length;
                           compares against partition-derived GCP length;
                           returns GcpValidationResult (Match | Mismatch + correct value)
```

Report behavior on mismatch: "Partition encoding error — tag encodes GCP as N digits;
GS1 registry indicates correct GCP length is M digits for this prefix." Does not block
data comparison. GCP table bundled as resource; version-tagged; PIPS/VCCS supply path.

### 0.4 — Cross-validation against DataMan barcode data

```
vtccp/DeviceInterface/Rfid/
  RfidValidator.cs       — takes EpcParseResult + VerificationRecord (barcode source);
                           extracts GTIN-14 + serial from both sources;
                           produces RfidValidationResult (field-level match/mismatch)
```

Barcode data sources (in priority order):
1. GS1 DataMatrix: AI (01) = GTIN-14, AI (21) = Serial
2. QR Code / GS1 Digital Link URI: `/01/{GTIN-14}` path, `/21/{serial}` qualifier
3. Linear UPC-A/EAN-13: GTIN only (no serial); serial match reported as N/A

For UPC-A/EAN-13: GTIN comparison only; serial field explicitly noted as "not available
in linear symbol" — not a mismatch, not an error.

### 0.5 — Trigger coordination with DataMan scan

```
vtccp/DeviceInterface/Rfid/
  RfidScanCoordinator.cs — generates Scan ID (sequential int) at trigger time;
                           stamps ScanId on VerificationRecord before DataMan trigger;
                           sends RFID trigger/read command (COM port) or arms HID window;
                           correlates RFID read arriving within configurable time window;
                           produces RfidEnrichedRecord (VerificationRecord + RfidValidationResult)
```

Scan ID: sequential integer, generated by CP at trigger time, stamped on both the
main Excel row and the RFID tab row. This is the authoritative correlation key.
Timestamp on both records for human readability; row number not used as a key.

### 0.6 — Excel: RFID child tab

```
vtccp/ExcelEngine/
  Schema/RfidTabSchema.cs     — column definitions for RFID worksheet
  Writer/RfidTabWriter.cs     — writes RfidValidationResult rows to "RFID" worksheet
```

Called from `ExcelWriter` after main sheet row is written. New worksheet "RFID" in
same workbook. Columns:

Scan ID | Timestamp | Raw EPC Hex | EPC Scheme | Header (hex) | Filter | Partition
(encoded) | Partition (correct per GS1) | Partition Check | GCP (digits) | GCP Value |
Item Ref | Indicator Digit | GTIN-14 (decoded) | Serial (decoded) | Barcode Source |
Barcode GTIN | Barcode Serial | GTIN Match | Serial Match | Notes

### 0.7 — TruCheck-style HTML report

Two additions to `DmstHtmlReport.cs` output:

**Header section** (always present when RFID configured):
Append to Characteristics table — rows for: EPC (hex), EPC Scheme, GTIN (RFID decoded),
Serial (RFID decoded), Barcode GTIN, Barcode Serial, GTIN Match (PASS/FAIL), Serial
Match (PASS/FAIL or N/A), Partition Encoding (PASS / MISMATCH + correct value).

**RFID Detail table** (optional, configurable):
HTML block generated in C# and injected after XSLT transform output. Existing XSLT
templates are not modified. Table shows all decoded EPC segment values. Section only
present when RFID reader is configured and produced a result. "No RFID read" noted
explicitly if trigger fired but no tag returned within time window.

### 0.8 — CP settings / configuration UI

Minimal additions to existing CP settings:
- RFID enabled (bool)
- COM port name (ASR-P35U VCP; e.g. COM4; auto-detect VID=0x339C / PID=0x271B via WMI)
- TX power dBm (int, 13–27; default 20)
- Read time window (seconds, default 3.0)
- CRC+PC prefix mode: Auto / Strip / None
- GCP table: use bundled / use AppData / custom path
- GCP auto-update: enabled (bool); check on startup; notify when newer version available
- GCP update URL: semi-private URL (provided by operator; stored in settings; not hardcoded)
- RFID detail table in report (bool)

---

## Phase 1 — Standalone product: RFID validation without Command Pilot

**Priority: immediately after Phase 0 POC is validated.**

Thin standalone Windows tray application using all Phase 0 parser/validator/report
code as shared library components. No DataMan dependency. No DMCC. No DMST.

### Barcode input (Tier 1 — no verifier)

```
standalone/InputSources/
  HidBarcodeReader.cs    — same HID hook pattern as EpcReader; captures barcode
                           data from USB barcode scanner (keyboard wedge)
  ManualBarcodeEntry.cs  — simple UI dialog for manual entry (demo/testing)
```

Barcode data parsed from raw scanned string using existing GS1Parser.cs or
gs1-syntax-engine wrapper — extracts AI (01) GTIN and AI (21) serial if present.

### Application shell

```
standalone/
  StandaloneRfidApp.cs   — NotifyIcon tray app; no main window
  ScanSession.cs         — manages scan cycle: arm barcode reader → arm RFID reader →
                           wait for both within time window → validate → output
  StandaloneReport.cs    — generates standalone verification report (PDF or HTML)
                           showing RFID decode + barcode data + match results;
                           does not require TruCheck-format (simpler layout)
```

### Installer

WiX MSI — same stack as full Command Pilot installer. Self-contained .NET 8 exe.
Single optional component in the WiX project (shared with future Command Pilot
installer; standalone is a feature subset of the same WiX project).

### Standalone report format

Does not replicate full TruCheck format (no ISO grades, no verifier data). Focused:
- RFID cross-validation header (GTIN match, serial match, partition check)
- Full EPC segment decode table
- Barcode data summary (source, GTIN, serial)
- Timestamp, operator, session info
- Optionally: appended to verifier's own PDF (as an addendum page) — future option

---

## Phase 2 — Competitive verifier adapters (Tier 2: verifier + RFID)

**Priority: follows Phase 1 standalone validation.**

Formalizes the `IVerificationSource` adapter pattern and builds first two non-DataMan
adapters. Standalone app gains verifier-coupled capability; no separate USB barcode
scanner needed.

### IVerificationSource interface

```
vtccp/DeviceInterface/
  IVerificationSource.cs        — event: VerificationResultReady(VerificationRecord)
  Adapters/
    DataManVerificationSource.cs  — wraps existing DMCC/Push XML infrastructure (rename)
    OmronLvsVerificationSource.cs — Phase 2
    AxiconVerificationSource.cs   — Phase 2
    WebscanVerificationSource.cs  — wraps existing DMST scraper (rename)
```

### OMRON LVS adapter

Integration path: query embedded SQLite database (default install path auto-detected;
user confirms on first run). CP polls for new records since last scan ID; parses ISO
grade parameters and decoded barcode data; produces VerificationRecord.

No per-scan user action required once configured. No manual export step.

### Axicon adapter

Integration path: file system watcher on ScanDB CSV output folder (configurable;
auto-detect Axicon default). New CSV rows parsed as they appear; produces
VerificationRecord. RepGen HTML parse as fallback.

### REA adapter

Research needed before scoping. Held.

---

## Phase 3 — Full EPC scheme coverage + GS1 Digital Link input

**Priority: follows Phase 2.**

### Remaining EPC decoders

Complete all scheme decoders stubbed in Phase 0 (SSCC, SGLN, GRAI, GIAI, GDTI, GID,
GSRN, CPI, USDOD). All decoders follow same pattern as SGTIN-96; difference is only
the field layout and GS1 key type.

### GS1 Digital Link QR code input

CP parses GS1 Digital Link URIs from QR code barcode data:
- Detect `https://` or `http://` prefix in decoded barcode string
- Route to gs1-syntax-engine `GS1Encoder.DataStr` setter (already handles DL detection)
- Extract AI (01) GTIN-14 and AI (21) serial from DL path
- Compare against RFID decoded values as for any other barcode source

Enables Sunrise 2027 use case: QR Code (GS1 DL) + RFID tag validated simultaneously.

### Multi-surface validation (linear + 2D + RFID)

When both a linear symbol and a 2D symbol are present (Sunrise 2027 dual-symbol
packaging): validate all three surfaces against each other.
Report: linear GTIN → 2D GTIN → RFID GTIN (three-way match table).

---

## Phase 4 — Command Pilot as full verifier-agnostic platform

**Priority: follows Phase 3; longer-term.**

Integrate all Phase 2 adapters into full Command Pilot (not just standalone). Command
Pilot becomes the unified management layer for verification sessions regardless of
verifier brand: DataMan, Webscan TruCheck, OMRON LVS, Axicon, and any future adapter.

All CP reporting (TruCheck-format HTML, Excel, CP native) available for any adapter
source. RFID enrichment layer available on all sessions regardless of adapter.

Selling proposition: "Universal barcode verification data management — any ISO-compliant
verifier, enhanced GS1 reporting, RFID cross-validation, TruCheck-format output."

---

## Phase 5 — 2D/Multi-mode verification (held)

**Priority: held pending Cognex roadmap clarity on Sunrise 2027 / 2DiR response.**

If Cognex does not release QR code grading to ISO 15415 / ISO 29158 standards within
the Sunrise 2027 window, CP builds it. Equivalent capability to Webscan MultiMode.
Significant scope — not scheduled until Phase 4 is complete and Cognex posture is clear.

---

## Summary: phase priority and dependencies

```
Prerequisites resolved: ASR-P35U SDK VCP; AsReaderP35UEpcReader.cs built; GCP table 2026-05-03 in vtccp/data/
    │
    ▼
Phase 0: RFID integrated into CP/DataMan POC ←── HIGHEST PRIORITY (acquisition layer built ✅)
    │  (validates all core components: acquisition ✅, parse, validate, Excel, report)
    │
    ▼
Phase 1: Standalone product (thin tray app, shared Phase 0 library)
    │  (immediately deployable at competitive accounts; monetization path opens)
    │
    ▼
Phase 2: Competitive verifier adapters (OMRON SQLite, Axicon CSV)
    │  (Tier 2 hardware config — verifier replaces separate USB barcode scanner)
    │
    ▼
Phase 3: Full EPC scheme coverage + GS1 DL QR input (Sunrise 2027 readiness)
    │
    ▼
Phase 4: Command Pilot as universal verifier-agnostic platform
    │
    ▼
Phase 5: 2D multi-mode verification (held — pending Cognex 2DiR response)
```

## Shared components across all phases

| Component | Phase introduced | Used in |
|---|---|---|
| `IEpcReader` + implementations | 0 | 0, 1, 2, 3, 4 |
| `EpcParser` + scheme decoders | 0 | 0, 1, 2, 3, 4 |
| `GcpLengthTable` + `GcpValidator` | 0 | 0, 1, 2, 3, 4 |
| `RfidValidator` | 0 | 0, 1, 2, 3, 4 |
| `RfidTabWriter` (Excel) | 0 | 0, 2, 4 |
| `DmstHtmlReport` RFID injection | 0 | 0, 4 |
| `IVerificationSource` interface | 2 | 2, 4 |
| GS1 Digital Link QR parse | 3 | 3, 4 |
