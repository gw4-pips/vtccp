# VCCS DMV TruCheck Command Pilot (VTCCP)

A native Windows desktop utility (WPF + C#/.NET 8) for Cognex DataMan DMV barcode verifier
verification-results logging (XLS/XLSX).

---

## Architecture Overview

```
vtccp/
├── ExcelEngine/          — Class library (Phase 1)
│   ├── Adapters/         — IExcelAdapter, XlsxAdapter (EPPlus), XlsAdapter (NPOI)
│   ├── Models/           — VerificationRecord, SessionState, ScanResult1D, ...
│   ├── Schema/           — ColumnSchema, TruCheckCompatibleSchema (167 columns), SchemaVersionWriter
│   ├── Session/          — SessionManager, SessionSidecar
│   └── Writer/           — ExcelWriter, ExcelFileManager, DataMatrix2DMapper,
│                           ISO15416Mapper, PerScanTableWriter, ElementWidthsWriter
├── DeviceInterface/      — Class library (Phase 2)
│   ├── Dmcc/             — DmccClient (async TCP), DmccCommand, DmccResponse, DmccStatus
│   ├── Dmst/             — DmstResultParser (XML → VerificationRecord), VerificationXmlMap
│   ├── Testing/          — MockDmccServer (loopback TCP for unit testing)
│   ├── DeviceConfig.cs   — Connection parameters (host, port, timeouts)
│   ├── DeviceInfo.cs     — Type, firmware, serial, name, calibration date
│   └── DeviceSession.cs  — High-level orchestration (ConnectAsync, TriggerAndGetResultAsync)
├── TestHarness/          — Console test driver (Phase 1 Tasks 1–4 + Phase 2)
└── README.md
```

---

## Phase 1 — Excel Engine (Offline)

### Task 1 — .NET scaffold & data models
- .NET 8 solution with ExcelEngine class library and TestHarness console project.
- All domain models: `VerificationRecord`, `SessionState`, `GradingResult`, `ScanResult1D`,
  `DataFormatCheckResult`, `ElementWidthData`, enumerations.
- `TruCheckCompatibleSchema` — 163-column schema producing DMV TruCheck-compatible output
  (Blocks A–I).

### Task 2 — 2D Data Matrix Excel engine
- `IExcelAdapter` abstraction; `XlsxAdapter` (EPPlus) and `XlsAdapter` (NPOI).
- `ExcelWriter` — title row, header row, data rows, GS1 DFC columns, append mode,
  row-limit check, `.xls` near-limit warning.
- `ExcelFileManager` — filename generation (`{Job}_{YYYY-MM-DD}.xlsx`),
  output-path resolution, file-lock detection, public `SanitizeFileName()`.
- `DataMatrix2DMapper` — maps `VerificationRecord` (2D) to 163-column value dictionary.

### Task 3 — 1D ISO 15416 Excel engine
- `ISO15416Mapper` — maps 1D records to the shared 163-column schema.
- `PerScanTableWriter` — writes LQZ/RQZ/HQZ individually per scan in sub-rows immediately
  below the summary row; capped at 10 scans per record; derives MinQZ.
- `ElementWidthsWriter` — writes element size / deviation tables to the "Element Widths"
  sheet with strict column-count enforcement.
- Mixed-symbology support (2D + 1D records in the same file).

### Task 4 — Session management & test harness
- `SessionManager` — lifecycle controller wrapping `ExcelWriter`.
  - `StartSession(state)` — opens or creates the output file; resumes from JSON sidecar if
    an interrupted session exists.
  - `AddRecord(record)` — appends one `VerificationRecord`; updates sidecar after each write
    for crash safety.
  - `CloseSession()` — saves the file, closes adapters, deletes the sidecar.
  - `SetNewOperatorAndRoll(operatorId)` — changes roll identifier mid-session.
- `SessionSidecar` — compact JSON snapshot (`{outputFile}.vtccp.json`) for resume support.
- `SchemaVersionWriter` — writes three metadata cells (marker / schema name / version) into
  row 1 past the last data column so downstream tooling can identify VTCCP-generated files.
- Full 6-record integration test: 3× GS1 DataMatrix, 1× DataMatrix, 1× UPC-A, 1× EAN-13.

---

## Phase 2 — Device Integration (DMCC / DMST)

### DMCC TCP Client (`DeviceInterface.Dmcc`)
- `DmccClient` — async TCP client that maintains a persistent connection to the DataMan device.
  - `ConnectAsync()` — connects, reads the welcome banner.
  - `SendAsync(command)` — writes a DMCC command line, reads the full response via idle-gap
    detection (`Socket.Receive` loop with `ReceiveTimeout = IdleGapMs`).
  - `DisconnectAsync()` / `DisposeAsync()` — clean teardown.
- `DmccCommand` — string constants for all used DMCC commands; `SanitizeForDmcc(s)` replaces
  characters illegal in DMCC job/operator strings (`& < > \r \n`) with `_`.
- `DmccResponse` — parses the `\r\n<status>\r\n[\r\n<body>\r\n]` envelope; exposes `StatusCode`,
  `IsSuccess`, `Body`, `IsXml`.
- `DmccStatus` — integer constants for well-known DMCC status codes.

### Idle-gap Read Strategy
`NetworkStream.ReadAsync` with a `CancellationToken` disposes the underlying socket on Linux
(.NET 8) when the token is cancelled, killing the connection. `DmccClient` therefore runs a
synchronous `Socket.Receive` loop on a thread-pool thread; `Socket.ReceiveTimeout` (idle-gap
interval) causes a `SocketError.TimedOut` exception that is caught to detect end-of-response
without socket teardown.

### DMST XML Parser (`DeviceInterface.Dmst`)
- `VerificationXmlMap` — maps DataMan DMST XML element names to `VerificationRecord` field names;
  `ClassifySymbology()` returns the `SymbologyFamily` enum value.
  Now includes `SCRlRd` element (Symbol Contrast Reflection Level / Reflection Difference pair).
- `DmstResultParser.Parse(xml, map, contextRecord)` — parses DataMan verification result XML
  into a fully-populated `VerificationRecord`, merging device context (serial, operator, etc.)
  from the supplied context record.  Handles both 2D (ISO 15415) and 1D (ISO 15416) payloads.
  Wires `SC_RlRd` through the record from the `<SCRlRd>` XML element.

### DMST Push Script (`DeviceInterface/Dmst/DmstPushScript_v1.js`)
Ready-to-paste JavaScript for **Format Data → Scripting → Data Formatting** in DMST.
Replaces the default plain-text output with a complete
`<DMCCResponse><DMSymVerResponse>…</DMSymVerResponse></DMCCResponse>` XML push that
`DmstResultParser` can fully parse, producing a 167-column row in XLSX.

**How to install:**
1. Open DMST → Format Data → Scripting → *Data Formatting* tab
2. Select the **Script-Based Formatting** radio button
3. Paste the entire contents of `DmstPushScript_v1.js`
4. Click **Save Settings → Write to device**

**Coverage (firmware 6.x / DMV475):**

| Section | XML elements emitted |
|---|---|
| Identity | `DateTime`, `SymbologyName`, `DecodedData` |
| Grade summary | `FormalGrade`, `OverallGrade`, `OverallGradeNumeric` |
| Verification settings | `ApertureRef`, `Wavelength`, `Lighting`, `Standard` |
| 2D ISO 15415 quality | `UECPercent/Grade`, `SCPercent`, `SCRlRd`, `SCGrade`, `MODGrade`, `RMGrade`, `ANUPercent/Grade`, `GNUPercent/Grade`, `FPDGrade`, `DecodeGrade`, `AGValue/Grade` |
| Matrix characteristics | `MatrixSize`, `HorizontalBWG`, `VerticalBWG`, `EncodedCharacters`, `TotalCodewords`, `DataCodewords`, `ErrorCorrectionBudget`, `ErrorsCorrected`, `ErrorCapacityUsed`, `ErrorCorrectionType`, `NominalXDim`, `PixelsPerModule`, `ImagePolarity`, `ContrastUniformity`, `MRD` |
| Quiet zones | `LLSGrade`, `BLSGrade`, `LQZGrade`, `BQZGrade`, `TQZGrade`, `RQZGrade` |
| Transition ratios | `TTRPercent/Grade`, `RTRPercent/Grade`, `TCTGrade`, `RCTGrade` |
| Quadrant (≥32×32) | `ULQZ…`, `URQZ…`, `RUQZ…`, `RLQZ…`, per-quadrant TTR/RTR/TCT/RCT |
| 1D ISO 15416 | `SymbolAnsiGrade`, `AvgEdge/RlRd/SC/MinEC/MOD/Defect/Dcod/DEC/LQZ/RQZ/HQZ/MinQZ`, `BWGPercent`, `Magnification`, `Ratio`, `NominalXDim1D` |
| 1D per-scan | `<ScanResults><Scan number="n">…</Scan></ScanResults>` (max 10 scans) |

The script is ECMAScript 5-compatible (no `const`/`let`, no arrow functions) and uses
XML entity escaping so barcodes containing `<`, `>`, or `&` survive transit intact.
All property accesses use a null-safe `prop()` helper — unknown paths emit empty elements
rather than crashing the script.

**Verifying firmware property names:** if a column arrives blank after enabling the script,
check the annotated property-name comments in the `.js` file against the *Scripting API
Reference* (DMST Help menu) for your exact firmware revision.

### Device Session (`DeviceInterface.DeviceSession`)
- Connects to a DataMan device, queries `DeviceInfo` (type, firmware, serial, name, calibration
  date), and exposes `TriggerAndGetResultAsync(contextRecord)` for the Poll acquisition mode
  (TRIGGER + GET SYMBOL.RESULT).

### MockDmccServer (`DeviceInterface.Testing`)
- Loopback `TcpListener` on an OS-assigned port for unit tests; exposes static sample XML payloads
  (`SampleDm2DXml`, `SampleUpcAXml`) matching what a real DataMan device sends.

---

## Schema

`TruCheckCompatibleSchema` defines 167 columns across nine blocks:

| Block | Range     | Content                                          |
|-------|-----------|--------------------------------------------------|
| A     | 1–20      | Common header (date, operator, device, ...)      |
| B     | 21–42     | 1D ISO 15416 summary parameters + QZ averages    |
| C     | 43–45     | 1D general characteristics (BWG, magnification)  |
| D     | 46–96     | 2D ISO 15415 quality parameters + QZ grades      |
| E     | 97–115    | 2D general characteristics (matrix, ECC, ...)    |
| F     | 116–122   | Common grading summary (formal grade, overall)   |
| G     | 123–132   | Custom/user fields, notes, session info          |
| H     | 133–138   | Device/calibration metadata                      |
| I     | 139–163   | GS1 Data Format Check (DFC_Standard + 8 rows)   |

---

## Output File Naming

| Condition            | Filename                              |
|----------------------|---------------------------------------|
| JobName set          | `{JobName}_{YYYY-MM-DD}.xlsx`         |
| No job, OperatorId   | `VTCCP_{OperatorId}_{YYYY-MM-DD}.xlsx`|
| Neither set          | `VTCCP_{YYYY-MM-DD}.xlsx`             |

Filesystem-illegal characters and spaces in job/operator names are replaced with `_`
by `ExcelFileManager.SanitizeFileName()`.

Default output directory: `%USERPROFILE%\Documents\VTCCP`

---

## Session Sidecar (JSON)

A `.vtccp.json` file is written alongside the Excel output whenever a session is open.
It stores job name, operator, roll number, batch, record count, and format so an
interrupted session can be resumed automatically via `StartSession()`.

The sidecar is deleted by `CloseSession()` to signal a clean job close.

---

## Roll Identifier Modes

Set `SessionState.RollIncrementMode` before calling `StartSession()`:

| Mode | Behaviour | `{Roll}` token |
|---|---|---|
| `Manual` *(default)* | Caller supplies roll value; `SetNewOperatorAndRoll(op, manualRoll: n)` required | Decimal integer |
| `AutoIncrement` | Starts at `RollStartValue`; increments by 1 on each `SetNewOperatorAndRoll(op)` call | Decimal integer |
| `DateTimeStamp` | `yyyyMMddHHmmss` generated at session open and on each `SetNewOperatorAndRoll(op)` call | 14-char timestamp |

All three modes survive a crash and resume correctly via the sidecar.

---

## Building and Running

```
cd vtccp
dotnet build VTCCP.sln          # build all projects
dotnet run --project TestHarness/TestHarness.csproj   # run the test harness
```

Test output is written to `%TEMP%/vtccp_*_output/` (Linux) or `%TMP%\vtccp_*_output\` (Windows).
Expected output ends with all checks reporting **PASS**.

---

## EPPlus Licensing

EPPlus is used under the NonCommercial context (`ExcelPackage.LicenseContext = NonCommercial`).
Production or commercial deployments require a valid EPPlus commercial license.

---

## Roadmap

| Phase | Status      | Description                                              |
|-------|-------------|----------------------------------------------------------|
| 1     | Complete    | Excel Engine (offline) — 167-column schema, session mgmt |
| 2     | Complete    | Device Integration — DMCC/DMST live connection           |
| 3     | Complete    | Config Templates GUI — WPF shell + ConfigEngine          |
| 4     | Complete    | Results History — live DataGrid, grade/pass-fail filters |
