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
1. In DMST, click **Format Data** in the Application Steps sidebar (left panel)
2. On the **Basic** tab: select the **Script-Based Formatting** radio button
   *(it sits below the "Basic Formatting" radio — confirm any DMST warning prompt)*
3. Click the **Scripting** tab at the top of the Format Data panel
4. Paste the entire contents of `DmstPushScript_v1.js` into the editor
5. Click **Save Settings → Write to device**

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

The file is marked with the NTFS `Hidden` attribute immediately after each write so
it does not appear in a normal Explorer folder view.  To see it: *View → Hidden items*
in Explorer, or `dir /a:h` in a command prompt.

**Implementation note — attribute handling:**  
`SaveSidecar()` calls `File.SetAttributes(path, FileAttributes.Normal)` *before*
every write (if the file already exists), then sets `FileAttributes.Hidden` *after*.
This ordering is mandatory.  `File.WriteAllText` opens with `FileMode.Create`; if the
file carries a stale `ReadOnly` attribute from a prior write, `WriteAllText` throws
`UnauthorizedAccessException`.  The ORing pattern `GetAttributes | Hidden` was the
original bug — it preserved any `ReadOnly` bit and caused the crash on the second sidecar
write.  Always set `FileAttributes.Normal` to clear, then set only `FileAttributes.Hidden`.

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

## Operational Notes & Field Observations

This section records behaviour confirmed or discovered during live testing against
a real DMV475 device.  It is the authoritative source for configuration requirements
that are not obvious from the DMST UI.

---

### Operator ID Persistence

The **Operator ID (override)** field in the Session Launcher is pre-filled with the
value that was typed the last time a session was started.  It is persisted in
`%APPDATA%\VTCCP\appsettings.json` under `lastOperatorId`.  
- To change operators: type a new value before clicking **Start Session**.  
- To clear it permanently: delete or blank the field before starting, or edit
  `appsettings.json` directly.

---

### Live Excel Monitoring (Real-Time Save)

`SessionManager.AddRecord()` calls `ExcelWriter.Save()` after every appended row.
The file can therefore be opened in Excel during a session and watched live.

**How this differs from Webscan:**  
Webscan writes an `.xls` (legacy binary format) file via an OLEDB / COM-layer
connection that keeps a shared file handle open for the duration of the session.
Excel recognises OLEDB sources and can auto-refresh the sheet without any prompt.

VTCCP writes `.xlsx` (Open XML ZIP archive) via EPPlus.  Each `Save()` rewrites
the entire ZIP file and closes it.  When Excel has the same `.xlsx` file open:

| Excel state | What you see |
|---|---|
| File open, no edits in progress | Yellow notification bar: *"This file has changed. Reload?"* — click **Reload** to see new rows. |
| File open, you are editing a cell | Excel holds its own copy; the disk version is updated silently. You will see a reload prompt when you next save or switch away. |
| File not open | No prompt; next time you open it all rows are present. |

The notification bar appears in Excel 2016 and later.  Earlier versions may show
a modal dialog instead.  In all cases the data is correctly on disk — open/reload
the file to see the current rows.

**Recommendation:** Open the file in Excel *after* starting the session and use
the yellow bar to refresh between scans.  Do not edit the file while a session is
running; edits will be lost when Excel reloads the EPPlus-saved copy.

If silent live updates (Webscan style) are a hard requirement in a future version,
the `XlsAdapter` (NPOI) path can be extended to use an OLEDB write-behind strategy,
but this requires Excel to be installed on the host machine.

---

### Software Trigger in Push Mode

In Push mode VTCCP does not maintain a persistent DMCC connection.  The
**⚡ Trigger Scan** button opens a short-lived TCP connection to port 23, sends the
`TRIGGER` command, reads and surfaces the device's response code, then disconnects.

#### Prerequisites for the software trigger to work

1. **Device trigger type must be set to "Single"** (or another type that honours
   software trigger commands).  In DMST:
   *Device Settings → Trigger → Trigger Type → Single*  
   Trigger types "Presentation" and "Self" ignore DMCC `TRIGGER` commands entirely.

2. **DMST must not be connected to the device at the moment the button is pressed.**  
   Port 23 accepts only one DMCC client at a time.  If DMST is open and connected,
   VTCCP's connection attempt will be refused and the status bar will show:
   `Trigger error: No connection could be made because the target machine actively
   refused it.`  
   Close DMST (or disconnect its Telnet session) before using the software trigger.
   Alternatively, use the physical trigger on the device directly — the push result
   arrives via the Network Client regardless of how the scan was initiated.

3. **A barcode must be in the field of view** when the trigger fires, otherwise the
   device returns DMCC status code 6 (No Read) and the status bar shows:
   `Trigger fired — no symbol in field of view.`

#### DMCC TRIGGER response codes

| Code | Meaning | VTCCP status bar |
|------|---------|-----------------|
| 0 | Trigger accepted; device is scanning | "Trigger sent — waiting for push result…" |
| 6 | Trigger fired; no symbol decoded | "Trigger fired — no symbol in field of view." |
| 8 | Device busy; trigger rejected | "Device busy — trigger rejected. Wait a moment and retry." |
| Other | Firmware-specific error | "Trigger: device returned code N — *body*" |

If the status bar shows code 0 but no push result arrives (Records stays at 0),
see the **Push Mode Network Setup** section below.

---

### Push Mode Network Setup

Push mode requires the DataMan device's **Network Client** to be configured to
push XML results to VTCCP's host PC.

**In DMST:**
1. *Device Settings → Communication → Network Client*
2. Set **IP Address** to the VTCCP host PC's IP (e.g. `10.10.10.19`)
3. Set **Port** to the `DmstListenPort` value in the VTCCP device profile (default `9004`)
4. Set **Format** to *Full XML* (or ensure the DMST push script is installed — see above)
5. Enable the Network Client checkbox
6. *Save Settings → Write to device*

**Windows Firewall:**  
VTCCP's push listener binds to `0.0.0.0:<DmstListenPort>`.  On a fresh Windows
install the Windows Defender Firewall will block inbound TCP on that port.  Either:
- Allow VTCCP through the firewall when prompted on first session start, or
- Manually add an inbound rule: *Control Panel → Firewall → Advanced Settings →
  Inbound Rules → New Rule → Port → TCP → Specific port: 9004 → Allow*

**Confirming the push path:**  
Use DMST's *Network Client → Test Connection* button (if available in your firmware)
or scan a barcode while watching the VTCCP session for a record increment.

---

### DMST Push Script vs. Default XML Format

If `DmstPushScript_v1.js` is NOT installed, the device pushes a minimal XML payload
that `DmstResultParser` can still parse for identity and overall-grade columns, but
the majority of the 167-column schema will be blank.

Install the script for full column coverage (see the **DMST Push Script** section
above and the installation steps in the script's header comments).

---

## Roadmap

| Phase | Status      | Description                                              |
|-------|-------------|----------------------------------------------------------|
| 1     | Complete    | Excel Engine (offline) — 167-column schema, session mgmt |
| 2     | Complete    | Device Integration — DMCC/DMST live connection           |
| 3     | Complete    | Config Templates GUI — WPF shell + ConfigEngine          |
| 4     | Complete    | Results History — live DataGrid, grade/pass-fail filters |
