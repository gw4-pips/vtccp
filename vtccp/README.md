# VCCS DMV TruCheck Command Pilot (VTCCP)

A native Windows desktop utility (WPF + C#/.NET 8) that replicates the Webscan TruCheck
Excel verification-results logging (XLS/XLSX) for Cognex DataMan DMV barcode verifiers.

---

## Architecture Overview

```
vtccp/
├── ExcelEngine/          — Class library (Phase 1)
│   ├── Adapters/         — IExcelAdapter, XlsxAdapter (EPPlus), XlsAdapter (NPOI)
│   ├── Models/           — VerificationRecord, SessionState, ScanResult1D, ...
│   ├── Schema/           — ColumnSchema, WebscanCompatibleSchema (163 columns), SchemaVersionWriter
│   ├── Session/          — SessionManager, SessionSidecar
│   └── Writer/           — ExcelWriter, ExcelFileManager, DataMatrix2DMapper,
│                           ISO15416Mapper, PerScanTableWriter, ElementWidthsWriter
├── TestHarness/          — Console test driver (Tasks 1–4)
└── README.md
```

---

## Phase 1 — Excel Engine (Offline)

### Task 1 — .NET scaffold & data models
- .NET 8 solution with ExcelEngine class library and TestHarness console project.
- All domain models: `VerificationRecord`, `SessionState`, `GradingResult`, `ScanResult1D`,
  `DataFormatCheckResult`, `ElementWidthData`, enumerations.
- `WebscanCompatibleSchema` — 163-column schema matching Webscan TruCheck output
  (Blocks A–I).

### Task 2 — 2D Data Matrix Excel engine
- `IExcelAdapter` abstraction; `XlsxAdapter` (EPPlus) and `XlsAdapter` (NPOI).
- `ExcelWriter` — title row, header row, data rows, GS1 DFC columns, append mode,
  row-limit check, `.xls` near-limit warning.
- `ExcelFileManager` — Webscan-convention filename generation (`{Job}_{YYYY-MM-DD}.xlsx`),
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
  - `SetNewOperatorAndRoll(operatorId)` — increments roll number mid-session.
- `SessionSidecar` — compact JSON snapshot (`{outputFile}.vtccp.json`) for resume support.
- `SchemaVersionWriter` — writes three metadata cells (marker / schema name / version) into
  row 1 past the last data column so downstream tooling can identify VTCCP-generated files.
- Full 6-record integration test: 3× GS1 DataMatrix, 1× DataMatrix, 1× UPC-A, 1× EAN-13.

---

## Schema

`WebscanCompatibleSchema` defines 163 columns across nine blocks:

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

Matches the Webscan TruCheck convention:

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

## Running the Test Harness

```
cd vtccp
dotnet run --project TestHarness/TestHarness.csproj
```

Expected output ends with all checks reporting **PASS**.

---

## EPPlus Licensing

EPPlus is used under the NonCommercial context (`ExcelPackage.LicenseContext = NonCommercial`).
Production or commercial deployments require a valid EPPlus commercial license.

---

## Roadmap

| Phase | Status      | Description                                              |
|-------|-------------|----------------------------------------------------------|
| 1     | Complete    | Excel Engine (offline) — 163-column schema, session mgmt |
| 2     | Planned     | Device Integration — DMCC/DMST live connection           |
| 3     | Planned     | Config Templates GUI (WPF)                               |
