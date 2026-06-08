# VTCCP Working Notes

> **Restore point**: commit `f474fa7` — AS-IS state as of 2026-05-29.
> All trigger investigation work begins AFTER this point.
> Do not build without explicit user instruction.

---

## RESOLVED: Image missing from TC/DMST screen post-verification

**Status**: FULLY RESOLVED 2026-05-31 — `CONFIG.DEFAULT + CONFIG.SAVE + REBOOT` on the device.
No code change needed; `DataManSdkClient.cs` no longer touches LIVEIMG.MODE.

**Symptom**: After a verification scan, the DMST TC screen and DMST verification panel
show no barcode image. HTML report still contains the image. Push XML `JpegImageBase64`
also confirmed populated. This is a result-channel delivery failure, not a capture failure.

**Confirmed root cause** (2026-05-31): One or more NVRAM parameters were corrupted by an
earlier VTCCP session that called the SDK's own `COM.DMCC-SAVE` mechanism (via
`SetResultTypes()` before that call was removed). `CONFIG.DEFAULT` resets all config
parameters to firmware factory defaults; `CONFIG.SAVE` persists them; `REBOOT` applies them.
The exact corrupted parameter was not identified — the fix is a full config reset.

**CRITICAL — correct known-good state** (confirmed from DM-KnownGood-Snapshot_2026-05-31):
- `LIVEIMG.MODE = 0` ← this is CORRECT; it is NOT 2
- `DATA.IMAGE-TYPE = 0` ← correct
- `DATA.RESULT-TYPE = 513` ← factory default for DM475V
- `TRIGGER.TYPE = 0` ← external single trigger
- `IMAGE.SIZE = 1`

**All prior LIVEIMG.MODE theories were wrong.**  
The SDK's `Connect()` does NOT set LIVEIMG.MODE to any value that causes image loss.
The root cause was a different NVRAM parameter written by a prior `COM.DMCC-SAVE`.
`LIVEIMG.MODE = 0` with the TC panel image working has been device-confirmed.

**False leads definitively ruled out**:
- LIVEIMG.MODE — value is 0 in known-good state; setting it to 2 achieves nothing
- `DATA.IMAGE-TYPE` — confirmed = 0 (correct) in known-good state; never the cause
- SDK `Connect()` setting LIVEIMG.MODE — CONFIRMED FALSE; no such effect observed on fw 6.1.16_sr4
- Port 23 blocked by DMST — CONFIRMED FALSE; VTCCP trigger on port 23 works while DMST is open;
  both can use port 23 simultaneously

**Recovery procedure** (if image disappears again):
1. Open Telnet/DMCC to device port 23
2. `CONFIG.DEFAULT` — resets all config to factory
3. `CONFIG.SAVE` — persists to NVRAM
4. `REBOOT` — applies the reset
5. After reboot: restore any custom config (TRIGGER.TYPE, aperture, lighting, etc.)

---

### Probe — Telnet to port 23 on device IP (10.10.10.7)

Run each `GET` command in sequence. Record the actual response against the expected value.

```
GET LIVEIMG.MODE
GET DATA.IMAGE-TYPE
GET DATA.RESULT-TYPE
GET DATA.RESULT-ENCODING
GET DATA.RESULT-ALWAYSSEND
GET IMAGE.FORMAT
GET IMAGE.SIZE
```

**Expected values (known-good state — confirmed 2026-05-31):**

| Command | Expected | Notes |
|---|---|---|
| `GET LIVEIMG.MODE` | `0` | Correct known-good value; NOT 2. TC panel image works with LIVEIMG.MODE=0. |
| `GET DATA.IMAGE-TYPE` | `0` | Correct. If non-zero, SDK has persisted via COM.DMCC-SAVE. |
| `GET DATA.RESULT-TYPE` | `513` | DM475V factory default; not 0. |
| `GET DATA.RESULT-ENCODING` | `0` | Default encoding. |
| `GET DATA.RESULT-ALWAYSSEND` | `1` | Results sent on every trigger. |
| `GET IMAGE.FORMAT` | `1` | JPEG. |
| `GET IMAGE.SIZE` | `1` | Confirmed in known-good snapshot. |

**If image panel goes blank again**, use the CONFIG.DEFAULT recovery procedure above.
Do NOT set LIVEIMG.MODE to 2 — this is not the fix.

---

## Active investigation: Trigger reset / DMST TC recovery

**Status**: PARKED — plan written, no code changed. Awaiting user to restart.

**The problem**: After VTCCP fires a CP Trigger Scan (Push mode), the DMST TC window
does not recover to live/ready mode after the scan completes.

**The plan**: See `.agents/memory/trigger-reset-plan.md` for the full investigation
sequence (one variable at a time) and the code locations that will need to change.

**Do not skip to implementation** — the Wireshark capture and baseline observation
steps must happen first.

---

## Parked issues

### Excel live-view via COM automation (Method D)

**Status**: IMPLEMENTED — `ExcelEngine.Adapters.ComExcelAdapter` — 2026-06-01.

`ComExcelAdapter` uses late-bound COM (`dynamic` + `Marshal.GetActiveObject`) to write
rows directly into a running Excel process holding the output XLSX.  No file I/O; no lock
contention; rows appear in the spreadsheet immediately as scans arrive.

**Behaviour at session start**:
- `SessionManager.StartSession` calls `ComExcelAdapter.TryAttach(outputPath)`.
- If Excel is running and has `outputPath` open: COM adapter selected → `[VTCCP-EXCEL] Adapter: COM (live Excel)` in debug output.
- If Excel is not running, file not open, or OS is non-Windows: falls back to EPPlus → `[VTCCP-EXCEL] Adapter: EPPlus (file-based)`.

**Key design decisions**:
- Late binding (`dynamic`) — no compile-time dependency on `Microsoft.Office.Interop.Excel`.
- `[SupportedOSPlatform("windows")]` on the class; caller guards with `OperatingSystem.IsWindows()`.
- `Save()` calls `_workbook.Save()` (Excel saves itself — no file lock issue).
- `Dispose()` does NOT close the workbook — the operator keeps Excel open.
- `WriteEmbeddedImage` / `WriteLogoImage`: write to temp file, `Shapes.AddPicture`, delete temp.
- `ExcelFileManager.CheckFileLocked` at session start still fires for new sessions (file does not yet exist); no conflict.

**Fallback**: if COM attach fails for any reason (TryAttach returns null), EPPlus takes over transparently.  The `[VTCCP-EXCEL] Adapter:` log line confirms which path was selected.

---

### Device Repair Sequence — "CP Reset"

**Status**: LOGGED CONCEPT — do not build without explicit instruction.

**Idea**: A single VTCCP command ("Repair Device" or "Restore Baseline") that replays
the exact DMCC SET sequence needed to put the DM475V back into a known-good VCCS
operating state — including all TruCheck configuration, imaging settings, and the
current JS push script. Targeted at advanced application repair without needing DMST.

**What the repair sequence would do (proposed order):**
1. Open a brief raw TCP connection to port 23 (same path as Push-mode trigger)
2. Issue DMCC SET commands to restore all required TruCheck settings to VCCS baseline:
   - `SET TRIGGER.TYPE 0` (external / single, VCCS baseline)
   - `SET IMAGE.FORMAT 1` (JPEG)
   - `SET IMAGE.SIZE 1` (confirmed known-good value — do NOT set LIVEIMG.MODE; known-good = 0)
   - All grading standard, aperture, wavelength, lighting, application standard params
     (list TBD — requires TC config screenshots per the VCCS Command Pilot feature plan)
3. Write the current production JS push script to the device via DMCC
   (exact SET key TBD — see "Push script auto-deploy on connect" below)
4. Verify: GET each key back and confirm value matches expected; report any mismatch
5. Close port-23 connection

**Why this is the right architecture:**
- Every Telnet command the user sent was already a raw DMCC SET on port 23
- VTCCP owns the port-23 TCP channel (`SendRawDmccAsync` / Push-mode trigger path)
- No DMST dependency — the repair runs even if DMST is not installed
- The repair sequence is idempotent — running it twice has no side effects
- Could be triggered from the UI, from a CLI flag, or automatically on connect
  if a config drift is detected

**Blocking unknowns (same as VCCS Command Pilot feature):**
1. **Full TruCheck DMCC key list** — requires detailed DMST TC window screenshots
   (all panels) to map every visible setting to its DMCC key. Cannot write the
   repair SET list without this.
2. **DMCC key for push script delivery** — the SET key that writes the JS script
   text to the device (see "Push script auto-deploy" item below). Must confirm
   before that step can be included.

**Relationship to other features:**
- Overlaps with the VCCS Command Pilot Full Device Configuration feature (same
  DMCC key research dependency)
- Tighter scope: repair sequence is a fixed write-only restore, not a full
  read/edit/write config GUI
- The push script delivery step is the same capability as "Push script auto-deploy"
- Could be implemented incrementally: ship the LIVEIMG.MODE + IMAGE.* + TRIGGER.TYPE
  restore first (all keys confirmed), then add TruCheck params as screenshots arrive

---

### Push script auto-deploy on connect

**Status**: PARKED — log only, no implementation until user directs.

Every VCCS Command Pilot installation currently requires the operator to manually
copy the push script into DMST → TruCheck Configuration → Event window. This is
a setup step that can fail silently if the wrong version is pasted.

**Requirement**: when VTCCP connects to a DataMan device, it should automatically
write the current production push script to the device via DMCC so the device is
always running the correct version — no manual paste required.

**Implementation path when approved**:
- Identify the DMCC key for the event/push script (likely in the DMCC guide A1 digest
  under the `TRIGGER.*` or `EVENT.*` namespace — exact key to confirm before coding)
- On `ConnectAsync`: after device info is read, issue a DMCC SET with the bundled
  script text; compare version tag first to avoid unnecessary writes
- Bundle the canonical push script as an embedded resource in VtccpApp
  (same pattern as `tessdata/eng.traineddata`) so no file path is needed at runtime
- Version check: read back the DiagTag from a lightweight GET after SET to confirm
  the device accepted the script correctly

**Blocked by**: Exact DMCC key confirmation + user approval.

---

### GS1 `<F1>` formatter — ]d1 vs ]d2

**Status**: PARKED at user request, 2026-05-29.

The user will re-demonstrate the failure. Do not draw conclusions about build state,
timing, or code correctness until they do. Do not revisit until user reopens.

---

## Confirmed working rules for this project

1. **DO NOT BUILD WITHOUT ASKING.** Planning and notes only unless user says to implement.
2. **Assume less, prove more.** Label every assumption. Confirm with device before acting.
3. **One variable at a time.** One change, observe, record, then move to next.
4. **AutoPoll mode** — ignore entirely for now.
5. **Quit button** — planned but not built. Downstream of trigger investigation.

---

## Manual vs Push mode — what is known vs assumed

### Confirmed facts (from code + user reports)

- **Manual mode**: VTCCP opens persistent DMCC connection on port 44444 via SDK.
  DMST must be closed (conflicts on port 44444). HTTP subscriber started on same port.
  CP Trigger Scan fires via `TriggerAndGetResultAsync` (SDK path).

- **Push mode**: VTCCP does NOT hold a persistent DMCC connection. DMST may remain
  open. CP Trigger Scan fires TRIGGER ON via a brief raw TCP connection to port 23,
  then closes immediately. Result arrives via HTTP subscriber on port 44444.
  `HttpEventSubscriber` is started on `StartHttpSubscriberAsync`.

- **TRIGGER.TYPE** is read (GET) on Manual mode connect and stored. It is NOT changed
  by VTCCP in either mode (SET is commented out). Restore is also commented out.

### NOT confirmed — needs observation

- Whether DMST TC window actually conflicts with Manual mode DMCC connection in practice.
- Exact DMST TC window state (frozen? grayed? error?) when trigger recovery fails.
- Whether the HTTP subscriber's open connection causes DMST to change its behaviour.
- What TRIGGER.TYPE value (if any) DMST sets when entering or leaving its live-feed loop.
- Whether the problem occurs on every Push scan or only on the first/subsequent ones.

---

## Quit button plan

- A Quit button with active-session warning ("Session in progress. Stop and exit?") is
  confirmed as a planned UI feature.
- The window X (Closing event) must run identical cleanup to the Quit button.
- Cleanup sequence: Stop → Disconnect → Restore TRIGGER.TYPE → allow close.
- **NOT to be built until trigger reset is stable and user approves.**

---

## FEATURE DESIGN — Cursor-Aware Insert (WTC-style Excel mode)

**Logged**: 2026-06-07  
**Status**: Design only — not yet scoped for implementation  
**Prerequisite**: COM adapter fully working (needed for `ActiveCell` access)

---

### Background

The WTC (Webscan TruCheck) Excel file stores a **next-available-row pointer** at
**col 111, Excel row 2** (the visible header row). In the template file this reads `3`
(first data row). The macro increments it by 1 after every scan. This is a dumb append
pointer — Webscan never went further.

PIPS OptiDoc (the original, conceived by VCCS, built by Viktor) implemented a
**cursor-aware insert** that Webscan never replicated. VTCCP will implement this as
a first-class feature when the WTC-style Excel mode is built.

---

### Feature: Cursor-Aware Insert

**Normal operation (append mode):**
- On each scan result, write to the row indicated by the stored pointer (col 111, row 2).
- Increment the pointer by 1.
- Store the updated pointer back to col 111, row 2.

**Correction mode (operator-initiated):**
1. Operator decides a previously captured row needs to be re-scanned (bad barcode,
   wrong sample, etc.).
2. Operator deletes the contents of that row in Excel (clearing the cells).
   - Design decision (retained from OptiDoc): **require explicit delete** rather than
     silent overwrite. Forces conscious intent. Silent overwrite was considered and
     rejected — keeps the operator accountable.
3. Operator positions the Excel cursor (ActiveCell) on that now-empty row.
4. Next scan: VTCCP detects that `ActiveCell.Row` ≠ pointer row AND that
   `ActiveCell` row is empty → writes to `ActiveCell.Row` instead.
5. **The stored pointer is NOT modified** during correction mode. It retains the
   "resume" position (the next row that would have been written in normal sequence).

**Multiple consecutive corrections:**
- After writing a correction scan, VTCCP checks the new `ActiveCell.Row`:
  - If the new ActiveCell row is **empty** → still in correction mode, write there next.
  - If the new ActiveCell row is **non-empty** (operator moved cursor back into data,
    or to the pointer row) → exit correction mode, revert to stored pointer.
- This allows the operator to make several delete/re-scan corrections in sequence
  before returning to normal capture — all without touching the pointer.

**Resuming normal capture:**
- When cursor moves off an empty correction row (or operator does nothing and the
  next scan fires normally), VTCCP detects `ActiveCell` is at the stored pointer row
  (or is non-empty) and reverts to append mode automatically.
- No explicit "exit correction mode" button needed.

**Open design question:**
- Should "require delete" be a session-level setting (default ON) or always enforced?
  Overwrite-without-delete is simpler but removes the intentional friction.
  Decision deferred until implementation is scoped.

---

### Pointer cell location (VTCCP convention)

**VTCCP will use Col A (column 1), hidden row 2.**

- WTC used col 111 (far right, obscure). VTCCP moves it to col A for predictability.
- Row 2 is hidden — not visible to the operator during normal use, but trivially
  unhide-able in Excel if manual inspection or reset is needed.
- A small read-only indicator in the VTCCP UI will show the current pointer value
  (e.g. "Next row: 47") so the operator never needs to dig into the spreadsheet.
- A "Reset Row Counter" button in VTCCP (exact placement TBD) resets col A / row 2
  to the first data row (row 3) — useful when starting a new pass on the same file.

**Open**: whether the counter lives only in the Excel cell or is mirrored in the
session sidecar JSON. Mirroring in JSON allows VTCCP to recover the pointer if the
Excel file is closed/reopened mid-session. Decision deferred.

---

### Implementation sketch (COM path only)

```
ExcelApplication app = GetActiveObject("Excel.Application");
Worksheet ws = app.ActiveSheet;
int pointerRow = (int)ws.Cells[2, 1].Value;   // Col A (1-indexed), hidden row 2
int cursorRow  = app.ActiveCell.Row;

bool cursorEmpty = IsCellRowEmpty(ws, cursorRow);  // check key data columns

int targetRow;
if (cursorRow >= 3 && cursorEmpty && cursorRow != pointerRow)
{
    targetRow = cursorRow;   // correction mode — write here, don't touch pointer
}
else
{
    targetRow = pointerRow;  // append mode — write here, increment pointer
    ws.Cells[2, 112].Value = pointerRow + 1;
}
WriteRecordToRow(ws, targetRow, record);
```

---

## WTC UI Reference — Screenshots Logged 2026-06-07

Screenshots archived in `attached_assets/` (image_1780846508356, image_1780850312245).

### "Webscan Excel Spreadsheet Functions" main dialog

| State | Display |
|---|---|
| No job open | "No Job is Open" / "No Operator is defined" |
| Job active | Job Name, Operator Name/Number, Roll Number shown |

**Buttons**: New Job · Open Job · Close Job · Reclaim / Archive Job · Close and Archive Job ·
Modify Template Spreadsheet · Set New Operator/Roll · Close Window · **Go Live**

"Go Live" is the WTC equivalent of VTCCP's **Start Session** — activates real-time
Excel capture. "Modify Template Spreadsheet" opens the blank template for layout editing.

### "New Job" dialog

Fields:
- Job Name / Number (free text)
- Operator Name / Number (free text)
- Roll Number (free text)

Button: **Create Spreadsheet** (creates the session Excel file from the template, opens it).

**VTCCP mapping**: New Job ≈ Start Session. Fields map 1:1 to our existing
`JobName`, `OperatorId`, `RollNumber` session fields. "Create Spreadsheet" ≈
`SessionManager.StartSession()` → file creation + COM open.

Key difference: WTC always creates a new file from template on "New Job". VTCCP
also supports resuming an existing file ("Open Job" equivalent not yet built).

---

## FEATURE DESIGN — Three-Level Excel Row Structure (2026-06-07)

> **STATUS — Level 1 IMPLEMENTED (2026-06-07, v1.2.2)**
>
> Files changed:
> - `IExcelAdapter.cs` — added `SetRowOutlineLevel`, `SetRowHidden`, `ScheduleRowHide`
> - `ComExcelAdapter.cs` — implemented all three; `ScheduleRowHide` spins a background STA thread
> - `XlsxAdapter.cs` — implemented all three; `ScheduleRowHide` is a no-op (file-based)
> - `XlsAdapter.cs` — `SetRowOutlineLevel` is a no-op (NPOI `IRow.OutlineLevel` read-only in this
>   NPOI build); `SetRowHidden` uses `ZeroHeight`; `ScheduleRowHide` no-op
> - `ParseDetailRowWriter.cs` — new class; writes sentinel "↳" + HRI string from `DataFormatCheckResult.Rows`
> - `ExcelWriter.cs` — added `ParseDetailRowWriter`, `LastParseDetailRow` property;
>   Level-1 row inserted after parent `_nextDataRow++` / before per-scan rows
> - `SessionManager.cs` — checks `LastParseDetailRow` after each `AppendRecord`; calls
>   `ScheduleRowHide` with `ParseDetailHideDelay = 20 s` (static field, easy to make configurable)
>
> **Actual HRI format** (revised 2026-06-07 — brevity-first):
> `]d1 | GS1 | Header | GTIN: 0123456789012 | BatchLot: A1234`
>
> Format rules:
> - Lead with AIM ID (`record.SymbologyId`), e.g. `]d2`
> - Abbreviated standard: "GS1 Application Data Format" → "GS1", "MIL-STD-130" → "MIL-130"
> - `AI:*` rows skipped (redundant AI code numbers)
> - `Chk Digit` row skipped (internal GS1 decomposition)
> - `GS1 Header` emitted as bare `Header` token, no `<F1>` data value (encoding artifact)
> - All other rows → `{Name}: {Data}`, pipe-separated
>
> **Deviation from design**: The design called for GS1 syntax engine `HRI` property
> (giving `"GTIN (01) 00355513710213"` format with AI number in parens). The implementation
> instead reads `DataFormatCheckResult.Rows[].Name` + `.Data` directly, because the DFC
> result is already parsed into that model by the time `AppendRecord` is called. The GS1
> engine re-parse path would require the raw `DecodedData` and the engine instance, which
> are not available at the write layer. The current format is readable; if the AI-code-in-
> parens format is needed later, add an `AiCode?` field to `DataFormatCheckRow` and populate
> it from the HTML scraper's DFC table.
>
> **DESIGN DECISION LOGGED (2026-06-07)** — HTML report as primary DFC source:
> The DMST `pcm_report.html` contains the full Data Format Check table as rendered by
> Cognex TruCheck itself — AI descriptions, AI codes, encoded data, and pass/fail status,
> in the exact format the operator sees on screen. The preferred architecture is:
>
> 1. **Primary** — extend `DmstHtmlScraper.ParseHtml()` to scrape the DFC table directly
>    from the HTML report. This gives us the official Cognex-rendered AI names + codes
>    (solving the `AiCode` gap) with no dependency on the GS1 syntax engine or push XML
>    DFC schema. Cognex's own parser is authoritative for their symbology stack.
> 2. **Backup / corroboration** — existing native `DataFormatCheckResult` from push XML
>    (via `DmstResultParser`) remains in place. If the HTML report arrives before the
>    record is written, HTML data wins; otherwise the push XML data is used.
>
> **Implementation note**: The `DmstHtmlScraper` currently merges HTML fields into a
> `VerificationRecord` via `TryMergeAsync()`. The DFC table scrape would populate
> `record.DataFormatCheck.Rows` (or a new `HtmlDfcRows` parallel field) at merge time,
> before `ExcelWriter.AppendRecord()` is called. `ParseDetailRowWriter` then reads
> whichever is present — no change needed to the write layer.
>
> **HTML DFC table structure** (from `pcm_report.html` inspection):
> The `pcm_report.html` / `codes.xml` General Characteristics block contains the full
> DFC result. The exact HTML table structure needs one live HTML capture with a GS1 symbol
> to confirm the `<td>` layout before implementing the scraper extension.
> *(Not yet implemented — unblocked, do on request.)*
>
> ---
> **DESIGN DECISION LOGGED (2026-06-07) — Parse-detail HRI format + source hierarchy**
>
> **Confirmed cell format** (Symbology ID omitted — it is already prominent in the parent row):
> ```
> GS1 | HEADER: <F1> | GTIN (01) 00355513710213 | SERIAL (21) 100003289347 | [GS] | USE BY (17) 280331 | BATCH/LOT (10) 1197170
> ```
> Rules:
> - Lead with abbreviated standard (`GS1` / `MIL-130` / `ISO-15434`)
> - `HEADER: <F1>` — include the GS1 Header row with its `<F1>` data value (not suppressed)
> - GS1 HRI style: `TITLE (AI) value` for each AI element
> - `| [GS] |` appears as its own pipe-delimited segment wherever a GS separator (`0x1D`) occurs in the encoded data
> - Symbology ID (`]d2` etc.) is NOT included — it is shown in the parent verification row
>
> **Source hierarchy for parse-detail content:**
> 1. **PRIMARY — `DmstHtmlScraper` HTML report** (`pcm_report.html`): The DMST TruCheck
>    report is the authoritative source. It contains not only the DFC AI table but also
>    TruCheck-specific warnings, quality notes, and other data that are ONLY available
>    in the HTML report and cannot be obtained from push XML or `DecodedData` alone.
>    The scraper extension to parse the DFC/AI section of the HTML is the correct path.
> 2. **FALLBACK — hand-parse of `VerificationRecord.DecodedData`**: Pure C#, no native
>    DLL, builds on all platforms. Split on `0x1D`, extract AI codes, look up titles from
>    the bundled `gs1-syntax-dictionary.txt`. Use only when HTML report data is absent.
>    The GS1 syntax engine DLL is NOT to be referenced from ExcelEngine (breaks Linux build).
>
> **Not yet resolved — deferred:**
> - Operator UI choices (how to surface warnings, what to show/hide)
> - Exact HTML `<td>` layout for the DFC/AI table in `pcm_report.html` — requires one
>   live HTML capture from a GS1 symbol scan to confirm before implementing the scraper extension
>
> **Level 2 (per-scan-line) outline level**: NOT yet set; `PerScanTableWriter` still writes
> rows without `OutlineLevel`. Wire `SetRowOutlineLevel(row, 2)` + `SetRowHidden(row, true)`
> inside `PerScanTableWriter.WriteScans()` loop — one-liner, unblocked, do on request.
>
> ---
> **ENHANCEMENT NOTE (2026-06-08) — AppData Valid. column (Col A)**
>
> Formal name: **"AppData Valid."** (two display lines in header cell).
>
> **Parent row (Level 0) behavior:**
> - If application data validation is NOT configured/selected in the session: cell is white,
>   content = `"Not Set"`
> - If application data validation IS active: cell follows the same color/label rule as the
>   child row (see below) — reflects the overall parse result for that scan
>
> **Level-1 parse-detail child row (Col A) behavior:**
> - Background = **strong medium green** + content = `"Pass"` when parse result = Pass
> - Background = **medium yellow** + content = `"Warning"` when parse result = Warning
>   (e.g. non-fatal AI anomaly, deprecated AI, etc.)
> - Background = **strong red** + content = `"Fail"` when parse result = Fail
>   (e.g. data format check failure)
>
> Source for pass/fail/warning: `DataFormatCheckResult` / HTML report parse result.
> Warning state is not currently exposed in push XML — may require HTML report scraper
> to surface. Do not implement warning without confirming the source.
>
> Currently Col A of the Level-1 row holds the `"↳"` sentinel. This enhancement
> replaces that sentinel with the colored Pass/Warning/Fail indicator.
> `ParseDetailRowWriter.WriteParseDetailRow()` is the method to update.
> **Note for next build — do not implement until instructed.**
>
> ---
> **ADJUSTMENT NOTES (2026-06-08)**
>
> **Level 1 fill color**: Change from pale amber (#FFF2CC) to **very pale blue**.
> Rationale: amber risks visual confusion with the yellow "Warning" state of the
> AppData Valid. indicator. Pale blue is neutral and unambiguous.
> Target: `ParseDetailRowWriter.AmberFill` constant — rename and recolor.
> **Do not implement until instructed.**
>
> **Level 1 auto-collapse — Phase II setting** (deferred):
> Current default: auto-collapse ON, 20-second timer.
> Phase II will expose a user-configurable setting with four options:
>   - Off (row stays open indefinitely)
>   - On (collapse immediately after a brief display)
>   - On with timer (current default — collapse after N seconds)
>   - Timer interval (user-set N, in seconds)
> This will be a session/job template setting surfaced in the Settings or Session UI.
> **Do not implement until instructed.**
>
> **Level 2 data capture — traceability emphasis** (adjustment note):
> Level 2 rows should capture ALL available data in the corresponding scan-line pass,
> with particular emphasis on traceability fields (decoded data, lot, serial, expiry
> where present). Do not limit to grade columns only — surface whatever the firmware
> provides per scan pass. Exact field list TBD when Level 2 is implemented.
> **Do not implement until instructed.**

### Concept origin
OptiDoc (PIPS/Viktor) was the only tool to put per-scan-line data into the
spreadsheet as hidden child rows beneath the summary parent. Webscan/WTC never
did this. VTCCP will implement and extend the concept to cover both scan-line
children AND parse-detail children, using Excel's native outline/grouping to
create a three-level hierarchy.

### Three-level row hierarchy

```
Level 0  ── Parent row (always visible)
             Summary scan result: grade columns, decoded data, overall grade.
             Identical to what VTCCP writes today.

Level 1  ── Parse-detail child row (OPEN by default; auto-collapses via COM timer)
             One row per scan. Present for ALL parsed symbols (GS1-DM, GS1-QR,
             GS1 Code 128, etc.) whenever a DFC parse result exists.
             "Free-form" — indented label/value pairs, color-coded.
             Does NOT need to align to parent column headings.
             Written as Hidden=false (visible). COM timer collapses it after
             user-configurable interval. Operator can expand/collapse freely
             at any time via Excel outline buttons.

Level 2  ── Per-scan-line child rows (further collapsed inside Level 1 group)
             One row per scan-line pass. Present for 1D symbols only.
             Data IS aligned to parent column headings (same grade columns).
             Count is variable — read from firmware data, never hardcoded.
             (Default in DMST is 10, but configurable.)
```

### Visual treatment

| Level | Outline level | Hidden at write? | Fill color | Indentation |
|---|---|---|---|---|
| 0 — Parent | 0 | No | White / alternating | None |
| 1 — Parse detail | 1 | **No (open)** → COM timer collapses | Pale amber | One indent (col A label) |
| 2 — Scan-line passes | 2 | Yes (further collapsed) | Pale blue | None — align to parent cols |

### Level 1 — Parse-detail row format

Column A contains the sentinel label `"↳ Ignore Col Labels"` (or similar) in
italic/gray, signaling to any reader that this row's column semantics differ
from parent rows. Color coding (pale amber fill) is the second visual signal.

**Content**: the GS1 syntax engine `HRI` property with `IncludeDataTitlesInHRI=true`
returns one string per AI, e.g.:

```
GTIN (01) 00355513710213
BATCH/LOT (10) A1234
USE BY or EXPIRY (17) 261231
SERIAL (21) SN-98765
```

These are joined with `" | "` and written as a single cell value starting at
column B (or a dedicated parse column), in encoded data order:

```
GTIN (01) 00355513710213 | BATCH/LOT (10) A1234 | USE BY or EXPIRY (17) 261231 | SERIAL (21) SN-98765
```

The GS1 syntax engine dictionary (2026-01-27 release, bundled at
`vtccp/lib/gs1-syntax-engine/`) contains all current GS1 AI short titles —
no separate AI table needed.

### Level 1 — COM auto-collapse timer

Because VTCCP holds a live COM connection to the open Excel instance during a
session, the "timer" lives entirely in C# — no VBA macro in the file needed:

1. Row written with `Hidden = false` (open/visible at scan time)
2. VTCCP starts a `System.Timers.Timer` (or `Task.Delay`) for the configured interval
3. On elapsed: COM call → `ws.Rows[rowIndex].Hidden = true` → row collapses
4. Operator can re-expand at any time via Excel's `+` outline button
5. Interval is a VTCCP session setting (default TBD — suggest 10–30 seconds)

If COM is not attached (EPPlus-only session, e.g. file not open in Excel),
the row stays open indefinitely — no auto-collapse. This is acceptable because
EPPlus sessions imply no live operator view.

**Why COM is the right tool here**: The timer behavior is an interactive UX
feature — it only makes sense when a human is watching the sheet in real time.
That is exactly the scenario where COM is already attached.

### Level 2 — Scan-line row format

Columns exactly match parent grade columns. Parent holds the summary
(average or minimum, per ISO 15416). Each Level-2 child holds the per-pass
individual measurement. Identity columns (Date, Time, Operator, etc.) are blank
on child rows — they inherit context from the parent above.

### Applicability by symbol type

| Symbol type | Level 1 (parse) | Level 2 (scan lines) |
|---|---|---|
| GS1 DM / GS1 QR | Yes (if DFC result present) | No (2D per-module data on separate sheets) |
| GS1 Code 128 / EAN / UPC | Yes | Yes (N scan-line passes) |
| Plain DM / QR (no GS1 parse) | No | No |
| Plain 1D (no GS1 parse) | No | Yes |

### Excel mechanics

- `row.OutlineLevel = 1` + `row.Hidden = true` → Level 1 (EPPlus / NPOI both support)
- `row.OutlineLevel = 2` + `row.Hidden = true` → Level 2
- `sheet.SetRowGroupCollapsed(rowIndex, true)` → collapsed by default
- Operator uses Excel's native `+`/`−` outline buttons to expand/collapse
- No macro required — all set at file-creation time

### Row pointer interaction (cursor-aware insert)

The row pointer (Col A, hidden row 2) must advance past ALL child rows of the
previous parent — not just `+1`. On each new scan, VTCCP must:
1. Read pointer (first empty parent slot)
2. Write parent row at pointer
3. Write Level 1 child row at pointer+1 (if parse data present)
4. Write Level 2 children at pointer+2 … pointer+1+N (if scan lines present)
5. Advance pointer to pointer + 1 + (1 if parse) + N_scan_lines

### Open questions

- Where does per-scan-line grade data come from for 1D? Push XML sub-table,
  or DmstHtmlScraper? Check `PerScanTableWriter.cs` — it writes a separate sheet
  today; determine if that data should also feed Level-2 child rows on main sheet.
- Level 1 column layout: dedicated far-right cols vs. repurposed left cols with
  color signal — decide at implementation time.
- Should Level-2 rows carry the Formal Grade / Overall Letter columns populated,
  or only the raw per-parameter values?

---

## Key code locations (reference only — do not edit without instruction)

| Topic | File | Lines |
|---|---|---|
| TRIGGER.TYPE read on connect | `DeviceInterface/DeviceSession.cs` | 133–138 |
| SET TRIGGER.TYPE (commented out) | `DeviceInterface/DeviceSession.cs` | 140–143 |
| Restore on disconnect (commented out) | `DeviceInterface/DeviceSession.cs` | 172–188 |
| Restore on reboot (commented out) | `DeviceInterface/DeviceSession.cs` | 210–220 |
| Push trigger (TRIGGER ON via port 23) | `VtccpApp/ViewModels/SessionViewModel.cs` | 453–570 |
| HTTP subscriber start | `DeviceInterface/DeviceSession.cs` | ~190 |
| GS1 formatter (parked) | `DeviceInterface/Dmst/BarcodeDataFormatter.cs` | all |
| Manual vs Push trigger dispatch | `VtccpApp/ViewModels/SessionViewModel.cs` | 421–446 |
