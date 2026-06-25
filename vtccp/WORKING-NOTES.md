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

**Diagnostic tool**: `vtccp/tools/Fix-LiveimgMode.ps1` — reads all 7 known-good parameters,
compares each to the expected value, corrects LIVEIMG.MODE in place if it is the only
deviation, and recommends CONFIG.DEFAULT for any other deviation pattern.

### Recurrence — 2026-06-20 **RESOLVED**

Blank-image symptom recurred.  Root trigger: a QR code scan with DMST set to
Full + Bitmap caused a device lockup (buffer overrun on ~14 MB full-frame BMP preparation).
Power cycle cleared the lockup but left one or more NVRAM parameters corrupted again.

**Fix applied**: CONFIG restore from last known-good config save.  Device confirmed
working after restore.  Same CONFIG.DEFAULT recovery pattern as 2026-05-31.

**Pattern confirmed**: Full+Bitmap on QR is a reliable way to corrupt NVRAM.
Do not use Full+Bitmap mode on QR scans.

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
> **LIVE VIEW WINDOW — SCOPE NOTES (2026-06-08)**
>
> Separate WPF window for Phase I. Final UI to closely resemble WTC (Webscan TruCheck) model.
>
> **Image panel**: occupies 1/3 to 1/2 of the CP window (NOT most of screen).
>
> **Live feed FPS**: 3 fps default (closer to DMST behavior).
>
> **Verify behavior (CORRECTED)**: After a Verify scan, live feed does NOT resume automatically.
> Image stays frozen on the result. Operator must click Go Live again to resume feed.
> State machine: Idle → Live (Go Live) → Frozen/Result (Verify) → Live (Go Live again).
>
> **Button layout confirmed**:
> - Go Live: connects if needed, starts IMAGE.SEND polling at 3 fps, label → "Verify"
> - Verify: sends software trigger, displays result image, stops — stays frozen until Go Live
> - Cancel Live Feed: stops polling, freezes last image, label → "Go Live"
>
> **IMAGE.SEND availability in DeviceSession**: TBD — user to confirm; check before implementing.
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

## v1.36 push script — scan #17, 2026-06-20T08:51:33

**PushScriptDiag**: `v1.36 q=r.trucheck m=found`

**Note**: v1.36 is running because the CONFIG restore reverted the device to the .dmb
save state (which had v1.36). v1.37 was installed after that backup — lost on restore.
This is the "push script auto-deploy on connect" problem in practice.

**Symbol**: GS1 DM 22×22, GS1 ]d2, overall grade F (DecodeGrade=F, proprietary Decode metric)

### JpegImageBase64 format — CONFIRMED JPEG regardless of DMST PNG setting

`JpegImageBase64` first bytes: `FFD8FFE0` — JPEG SOI+APP0 marker.
DMST Image Panel toolbar shows "PNG" selected — has NO effect on push XML image.
The push XML image path (`JpegImageBase64`) is firmware-generated JPEG, always.
The DMST format dropdown controls only the separately-saved decoded file (DMST logging
path). These are two independent image delivery channels:
- **Push XML path** (`JpegImageBase64`): firmware-controlled, always JPEG, no knob.
- **DMST logging path** ("native save decoded file"): format dropdown applies here.

Excel always receives JPEG via push XML. Do not attempt PNG conversion on the CP side.

### EncodedCharacters: HTML=49, push XML (eaLen fallback)=31 — MISMATCH CONFIRMED

HTML report `Encoded characters = 49`; push XML emits 31 (eaLen fallback).
For this 22×22 GS1 symbol the firmware says 49 encoded characters.
eaLen=31 is the encodationAnalysisArray.length, which counts encoding segments, not characters.
The prior observation (scan #16 eaLen=31 ≈ correct) was a coincidence for that data.
Dead-path status confirmed again. `DmstHtmlScraper` is the only correct source.

### DataCodewords / Error Correction Budget — from HTML report

| Field | Push XML | HTML report |
|---|---|---|
| DataCodewords | (empty) | 30 |
| ErrorCorrectionBudget | (empty) | 20 |
| TotalCodewords | 50 | 50 ✓ |
| ImagePolarity | (empty) | White on black |
| EncodedCharacters | 31 (wrong) | 49 |

### DebugBarcodeAssignment — NEW sub-key: `stats=[obj]`

v1.36 output: `result=-1;stats=[obj];`
Compared to scan #16 (v1.37): `result=-1` only.
Either `stats` is conditionally present or was missed in the prior ekv pass.
Queue a **DebugBarcodeAssignmentStats** probe in v1.38: enumerate `stats` sub-keys.
Note: `result=-1` on Grade F scans confirmed on two independent symbols — but Grade F
does NOT mean no decode; symbol was decoded in both cases. Interpretation TBD.

### ISO 15415:2024 — new edition string

GradingStandard: `ISO 15415:2024` — first observation of the 2024 edition.
Prior test data used 2011/2016. Parser must echo whatever the device reports; confirmed.

### Custom Note: "Test of 1 4th PNG Image"

Operator-set probe note confirming the user was testing PNG save functionality.
CustomNote field wired and populating correctly.

### DMST "native save decoded file" — PNG confirmed working, resolution confirmed

Saved file: `2026-06-20_09-00-37-256.png`
- **Dimensions: 1224 × 1024 px** (8-bit grayscale, 1,254,709 bytes / ~1.2 MB)
- This is **Quarter resolution** in DMST terms = IMAGE.SIZE=1 = ¼ pixel count = ½ each linear dimension of the 2448×2048 full sensor.
- The saved file is the **full IMAGE.SEND frame** at the configured resolution — NOT a barcode ROI crop. Full scene context is preserved.
- PNG format confirmed working. DMST format dropdown DOES control this save path.

**⚠ TRIGGER DEPENDENCY — CONFIRMED 2026-06-20**: DMST native save ONLY fires when the scan
is triggered FROM WITHIN DMST. A CP-triggered scan (DMCC software trigger) does NOT cause
DMST to save a decoded image file. DMST must be both running AND must have initiated the
trigger for the native save path to activate. This eliminates DMST native save as a
viable L0 archive path for CP-triggered workflows. IMAGE.SEND (DMCC, CP-initiated) is
the only image retrieval path that works independently of DMST trigger ownership.

**Resolution map (DM475V, 2448×2048 sensor):**

| DMST label | IMAGE.SIZE | Saved dimensions | Pixels | Approx size (PNG) |
|---|---|---|---|---|
| Full | 0 | 2448 × 2048 | 5.0 MP | ~5 MB |
| Quarter | 1 (known-good) | 1224 × 1024 | 1.25 MP | ~1.2 MB ✓ confirmed |
| 1/16 | 2 | 612 × 512 | 0.31 MP | ~0.3 MB |
| 1/64 | 3 | 306 × 256 | 0.08 MP | ~0.08 MB |

**For DPM multi-lab grading study**: switch DMST to **Full + PNG** to get lossless 2448×2048 images. The Quarter PNG is usable for most purposes but lacks the pixel density needed for manufacturer-standard DPM analysis at 1:1.

---

## v1.37 push script — DEVICE CONFIRMED (scan #16, 2026-06-20)

**PushScriptDiag**: `v1.37 q=r.trucheck m=found`

**Symbol**: GS1 DM 22×22, overall grade F (DecodeGrade=F)

### Probe: DebugBarcodeAssignment — ANSWERED

`r.barcodeAssignment` is an object with a single key: **`result=-1`**.

`result=-1` meaning is **not yet established**. Both observations so far were Grade F
scans, but Grade F means one or more grading parameters failed threshold — the symbol
was decoded successfully in both cases. `result=-1` is NOT "decode failed."
Two candidate interpretations:
- "No barcode assignment rule configured on this device" — would return -1 on every
  scan regardless of grade.
- Tracks something about grade level or grading outcome — would change on a Grade A scan.

A Grade A scan of the same symbol (same device, same firmware) will distinguish these:
- If result is still -1 → device-config sentinel, grade-independent.
- If result changes → tracks grade or grading outcome.
No other sub-keys present on these Grade F scans. Probe is partially answered;
passing-scan observation pending (incoming — user running now).

### FormalGrade `0/F` — confirmed correct device behavior

v1.37 constructs `FormalGrade` as `op("gradeValue") + "/" + op("gradeLetter")` (script
line: `elem("FormalGrade",(op("gradeLetter"))?(op("gradeValue")+"/"+op("gradeLetter")):"")` ).
For a fail scan: gradeValue=0, gradeLetter=F → `0/F`.  This is correct and expected.
The full ISO formal notation (`1.0/16/660/45Q` etc.) is only emitted on passing scans
where the aperture/wavelength/lighting fields are all populated.  No CP code change needed.

### EncodedCharacters `31` — eaLen fallback, not a firmware fix

v1.37 tries `gp("encodedCharacters")` first; falls back to `s(_el)` (encodationAnalysisArray.length)
if that is absent.  The 22×22 symbol encoded 31 characters; the eaLen fallback happened to
equal the true encoded character count for this symbol.  This is NOT a resolution of the
dead-path finding from v1.32/v1.33 — `q.general.encodedCharacters` is still absent on
fw 6.1.16_sr4.  The fallback remains unreliable for other symbol sizes (confirmed wrong on
16×36: eaLen=33 vs DMST=38).  Field status: unresolvable from push XML; populate via
`DmstHtmlScraper` (General Characteristics block).

---

## CP file-rename architecture — PLANNED 2026-06-20

**Context**: DMST native save produces timestamp-rooted filenames (e.g.
`2026-06-20_09-00-37-256.png` / `2026-06-20_09-00-37-256.html`). DMST's
built-in naming convention control is limited and cannot be set programmatically
via DMCC. The save path is also DMST-configured only (no DMCC path command).
Both path and naming convention are sticky (persist across DMST restarts).

**CP rename approach** (agreed, not yet scheduled for build):
1. CP knows the DMST save directory (configured once at CP setup)
2. CP knows the exact scan timestamp from push XML `<DateTime>`
3. After each DMST-triggered scan CP watches for new files whose timestamp root
   matches → immediately renames both the `.html` and decoded image (PNG/JPEG)
4. Pure filesystem operation — no DMCC, no DMST API, no UI automation

**CP naming token set** (proposed — far exceeds DMST's built-in options):
`{Date}`, `{Time}`, `{DeviceName}`, `{Operator}`, `{JobName}`, `{ScanSeq}`,
`{OverallGrade}`, `{Symbology}`, `{MatrixSize}`, `{GTIN}`, `{Lot}`, `{Serial}`,
`{Expiry}`, `{NominalXDim}`, `{GradingStandard}` — any field in VerificationRecord.

**HTML file integrity rule**: CP must NOT modify the HTML content. The rename is
sufficient to make the filename meaningful. Audit identity lives in CP's session
sidecar JSON (stores renamed filepath per scan record). The renamed file IS the
official record; its name is its identity.

**Blocked by**: DMST must be open AND must have triggered the scan (confirmed
trigger dependency 2026-06-20). Rename path only applies to DMST-triggered scans.

---

## Re-grade grade differences — JPEG artifact hypothesis 2026-06-20

**Observation**: When a previously-graded symbol image is re-uploaded via
IMAGE.LOAD and re-graded, small numeric grade differences appear vs the original
live-scan grade (e.g. decimal-place differences in SC%, Modulation, etc.).

**Hypothesis (user, supported)**: The image being re-submitted is the L1 JPEG
(`JpegImageBase64` from push XML) — a lossy compressed crop of the original
sensor image. JPEG DCT compression introduces:
- **Ringing artifacts** at high-contrast module edges → alters SC% and Modulation
  measurement values
- **Blocking artifacts** at 8×8 DCT tile boundaries → particularly damaging for
  dense Data Matrix where tile boundaries cross module patterns
- **DC coefficient drift** in uniform regions → slightly shifts Rmax/Rmin,
  compressing the measured contrast range

Differences of 0.1–0.2 numeric grade units at borderline grades are expected and
consistent with this mechanism. "Rounding error" is plausible only for differences
< 0.05 at the second decimal place.

**Shared with inventor/chief scientist**: offered as JPEG artifact explanation;
inventor was surprised but did not engage with the explanation. User's read: the
inventor may have assumed IMAGE.LOAD receives original sensor data.

**Proposed confirmation test**: DMST-trigger a scan to capture L0 PNG (lossless,
full frame). Crop to L1Prime (lossless). Re-submit L1Prime via IMAGE.LOAD. If
re-grade matches original to more decimal places than L1 JPEG re-grade, JPEG
artifact hypothesis is confirmed. Has implications for any manufacturer receiving
CP-provided images for independent grading.

**Status**: Hypothesis, not yet confirmed on bench.

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

---

## CP Installation File Path Architecture — DECIDED 2026-06-20

### Two distinct categories — must not conflate

**Category 1: CP Application Data** (config, templates, settings)
Managed by CP internally; not user-configurable; survives app updates.

| Item | Path |
|---|---|
| CP settings / device profiles | `%APPDATA%\VCCS\CommandPilot\Settings\` |
| Job templates | `%APPDATA%\VCCS\CommandPilot\JobTemplates\` |
| File name format templates | `%APPDATA%\VCCS\CommandPilot\FileNameTemplates\` |
| Session history index | `%APPDATA%\VCCS\CommandPilot\Sessions\` |
| Factory defaults (read-only, shipped) | `%PROGRAMFILES%\VCCS\CommandPilot\Templates\` |

**Category 2: Verification Output Files** (reports, images, Excel logs)
Defaults shown below; operator-configurable via Data Logging Paths screen.

| Item | Default path |
|---|---|
| HTML/PDF reports | `%USERPROFILE%\Documents\VCCS Command Pilot\Reports\` |
| Decoded images (JPEG) | `%USERPROFILE%\Documents\VCCS Command Pilot\Images\Decoded\` |
| No-read images | `%USERPROFILE%\Documents\VCCS Command Pilot\Images\NoRead\` |
| Excel / CSV log | `%USERPROFILE%\Documents\VCCS Command Pilot\ExcelLog\` |

### Installer behavior
- Installer: binaries + factory templates to `%PROGRAMFILES%` ONLY — no writes to Documents or AppData during install.
- First run: CP creates `%APPDATA%\VCCS\CommandPilot\`, copies factory templates, creates default Documents output tree.
- First-run notice shown: "Verification reports and images will be saved to Documents\VCCS Command Pilot\. Change in Settings → Data Logging."
- No setup wizard needed — keeps installer silent-install-friendly for enterprise IT.

### Enterprise / network-share (V2)
- QA managers at multi-station sites point output paths to UNC share (`\\QA-SERVER\VerifLogs\`).
- Machine-wide path lock: `%PROGRAMDATA%\VCCS\CommandPilot\policy.json` — prevents operators from redirecting output. V2 scope, architecture noted.

### DMST-managed paths (CP cannot change)
- DMST PNG + HTML saves: `%USERPROFILE%\Documents\{DeviceName}\CodeQuality\{timestamp}.*`
- CP filesystem watcher monitors this path; does not write to it.

---

## CP File Path Architecture — REVISED 2026-06-21 (ProgramData decision)

**Decision**: Primary configuration location is `%ProgramData%\VCCS\CommandPilot\` — NOT `%AppData%`.

### Rationale
Verification lab workstations are shared across multiple operators. `%AppData%` is per-user-isolated — operator A's job templates would be invisible to operator B. `%ProgramData%` is machine-wide and shared across all Windows user accounts on that PC, which is correct for a shared instrument. This is NOT old school — it is the Microsoft-recommended location for machine-wide application data (Vista+). Webscan's use of ProgramData was architecturally correct for the same reason.

### Revised path layout

| Item | Path | Notes |
|---|---|---|
| Job templates | `%ProgramData%\VCCS\CommandPilot\JobTemplates\` | Shared across all operators on this machine |
| Device profiles | `%ProgramData%\VCCS\CommandPilot\Devices\` | Machine-specific (GigE IP, device config) |
| File name format templates | `%ProgramData%\VCCS\CommandPilot\FileNameTemplates\` | Global, shared |
| Output path config | `%ProgramData%\VCCS\CommandPilot\Settings\Paths.json` | Machine-level; points at server share in enterprise |
| Factory defaults (read-only) | `%ProgramFiles%\VCCS\CommandPilot\Templates\` | Ships with installer; never written by CP |
| Per-user preferences | `%AppData%\VCCS\CommandPilot\UserPrefs\` | Window layout, recent files only |
| Output files | Configurable (default local; enterprise = UNC share) | Path stored in Paths.json above |

### Enterprise output path
In an enterprise environment, `Paths.json` output paths point at a UNC network share (e.g. `\\QA-SERVER\VerifLogs\`). CP does NOT restrict the path — operator can point anywhere they have write access. Access CONTROL (who can change the path) is handled by the password tier system (see below), not by the app enforcing a fixed location.

### Multi-level password protection — gating for path settings
Based on Webscan screenshots (user-provided) and prior discussion, CP should adopt a similar multi-tier access scheme:

| Tier | Role | What they can do |
|---|---|---|
| 0 — Operator | Day-to-day scanning | Scan, view results, select pre-loaded job template |
| 1 — Supervisor | Job management | Run jobs, load/select templates, view logs |
| 2 — QA Manager | Configuration | Edit templates, configure output paths, device profiles |
| 3 — Admin | Full control | Password management, factory reset, installer-level config |

**Data Logging Paths screen** → Level 2 minimum.
**Excel Column Options screen** → Level 1 or 2 (TBD — whether operators adjust log columns).
**Device configuration screen** → Level 2.
**Access control / password setup** → Level 3.

Full design of the access control UI is pending — screenshots of Webscan's scheme referenced but detailed implementation not yet scoped. DO NOT IMPLEMENT until scoped.

---

## TD: Webscan TruCheck Replica Excel template — 2026-06-21

### Decision log
- A "Webscan TruCheck Replica" preset will be added at the **bottom** of the Excel Column Options preset list, below all VTCCP-native presets.
- Layout is **frozen and non-configurable** — column toggles are disabled for this preset; it is rendered exactly as TruCheck exports it, column for column. UI note: *"Layout is fixed to match Webscan TruCheck output. Column selection is not available for this preset."*
- TruCheck product line is **EOL** — no future firmware changes expected, so the column layout will not drift. No versioning tag required.
- Webscan TruCheck **does support QR Code** — the replica preset must cover DM, QR, and linear (1D) symbologies. Edge cases: 4-quad DM (2-row symbol layout?), larger-content QR (more codewords, higher version).

### Authoritative source
Live TruCheck Excel exports are the ground truth for column layout — **not documentation**.
User to provide live exports for:
- [ ] Key linear (e.g. Code 128, EAN-13)
- [ ] Standard DataMatrix (e.g. 16×36)
- [ ] 4-quad DataMatrix
- [ ] Standard QR Code (e.g. Version 3, 29×29)
- [ ] Larger-content QR (higher version, more codewords)

Column mapping work begins only after these exports are received. DO NOT IMPLEMENT column layout until source data arrives.

### Migration import — TD (low priority, easy to offer)
Offer a future one-time import operation: **"Import data from Webscan TruCheck Excel export into a CP log file."** Mechanism: CP reads a user-selected TruCheck `.xlsx`, maps WTC columns to CP schema, appends (or creates) a CP-format output file. Rows that map cleanly transfer verbatim; fields with no WTC equivalent emit empty.

Surface this option:
- In the "Open / Import" menu or File menu
- Possibly as a first-run prompt if CP detects an existing TruCheck file in the default output path

Scope and build only after the WTC column layout is fully mapped from live export data.

---

## TD: Excel column visibility architecture — LOCKED 2026-06-21

### Decision 1 — Column presence model
**All CP columns are always written to every Excel file.** Presets (Standard, Minimal, QA Audit, WTC Replica, etc.) control column *visibility* (hidden/shown) only — not column presence. A column hidden by the active preset still receives its data value. This means:
- Data is never silently discarded
- Switching presets retroactively is possible — just unhide the relevant group
- No schema migration when a preset changes

### Decision 2 — WTC Replica trailing columns
In WTC Replica mode, WTC-layout columns appear first in their native order (visible). All CP columns that have no WTC equivalent are appended as **hidden trailing columns** after the last WTC column. These trailing columns use the **same header names as the main CP template** — identical strings, not aliases. Rationale: when an operator later migrates a WTC Replica file to a CP template file, the column headers match exactly and no mapping table is required.

### Decision 3 — Hidden column warning behavior
When a hidden column receives a **non-null value** for the first time in a session, CP surfaces a **category-level toast/warning** in the application status area (not in Excel). Warning rules:
- Fires once per category group per session (e.g. "QR Pattern Grades", "Modulation Values", "HTML Scraper fields") — not once per column or per scan
- Does NOT fire if the field is null/empty (normal for symbology-specific fields on the wrong symbology)
- Action offered: *"Show [category] columns"* — CP unhides that group in the live file
- If operator dismisses: suppressed for the remainder of the session
- **Global override**: a future config UI (Level 2 or Level 3) can offer a persistent setting — "Always warn / Never warn / Warn once per session (default)". Until that config UI exists, per-session suppression is the only option.

---

## TD: CP Verifier Config Preset — Post-Calibration Conformance Challenge (LOGGED 2026-06-21)

**Concept**: CP offers a one-click verifier configuration preset that sets the device into a standardised state for running a post-calibration conformance challenge on a 2D symbol (e.g. NIST/ISO reference target).

### Known required settings

| Setting | Value | Notes |
|---|---|---|
| Grading standard | ISO 15415 / ISO 15416 | Combined 2D+1D standard selection |
| Lighting | 45° | Locked/implied by ISO 15415/15416 standard selection — device restricts to 45Q automatically |
| Aperture | 50 / 80 | Dual-aperture (TBD which aperture per symbol type or operator choice) |
| Batch Number (user field) | TBD | DeviceName already captured in Block A — Batch purpose for conformance challenge TBD |
| Custom Note | "Post-calibration Conformance Challenge" | Fixed string — identifies the record type in log |

### TBD items
- Additional DMCC settings locked by the standard selection (confirm from device)
- Whether aperture is operator-selectable (50 vs 80) or both required in sequence
- Any required symbology enables/disables for conformance challenge
- Whether a specific reference target (part number / lot) should be logged
- Whether CP should enforce a minimum number of repeat scans before accepting the session as a valid conformance record

### Architecture notes
- This preset writes DMCC settings to the device (not just CP UI state) — it is a full device configuration operation, not just a log-format preset
- Requires DMCC write capability in CP (already planned via raw TCP port 23)
- After the challenge session ends, CP should offer to restore the previous device config (or leave it to the operator)
- Blocked by: Device Config feature (DMST TC window screenshots required — see LOGGED FEATURE PLAN 2026-05-25)


---

## FTP-IMAGE — Full-Frame Archival Channel (2026-06-24)

**Context**: Glenn Reuss (Cognex chief eng) suggested using device FTP to capture bits missing from current extraction methods — specifically FTP-PCM-REPORT for HTML delivery. Investigation expanded to all three FTP result channels.

### Assessment of all four FTP subsystems

| Channel | Content | Delta vs current CP? |
|---|---|---|
| `FTP-PCM-REPORT` | pcm_report.html — identical to PUT /pcm_report.html already received via HTTP subscriber | **None** — we already get this |
| `FTP-RESULT` | Formatted decode string (Data Formatting template output) — subset of codes.xml | **None** |
| `FTP-IMAGE` | Image per scan — **content TBD, likely full-frame at IMAGE.SIZE** | **Potentially high** — see below |
| `CONFIG-BACKUP` | .dmb config backup file — not scan data | N/A |

### FTP-IMAGE — the critical channel

**Hypothesis (Glenn-confirmed direction)**: FTP-IMAGE delivers whatever `IMAGE.FORMAT` and `IMAGE.SIZE` are set to via DMCC. This is almost certainly the full-frame acquisition path — distinct from `JpegImageBase64` in codes.xml (which is always a firmware-generated ROI crop, always JPEG, no knobs).

**Why this matters**: DMST's "Log All Decoded Images" function (Data Logging tab → Decoded Images path/prefix) and FTP-IMAGE are almost certainly the same underlying firmware mechanism — one saves to filesystem, the other pushes to FTP server. The image content is the same: full-frame at IMAGE.SIZE resolution. The DMST path settings are NOT sticky (Glenn confirmed Cognex will not fix this — internal emphasis on browser migration). FTP-IMAGE settings ARE sticky (DMCC-settable, NVRAM-persisted).

**Image format control:**

| DMCC Key | Values | Notes |
|---|---|---|
| `IMAGE.FORMAT` | `0`=JPEG, `2`=PNG | PNG support from v5.6.3+; PNG confirmed on DM475V |
| `IMAGE.SIZE` | `0`=Full (2448×2048), `1`=1/4 (1224×1024), `2`=1/16, `3`=1/64 | Same key that controls IMAGE.SEND downscale |

**⚠ UNCONFIRMED — ONE PROBE NEEDED**: Enable `FTP-IMAGE.ENABLE ON` pointing at a FileZilla instance, run one scan, check file dimensions. If 2448×2048 → confirmed full-frame. If ~200–600px → duplicates JpegImageBase64 (low value). Result will definitively answer this.

**What FTP-IMAGE solves if confirmed full-frame:**
1. L3 full-frame archival without requiring Cognex SDK `GetResultImage()` (eliminates D4 SDK dependency)
2. Sticky path/prefix via DMCC — replaces non-sticky DMST Data Logging settings entirely
3. DMST-independent — works for CP-triggered scans (DMST filesystem save does NOT fire for CP-triggered scans)
4. Format choice: PNG for lossless archival; JPEG for space-efficient daily use — operator-configurable via DMCC

**Replacement mapping for DMST Data Logging settings (screenshot 2026-06-24):**

| DMST "Data Logging" field | FTP-IMAGE replacement |
|---|---|
| Decoded Images → Path | `FTP-IMAGE.IP-ADDRESS` + `FTP-IMAGE.SERVER-PATH` |
| Decoded Images → Filename Prefix | `FTP-IMAGE.FILE-NAME` or `FTP-IMAGE.FILE-NAME-USE-SCRIPT` |
| Image resolution (not shown in UI, controlled by DMST connection settings) | `IMAGE.SIZE` via DMCC |
| Image format (PNG / BMP) | `IMAGE.FORMAT` via DMCC |
| Include Overlay Graphics (SVG) | **No FTP equivalent** — SVG overlay is DMST-only rendering |

### FTP-IMAGE DMCC key inventory

All keys: `FTP-IMAGE.ENABLE`, `FTP-IMAGE.IP-ADDRESS`, `FTP-IMAGE.PORT`, `FTP-IMAGE.USER-NAME`, `FTP-IMAGE.PASSWORD`, `FTP-IMAGE.SERVER-TYPE` (FTP or SFTP — SFTP confirmed on DM380/390/580/590/8700, DM475V support TBD), `FTP-IMAGE.FILE-NAME`, `FTP-IMAGE.CUSTOM-FILE-NAME`, `FTP-IMAGE.FILE-NAME-USE-SCRIPT`, `FTP-IMAGE.SERVER-PATH`, `FTP-IMAGE.SERVER-PATH-SCRIPT`, `FTP-IMAGE.SERVER-PATH-SCRIPT-ERROR` (GET-only error string), `FTP-IMAGE.MAX-APPEND`, `FTP-IMAGE.MAX-APPEND-START-VALUE`, `FTP-IMAGE.IDLE-LIMIT`, `FTP-IMAGE.IDLE-TIME`, `FTP-IMAGE.SERVER-FINGERPRINT` (SFTP host key).

Platforms: all Ethernet readers. DM8000 wireless with Ethernet base station also supported.

---

## TRIGGER.TYPE=4 "Self (internal)" — Live View Mechanism Identified (2026-06-24)

**Source**: DMCC reference `TRIGGER.TYPE` detail page confirmed 2026-06-24.

### All six TRIGGER.TYPE modes

| Value | Name | Class | Notes |
|---|---|---|---|
| 0 | Single | External | **CP idle/verify state — confirmed** |
| 1 | Presentation | Internal | Motion-detect auto-trigger; our earlier wrong hypothesis |
| 2 | Manual | Button | Physical trigger button |
| 3 | Burst | External | Multi-image per trigger |
| **4** | **Self** | **Internal** | **NEW FINDING — self-triggers at CAMERA.INTERVAL-US** |
| 5 | Continuous | External | Continuous external trigger |

### Self (internal) trigger — key facts

**`CAMERA.INTERVAL-US`**: interval between successive camera acquisitions in microseconds. Step size: 250 µs. Supported for Burst (3), **Self (4)**, and Continuous (5).

- 3 Hz = 333,333 µs
- 4 Hz = 250,000 µs

**Conclusion**: The 3–4 Hz LED flicker observed during DMST live view is the device running in **TRIGGER.TYPE=4 (Self)** at CAMERA.INTERVAL-US ≈ 250,000–333,333. DMST sets TRIGGER.TYPE=4 when entering live view mode and restores TRIGGER.TYPE=0 when exiting.

**MOTION-DETECTION.ACTIVE** (GET-only, v5.7.0) — covers both Presentation (1) and Self (4). Both trigger types share the motion-detection hardware path. This is how Glenn recognized "self internal" as a motion-related context.

### CP Live View architecture using TRIGGER.TYPE=4

```
Enter live view:
  SET TRIGGER.TYPE 4
  SET CAMERA.INTERVAL-US 333333   ← 3 Hz; configurable
  → device self-triggers at 3 Hz
  → each self-trigger fires a full TruCheck scan
  → IMAGE.SEND after each trigger → JPEG frames for viewfinder UI
  → HTTP subscriber receives PUT /codes.xml per frame (⚠ unconfirmed — see below)

Cancel live view:
  SET TRIGGER.TYPE 0
  → self-triggering stops immediately
  → device returns to idle single-external mode

VERIFY (trigger one canonical scan):
  SET TRIGGER.TYPE 0              ← freeze the label / stop continuous triggers
  TRIGGER ON                      ← fire one deliberate TruCheck scan
  → result arrives via HTTP subscriber (codes.xml + pcm_report.html)
```

**User-confirmed trigger state model (2026-06-24)**:
- Cancel live view = `SET TRIGGER.TYPE 0`
- VERIFY = `SET TRIGGER.TYPE 0` + `TRIGGER ON` (freeze then fire)

The VERIFY step serves operator intent: stop the continuous viewfinder to freeze the label in a known position, then fire a clean deliberate capture. This distinguishes a QA-grade verification record from a random frame from the continuous viewfinder cycle.

**⚠ CORRECTED 2026-06-24**: Live view does NOT produce TruCheck verifications. User confirmed: no decode XML generated, no result history entry, no codes.xml event. Live view is frame acquisition for the viewfinder only — camera exposure + display, not a verification cycle. The 3–4 Hz LED flicker is camera acquisition, not TruCheck.

**★ PROBE RESULT — 2026-06-24**: `||>GET TRIGGER.TYPE` → `||0` while DMST is in live view.

**TRIGGER.TYPE=4 hypothesis is WRONG. TRIGGER.TYPE=0 (Single external) during live view — confirmed.**

DMST does not change TRIGGER.TYPE when entering live view. The 3–4 Hz LED flicker is NOT from Self (internal) triggering.

**Live view mechanism is still unknown.** Most likely: DMST sends rapid `TRIGGER ON` DMCC commands via its own connection at 3–4 Hz while suppressing the result from the verification pipeline — or DMST uses an internal SDK streaming path that bypasses standard DMCC trigger state entirely. Neither path is accessible to CP via raw DMCC. The live view LED behavior and TRIGGER.TYPE are decoupled.

Note: `CAMERA.INTERVAL-US` query was sent without `||>` prefix — silently ignored by device. Moot regardless: CAMERA.INTERVAL-US only applies to TRIGGER.TYPE 3/4/5; with TRIGGER.TYPE=0 confirmed it is irrelevant to live view.

### Implication for CP session lifecycle

- On connect: GET TRIGGER.TYPE → confirm returns 0 (single external) at idle
- If GET returns 4: DMST is actively in live view mode — wait or alert operator
- Live view feature: SET TRIGGER.TYPE 4 + SET CAMERA.INTERVAL-US → read IMAGE.SEND frames
- No LIVEIMG.SEND needed (confirmed dead on all ports/modes)
- Restore to 0 on session end — same restore path as existing TRIGGER.TYPE management


---

## SDK DLL inspection — PowerShell procedure (2026-06-24)

**Purpose**: Identify the correct Cognex DataMan SDK v25.4.1 discovery class name to
replace the `EthSystemDiscoverer` stub in `NetworkDiscoverer.cs`.
`EthSystemDiscoverer` does not exist in the v25.4.1 DLL; correct class name unknown.

**Requirements**: Windows PowerShell 5.1 or PowerShell 7+ (both ship with Windows 10/11).
No third-party tools required.

### Step 1 — Open PowerShell

Press `Win+R`, type `powershell`, press Enter.

### Step 2 — Run the inspection script

Copy and paste the following block into the PowerShell window, then press Enter:

```powershell
$dll = "C:\Program Files (x86)\Cognex\DataMan\DataMan Software v25.4.1\Cognex.DataMan.SDK.PC.dll"
$asm = [System.Reflection.Assembly]::LoadFile($dll)

Write-Host "`n=== Types matching *Discover* ===" -ForegroundColor Cyan
$asm.GetTypes() | Where-Object { $_.Name -like "*Discover*" } | Select-Object -Expand FullName

Write-Host "`n=== All public types in Cognex.DataMan.SDK ===" -ForegroundColor Cyan
$asm.GetTypes() | Where-Object { $_.Namespace -eq "Cognex.DataMan.SDK" } |
    Select-Object -Expand FullName | Sort-Object
```

### Step 3 — Read the output

The first block lists any type whose name contains "Discover" — the discovery class will
appear there if it exists in this DLL. The second block lists every public type in the
`Cognex.DataMan.SDK` namespace — scan for anything related to network scanning, discovery,
or system enumeration.

### Step 4 — Report back

Paste the full output here. Once the correct class name is confirmed, the stub in
`NetworkDiscoverer.cs` will be replaced with the real implementation (one pass, ~20 lines).

**If the DLL path is different** (e.g. different version directory), adjust `$dll` on line 1.
To list all Cognex DLLs available on the machine, run first:
```powershell
Get-ChildItem "C:\Program Files (x86)\Cognex" -Recurse -Filter "*.dll" |
    Where-Object { $_.Name -like "*DataMan.SDK*" } | Select-Object FullName
```

---

## Image layer model — clarification 2026-06-24

### Correction: L2 (IMAGE.SEND) is the full camera scene, NOT the ROI frame

`IMAGE.SEND` returns the **full camera scene** at `IMAGE.SIZE` resolution. It is not
cropped to `DECODER.ROI`. The field name `RoiJpegImageBase64` in `VerificationRecord` is a
legacy misnomer — the content is the full scene. `image-capture-pipeline.md` L2 section
corrected accordingly.

### Image layer model (authoritative, 2026-06-24)

| Level | Name | Source | Content | Captured? |
|---|---|---|---|---|
| L0 | DMST Native PNG | DMST filesystem save | Full sensor frame (2448×2048), lossless | DMST-triggered scans only |
| L1 | Barcode crop | Push XML `JpegImageBase64` | Firmware-generated tight crop around symbol only (~200–600 px) | Every scan, automatically |
| L2 | Full camera scene | `IMAGE.SEND` DMCC | Full camera frame at `IMAGE.SIZE` resolution (e.g. 1224×1024 at SIZE=1) | On demand, CP-triggered |
| L2.ROI | Virtual ROI crop | L2 cropped to `DECODER.ROI` | Operator-defined scan region — wider than L1, tighter than L2 | **Virtual** — derived in software |
| L3 | Full sensor frame | SDK `GetResultImage()` | 2448×2048 full sensor | D4 scope |

**Key distinction (user-confirmed 2026-06-24)**:
- L1 (barcode crop) is the **firmware-generated bounding box** around the symbol only —
  equivalent to what DMST shows in the verification panel crop. This is a fixed firmware
  output; CP has no control over its extent.
- L2.ROI (Virtual ROI crop) is the **operator-configured scan region** (DECODER.ROI) —
  always larger than L1; includes surrounding HRI text, lot number, expiry date, etc.
  It does NOT exist as a separately-captured image anywhere in the firmware output.
  It is derived by cropping L2 (or L3) to the `r.image.RoI` coordinates.

### Virtual ROI Crop — derivation at any time

The ROI coordinates are available from two sources:
- `r.image.RoI` — embedded in every push result (confirmed: 28-key r.image inventory, scan #12)
- `DECODER.ROI` DMCC GET — queryable at any time

To construct the Virtual ROI Crop:
1. Obtain L2 full camera scene (IMAGE.SEND)
2. Scale DECODER.ROI sensor-space coords to IMAGE.SIZE pixel space (factor 0.5 at SIZE=1)
3. Crop L2 to scaled ROI rect → Virtual ROI Crop

### PNG vs BMP — planned evaluation pass

`IMAGE.SEND` currently delivers JPEG bytes. For L2 archival, the question is JPEG vs
lossless (PNG or BMP). The user has noted:

| Property | PNG | BMP |
|---|---|---|
| Lossless | ✓ | ✓ |
| Compression | ✓ (typical 30–60% smaller than BMP for grayscale scanner images) | ✗ (raw pixels) |
| Metadata support | ✓ `tEXt` / `iTXt` chunks — arbitrary key-value pairs | ✗ (no standard metadata) |
| Ecosystem | Universal | Universal |
| Decode complexity | Minor decompression overhead | Trivial |

**Key opportunity**: PNG `tEXt` metadata chunks can embed `DECODER.ROI` coordinates
(from `r.image.RoI`), scan timestamp, device name, grade letter, and any other
`VerificationRecord` field — all in the same file as the pixel data. This makes the
**Virtual ROI Crop reconstructable from the PNG alone** at any future time, with no
additional database or sidecar file required.

If full-frame archival is saved as PNG + embedded metadata, the L2.ROI is a derived
read-time operation, not a stored image — reducing per-scan storage by one image level.

**Planned pass** (not yet scheduled): evaluate PNG vs BMP/JPEG for L2 archival:
- Confirm IMAGE.FORMAT=2 (PNG) returns valid PNG bytes from IMAGE.SEND
- Measure file size vs JPEG at quality=50 and quality=85
- Prototype writing ROI coords + scan metadata to PNG tEXt chunks at scan time
- Confirm round-trip: save PNG → read PNG → crop to embedded ROI coords → verify pixel fidelity

**Not implementing until instructed** — log only.


---

## Scan #18 — QR IMAGE.LOAD Grade A, DPM device, new pre-release firmware (2026-06-24)

**Device**: DM475-DPM-866D76-VCCS-Verif-Lab (DM475V-DPM unit)
**Firmware**: Pre-release (version TBD — not yet queried)
**Push script**: v1.37 confirmed (`<PushScriptDiag>v1.37 q=r.trucheck m=found</PushScriptDiag>`)
**Symbol**: QR 29×29 (Version 3), IMAGE.LOAD, `CustomNote: 'Post-calibration Conformance Challenge'`
**Decoded**: `0e95e424-3a33-eb11-a816-001dd80187c1` (GUID, byte mode)
**OverallGrade**: A (all parameter grades A; VIBGrade=`'-'`)
**GradingStandard**: `ISO 29158:2025 (AIM-DPM)` ← NEW 2025 EDITION (previously 2024 max observed)
**OpticsSource**: `LoadedImage` (CU=-1, MRD=-1 confirmed)
**JpegImageBase64**: 19,580 chars — JPEG present and correct

### barcodeAssignment — RESOLVED 2026-06-24

`DebugBarcodeAssignment: 'result=-1;stats=[obj];'` on this Grade A scan.

Combined with scan #16 (Grade F DM, `result=-1` no stats):
- **`result=-1` is grade-independent** — confirmed device-config sentinel: "no barcode assignment rule configured on this device." NOT a decode or grade signal.
- **`stats=[obj]`** present here (Grade A QR IMAGE.LOAD) and in v1.36 scan #17, but ABSENT in scan #16 (Grade F DM live). Tentative pattern: `stats` correlates with QR symbology or IMAGE.LOAD scans, not with grade. v1.38 probe still needed to enumerate `stats` sub-keys.

### NEW field — `FactoryCalibrated`

`FactoryCalibrated: 'false'` — new field in pre-release firmware push XML.
Not previously seen (only `FieldCalibrated` existed before). Not yet in `VerificationXmlMap` or `VerificationRecord`. **Do not implement until instructed** — log only.

### No-image in DMST TC — resolved by config restore

**CORRECTION 2026-06-24**: CP was not open and was not a factor. The post-verification
image problem was entirely within DMST TruCheck's TC panel. Resolved by config restore
(same NVRAM corruption pattern as the prior episode — see `dmst-image-blank-root-cause.md`).
`JpegImageBase64` is present and correct (19,580 chars). Firmware working normally.
All is well following the restore.

### ISO 29158:2025 (AIM-DPM) — new edition

First observation of 2025 edition. Prior: 2024 max. Parser echoes whatever device
reports — no code change needed.
**CORRECTION 2026-06-24**: The 29158 grading standard on this scan was operator
configuration error — ISO 15415 should have been selected for this printed QR symbol.
Grading standard is user-configurable in DMST TruCheck and is independent of symbology.
No inference about DPM device applying ISO 29158 to QR scans is valid from this data point.

### Other field readings

| Field | Value | Notes |
|---|---|---|
| QR grades | All A | ULP/URP/LLP/HCT/VCT/ALP=A, VIBGrade='-', FIBGrade=A |
| SymbologyId | `]Q1` | Consistent |
| ANUPercent | 1.305032730102539 | Raw value; no ÷100 issue |
| EncodedCharacters | 39 | v1.37 eaLen fallback; unchanged |
| DataCodewords | '' | Still empty — bug #5 |
| ErrorCorrectionBudget | '' | Still empty — bug #6 |
| FactoryCalibrated | 'false' | NEW FIELD — not yet parsed |
| FieldCalibrated | 'false' | Consistent |
| DateTime | 2026-06-23T22:32:33 | Device RTC off by ~1 day |


---

## Scan #19 — QR IMAGE.LOAD Grade A, ISO 15415:2011, DPM device (2026-06-24)

**Device**: DM475-DPM-866D76-VCCS-Verif-Lab  
**Firmware**: Pre-release  
**GradingStandard**: `ISO 15415:2011` ✓ (correct standard; scan #18 used wrong 29158)  
**Symbol**: QR 29×29, IMAGE.LOAD, GUID `0e95e424-3a33-eb11-a816-001dd80187c1`  
**FormalGrade**: 4/A, OverallGrade: A, SymbolQuality: 100  
**ApertureRef**: 16 (was 17 on scan #18 with 29158 — aperture is standard-dependent)  
**OpticsSource**: LoadedImage (CU=-1, MRD=-1 confirmed)  
**JpegImageBase64**: 18,716 chars — JPEG present  
**DebugBarcodeAssignment**: `result=-1;stats=[obj];`  
**FactoryCalibrated**: 'false' (confirmed again — new firmware field)  

All QR grades A; VIBGrade='-'; FIBGrade=A. Consistent with scan #13/#15/#18.

---

## Scan #20 — QR IMAGE.LOAD Grade F (TOTAL FAIL), ISO 15415:2011, DPM device (2026-06-24)

**Device**: DM475-DPM-866D76-VCCS-Verif-Lab  
**Firmware**: Pre-release  
**GradingStandard**: `ISO 15415:2011`  
**Symbol**: QR 29×29, IMAGE.LOAD, same GUID (intentional degraded image)  
**FormalGrade**: 0/F, OverallGrade: F, SymbolQuality: 0  
**ApertureRef**: 16  
**OpticsSource**: LoadedImage (CU=-1, MRD=-1 — confirmed IMAGE.LOAD even on total fail)  
**JpegImageBase64**: 18,496 chars — **JPEG present even on total fail** ← notable  
**DebugBarcodeAssignment**: `result=-1;stats=[obj];` — **grade-independent CONFIRMED**  
**FactoryCalibrated**: 'false'

### Total-fail QR field behaviour (ISO 15415, IMAGE.LOAD)

| Field | Total-fail value | Notes |
|---|---|---|
| SymbolQuality | '0' | Not '' or '-1' |
| UECPercent / SCPercent | '0' | Zero, not empty |
| ANUPercent / GNUPercent | '0' | Zero, not empty |
| FPDValue | '0' | Zero |
| All grade letters (ULP…FIB) | 'F' | All F including VIBGrade='F' |
| MatrixSize | '' | Empty — decode failed |
| EncodedCharacters | '' | Empty |
| NominalXDim | '' | Empty |
| ErrorsCorrected / ErrorCapacityUsed | '' | Empty |
| ApplicationPass | 'Fail (Quality)' | NEW variant — see below |
| ApplicationPassReason | 'Quality' | NEW variant |
| DDGrade / AverageGrade | 'X' | Same as passing scans — no change |
| AverageGradeNumeric | '8.8' | Same sentinel — no change |
| JpegImageBase64 | 18,496 chars | Image sent even on total fail |

**VIBGrade='F' on total fail** — reconfirms prior finding (live QR fail scan #3). Parser handles both '-' (passing v3) and 'F' (total fail).

### ApplicationPass variants — now 3 observed

| ApplicationPass | ApplicationPassReason | Condition |
|---|---|---|
| 'Pass' | '' | ISO grade pass |
| 'Fail (Quality)' | 'Quality' | ISO grade total fail (scan #20) ← NEW |
| 'Fail (Data Format)' | 'Data Format' | GS1 format fail (Code 128 scan #9) |

### ★ barcodeAssignment — FULLY RESOLVED 2026-06-24

`result=-1;stats=[obj];` on both pass (A) and fail (F) QR IMAGE.LOAD scans.

Combined evidence across all scans:

| Scan | Symbology | Grade | OpticsSource | result | stats |
|---|---|---|---|---|---|
| #16 | DM 22×22 | F | Live | -1 | absent |
| #18 | QR 29×29 | A | IMAGE.LOAD | -1 | [obj] |
| #19 | QR 29×29 | A | IMAGE.LOAD | -1 | [obj] |
| #20 | QR 29×29 | F | IMAGE.LOAD | -1 | [obj] |

**CONCLUSION**: `result=-1` = "no barcode assignment rule configured on this device."
Completely grade-independent, standard-independent, and symbology-independent.
Device-config sentinel. Passing-scan probe closed — no further investigation needed.

**stats=[obj] pattern**: present on all QR scans (pass and fail); absent on DM live.
Tentative: `stats` correlates with QR symbology. One more DM pass scan would confirm.
Sub-key enumeration remains a v1.38 candidate but is no longer blocking anything.


---

## Scan #21 — QR IMAGE.LOAD Grade F (DECODED), ISO 15415:2011, DPM device (2026-06-24)

**Device**: DM475-DPM-866D76-VCCS-Verif-Lab  
**GradingStandard**: `ISO 15415:2011`  
**Symbol**: QR 29×29, IMAGE.LOAD, same GUID — partial degradation (not total failure)  
**FormalGrade**: 0/F, OverallGrade: F, SymbolQuality: **46** (decoded, partial fail)  
**DecodedData**: `0e95e424-3a33-eb11-a816-001dd80187c1` — **decode succeeded**  
**DecodeGrade**: 'A' — decode perfect  
**DebugBarcodeAssignment**: `result=-1;stats=[obj];` — consistent  
**JpegImageBase64**: 18,600 chars present  
**FactoryCalibrated**: 'false'

### Three-way Grade F comparison — decoded F vs total fail F

This scan establishes the critical distinction between:
- **Scan #20**: total decode failure (no decode, SymbolQuality=0, all grades F)
- **Scan #21**: decoded but grade F (MOD+RM fail, SymbolQuality=46)

| Field | Scan #19 (A, pass) | Scan #21 (F, decoded) | Scan #20 (F, no decode) |
|---|---|---|---|
| SymbolQuality | 100 | 46 | 0 |
| DecodeGrade | 'A' | 'A' | 'F' |
| MatrixSize | '29x29' | '29x29' | '' |
| EncodedCharacters | '39' | '39' | '' |
| NominalXDim | '21.5 mil' | '21.4 mil' | '' |
| ErrorsCorrected | '0' | '8' | '' |
| ErrorCapacityUsed | '0' | '16' | '' |
| UECGrade | 'A' | **'C'** | 'F' |
| MODGrade | 'A' | **'F'** | 'F' |
| RMGrade | 'A' | **'F'** | 'F' |
| FPDGrade | 'A' | **'D'** | 'F' |
| ALPGrade | 'A' | **'D'** | 'F' |
| ULPGrade/URPGrade/LLPGrade | 'A' | 'A' | 'F' |
| HCTGrade/VCTGrade | 'A' | 'A' | 'F' |
| **VIBGrade** | **'-'** | **'-'** | **'F'** |
| FIBGrade | 'A' | 'A' | 'F' |
| HorizontalBWG | '0' | **'1'** | absent |
| VerticalBWG | '0' | **'5'** | absent |

### ★ VIBGrade — rule now fully established

- `'-'` = "not applicable" for v1–6 QR (no VIB to grade). Appears on **any decoded scan**, pass OR grade F.
- `'F'` = total decode failure only (no decode → no grade → 'F' sentinel for all grades).
- `'-'` does NOT mean "passing VIB." It means the VIB grading parameter does not apply to this version.

### ErrorsCorrected / ErrorCapacityUsed population rule — confirmed

- Populated (non-empty) whenever symbol decodes successfully: `'0'` on clean pass, `'8'`/`'16'` on error-corrected decode.
- Empty (`''`) only on total decode failure (scan #20).
- Aligns with expected behaviour: these fields require a successful decode to compute.

### DataCodewords / ErrorCorrectionBudget — still empty

Both still `''` on this scan (decoded, non-trivial error correction). Confirms these remain
unresolvable from push XML. C# table lookup (bug #5/#6) still required.


---

## Wireshark capture #2 — DM475V-DPM (10.10.10.4) live mode toggle — 2026-06-24

**Capture condition**: Verifier idle → live mode ON → OFF → ON → OFF (two complete cycles)  
**Full analysis**: `vtccp/architecture/wireshark-protocol-analysis.md` §8

### Live mode control — CONFIRMED

```
GET /monitormode?enable=true HTTP/1.1   →  HTTP/1.1 204 No Content
GET /monitormode?enable=false HTTP/1.1  →  HTTP/1.1 204 No Content
```

- Port 44444, plain HTTP GET
- **NOT** a raw DMCC text command
- Response is immediate `204 No Content` — clean ack

Four toggle events observed, all confirmed. Same protocol as LBL device.

### VERIFICATION.ENABLE — ABSENT

Not present anywhere in the capture. TruCheck verification is always active during monitor
mode on the DM475V. There is no separate per-session verification enable/disable toggle.
The DMCC reference `VERIFICATION.ENABLE` command applies to older models (DM370/DM390/DM470)
and was not observed in any DM475V session capture.

### GET /svg_image.img — CONFIRMED 500 on all polls

Live image streaming is inaccessible to third-party HTTP clients. Every poll returns
`HTTP/1.1 500 Internal Server Error`. DMST's live view uses a different mechanism.

### Two-connection architecture — confirmed (DPM device, same as LBL)

| Conn | Local port | Role |
|---|---|---|
| Subscription | 55653 | Device pushes PUT events here |
| Command | 55654 | DMST sends GET /monitormode and GET /svg_image here |

### For VTCCP live mode control (future)

```
GET /monitormode?enable=true HTTP/1.1\r\n
Host: {device-ip}:44444\r\n
\r\n
```

Use a separate `TcpClient` — do NOT reuse the SDK DMCC connection.
No `MONITOR-MODE.ENABLE` DMCC command required.

---

## Full cold-start session capture — DPM (10.10.10.4) — 2026-06-24

**Capture**: DMST cold start → TC connect → Go Live → Trigger → Go Live → Cancel
**Full analysis**: `vtccp/architecture/wireshark-protocol-analysis.md` §9

### Complete connection sequence confirmed

1. UDP port 1069 discovery broadcasts (device → before TCP)
2. Two TCP connections: 54767 (events/sub) + 54768 (command)
3. Continuations → 200 + 204 on 54768 (init, URLs unknown)
4. `GET /events?enable` → 204 on 54767
5. `GET /vs.cfg` ×2, `GET /parameters.xml` (NEW), `GET /status.xml` ×2, `GET /device_info.xml` → 401 ×2
6. `GET /monitormode?enable=true` → Sleep
7. Trigger: Continuation ×2 → 204 each (URL unknown)
8. Scan result: vs.cfg + pcm_report.html + codes.xml + **svg_image.img** + status.xml ×2
9. `GET /monitormode?enable=true` → Sleep (repeat)
10. `GET /monitormode?enable=false` → Cancel

### ★ PUT /svg_image.img — device PUSHES the image

Device PUTs the scan/live image on the events channel (not DMST polling via GET).
GET /svg_image.img → 500 because direction is wrong. PUT is correct.
For VTCCP live image: receive PUT /svg_image.img on events subscription connection.

### GET /parameters.xml — new, large, content unknown

Fetched during TC connect init. Large response. Likely full device parameter dump.
Worth capturing the body — may replace DMCC GET ALL.

### Still unknown

- Trigger URL (need Follow TCP Stream on pkt 2035 or 4307)
- Init URLs (Continuation pkts 505/508 at connection open)
- Image format in PUT /svg_image.img body
- Content of GET /parameters.xml response
