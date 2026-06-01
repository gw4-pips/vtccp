# VTCCP Working Notes

> **Restore point**: commit `f474fa7` — AS-IS state as of 2026-05-29.
> All trigger investigation work begins AFTER this point.
> Do not build without explicit user instruction.

---

## RESOLVED: Image missing from TC/DMST screen post-verification

**Status**: ROOT CAUSE CONFIRMED 2026-05-31 — fix deployed in DataManSdkClient.cs.

**Symptom**: After a verification scan, the DMST TC screen and DMST verification panel
show no barcode image. HTML report still contains the image. Push XML `JpegImageBase64`
also confirmed populated (v1.36 scan #16). This is a result-channel delivery failure,
not a capture failure.

**Confirmed root cause**: `LIVEIMG.MODE = 0` while VTCCP is connected.
The Cognex SDK's `Connect()` sets `LIVEIMG.MODE` to `0` ("no image with result")
and persists it. `LIVEIMG.MODE = 2` means "send image with each result."
`COM.DMCC-RESET` does NOT restore this — it is a CONFIG parameter, not a DMCC
session parameter. `DATA.IMAGE-TYPE` was confirmed correct (= 0) and was never the cause.

**Fix**: `SendDmccRestoreAsync` now sends four commands on port 23 after every
connect and on disconnect:
1. `COM.DMCC-RESET` — clears SDK-corrupted DMCC session params
2. `COM.DMCC-SAVE` — persists them to NVRAM
3. `SET LIVEIMG.MODE 2` — restores image delivery ← the actual fix
4. `CONFIG.SAVE` — persists LIVEIMG.MODE to flash

**False leads ruled out**:
- `DATA.IMAGE-TYPE` — confirmed = 0 (correct) while VTCCP connected; never the cause
- `DATA.RESULT-TYPE = 513` — factory default; not the cause

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

**Expected values (known-good state):**

| Command | Expected | Notes |
|---|---|---|
| `GET LIVEIMG.MODE` | `2` | Send image with each result. Any other value → no image in TC panel. |
| `GET DATA.IMAGE-TYPE` | `0` (or firmware default) | If non-zero, SDK has persisted an image-strip value via COM.DMCC-SAVE. |
| `GET DATA.RESULT-TYPE` | `0` | Default DMCC result type. |
| `GET DATA.RESULT-ENCODING` | `0` | Default encoding. |
| `GET DATA.RESULT-ALWAYSSEND` | `1` | Results sent on every trigger. |
| `GET IMAGE.FORMAT` | `1` | JPEG. |
| `GET IMAGE.SIZE` | `0` | Full resolution. |

**If any value is wrong**, the corrective sequence (in order):

```
COM.DMCC-RESET
SET LIVEIMG.MODE 2
SET IMAGE.FORMAT 1
SET IMAGE.SIZE 0
CONFIG.SAVE
```

`COM.DMCC-RESET` resets all DMCC session settings (`DATA.IMAGE-TYPE`, `DATA.RESULT-TYPE`,
`DATA.RESULT-ENCODING`, `DATA.RESULT-ALWAYSSEND`, `COM.DMCC-RESPONSE`, `COM.DMCC-CHECKSUM`,
`COM.DMCC-HEADER`) back to firmware defaults in one command. Run it first; it likely clears
the image-strip condition in a single step. Then verify LIVEIMG.MODE and IMAGE.* separately.

`CONFIG.SAVE` persists the restored settings to flash so a power cycle doesn't revert them.

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

**Status**: PARKED — implement when user directs.

Webscan TruCheck Excel uses Excel COM automation (`Microsoft.Office.Interop.Excel`)
to push each row directly into the already-open workbook in memory. No file write,
no lock contention, rows appear instantly as scans arrive — the Excel file stays open
throughout the session.

VTCCP's current per-record `_writer.Save()` (EPPlus `SaveAs`) cannot update an
XLSX that Excel has locked. The sidecar JSON always writes correctly; the XLSX only
reflects new records after the session is closed (CloseSession finalises the file).

**Implementation path when approved**:
- Add `ExcelEngine.Adapters.ComExcelAdapter` implementing `IExcelAdapter`
- Uses `Microsoft.Office.Interop.Excel` (COM, no NuGet — present on any machine with Excel installed)
- On `AddRecord`: `_worksheet.Rows[nextRow].Value = rowData` — writes directly to
  the live workbook object in the running Excel process; no SaveAs required
- `Save()` on the COM adapter becomes a no-op (or `_workbook.Save()` at CloseSession)
- Fallback: if Excel is not running / COM binding fails, fall back to current EPPlus path
- VtccpApp.csproj: add `<COMReference>` for `Microsoft.Office.Interop.Excel` or use
  dynamic binding to avoid a hard compile-time dependency on the interop assembly

**Blocked by**: User approval + decision on whether to require Excel to be open at session
start, or lazily bind on first record.

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
   - `SET LIVEIMG.MODE 2` (image delivered with each result — required for TruCheck)
   - `SET IMAGE.FORMAT 1` (JPEG)
   - `SET IMAGE.SIZE 0` (full resolution)
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
