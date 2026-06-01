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
