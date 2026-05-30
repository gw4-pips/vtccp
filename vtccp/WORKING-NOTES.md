# VTCCP Working Notes

> **Restore point**: commit `f474fa7` — AS-IS state as of 2026-05-29.
> All trigger investigation work begins AFTER this point.
> Do not build without explicit user instruction.

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
