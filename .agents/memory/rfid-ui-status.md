---
name: RFID UI wiring
description: Session Launcher RFID panel exists and is wired to the DeviceInterface backend; key lifetime/dep decisions to stay consistent with.
---

# RFID UI Wiring (built 2026-08-17)

Session Launcher (SessionView.xaml / SessionViewModel) now has an RFID panel:
COM port dropdown + refresh, Connect/Disconnect, status dot + message.

## Durable decisions
- **Reader lifetime is owned by the ViewModel, not the coordinator.** The
  ASR-P35U reader connects once via the panel and survives across sessions.
  `RfidScanCoordinator` takes `ownsReader: false` so disposing the per-session
  coordinator does not disconnect the reader. Disconnect happens only via the
  panel button or app exit (`StopSessionOnExitAsync`).
  **Why:** operators run many short sessions; reconnecting per session is slow
  and flaky on the ASR-P35U VCP.
- **Port enumeration in the UI uses `System.IO.Ports.SerialPort.GetPortNames()`
  directly**, not `EpcReaderFactory`/`AsReaderP35UEpcReader` — those files are
  excluded from compilation when `AsReaderP3xU.dll` is absent, but the picker
  must compile everywhere. Only reader creation sits behind `#if ASREADER_SDK`.
- **Selected port is persisted to `AppSettings.RfidComPort` on successful
  connect**; session start auto-connects when a port is selected but not
  connected yet.
- **"RFID Scans" tab writing:** `RfidTabWriter` (DeviceInterface) is driven from
  SessionViewModel using `SessionManager.Adapter` / `LastSummaryRow` /
  `MainSheetName` — ExcelEngine cannot reference DeviceInterface (dep rule), so
  the app layer bridges. Callers must restore the main sheet after aux writes.
