---
name: RFID UI not yet built
description: RFID DeviceInterface backend is complete but VtccpApp has no UI wired to it as of 2026-08-17.
---

# RFID UI Status (as of 2026-08-17)

## Backend — complete
- `vtccp/DeviceInterface/Rfid/IEpcReader.cs` — interface
- `vtccp/DeviceInterface/Rfid/AsReaderP35UEpcReader.cs` — ASR-P35U implementation (USB/VCP, COM port, 6-delegate SetDelegate)
- `vtccp/DeviceInterface/Rfid/EpcReaderFactory.cs` — factory + port enumeration (lines 47-54)
- `vtccp/DeviceInterface/Rfid/RfidScanCoordinator.cs` — scan orchestration (lines 18-117)
- `vtccp/ExcelEngine/Models/VerificationRecord.cs` — RFID fields at lines 443-528
- `vtccp/ExcelEngine/Schema/RfidTabSchema.cs` — Excel tab schema

## UI — not built
- No COM port selector, no connect/disconnect button, no RFID status in VtccpApp
- No ViewModel wires RfidScanCoordinator
- SettingsView.xaml has no RFID controls
- SessionView.xaml has no RFID panel

## What needs building (Task #133)
- RFID section in Session Launcher: COM port dropdown, Connect/Disconnect, status indicator
- Wire RfidScanCoordinator into SessionViewModel
- Auto-connect on session start if port configured
- EPC tags merge into VerificationRecord per scan

**Why:** Confirmed by subagent exploration of all VtccpApp XAML and ViewModels — exhaustive search found zero RFID/AsReader/IEpcReader/EPC bindings.
