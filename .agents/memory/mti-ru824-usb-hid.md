---
name: MTI READ ME RU-824-100 USB connection
description: Device presents as USB HID (not VCP/COM port); SDK required for USB mode; MSI installer in repo
---

## Rule
The MTI READ ME RU-824-100 connects via USB HID — it does NOT create a virtual COM port. `System.IO.Ports.SerialPort` approach in `MtiLlcsEpcReader.cs` must be rewritten to use the `LakeChabotReader` SDK class.

**Why:** Windows Device Manager confirmed +2 HID entries on plug/unplug; no COM port created. ReadMe.txt confirms "defaults to enumerating R1000 radios attached via USB" via native rfid.dll HID transport.

**How to apply:**
- UART mode IS available in SDK but requires physical RS-232 port (device has none) + `RFIDcomm.cfg` with `Port=N`
- USB mode: use `LakeChabotReader.EnumerateReaders()` → connect → inventory. SDK handles HID internally.
- Native DLLs required at runtime: `rfid.dll`, `Linkage.dll`, `RFIDInterface.dll`, `rfidtx.dll`, `cpl.dll`
- MSI installer: `vtccp/references/mti-sdk/RFID_Explorer/MTI RFID Explorer v2.0.1.msi` (1.9 MB, in repo)
- After install on Win machine: DLLs at `C:\Program Files\MTI\MTI Explorer v2.0.1\`
- SDK source: `vtccp/references/mti-sdk/RFID_Explorer/MTI RFID Explorer v2.0.1 Source/RFIDInterface/Source/LakeChabot.cs`
- RU-824 Command Reference Manual: `vtccp/references/mti-sdk/RFID_Explorer/MTI RU-824 RFID Module Command Reference Manual v3.3.pdf`
- Correct full product name: **MTI READ ME RU-824-100** (not "ME RU-824-100")
