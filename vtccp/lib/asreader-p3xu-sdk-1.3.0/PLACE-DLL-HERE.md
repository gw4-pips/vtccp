# AsReader P3xU SDK DLL — Placement Instructions

Place `AsReaderP3xU.dll` (SDK v1.3.0) in **this folder** before building.

The DLL is not committed to the repository (vendor binary, not redistributable).

---

## How to obtain

1. Download the AsReader SDK zip from the AsReader developer portal.
   Filename: `AsReader_P35U_SDK_c#_1_3_0.zip` (or similar).
2. Extract and copy `AsReaderP3xU.dll` into this folder:
   `vtccp\lib\asreader-p3xu-sdk-1.3.0\AsReaderP3xU.dll`

## Hardware reference (unit KE00048)

| Field          | Value                       |
|----------------|-----------------------------|
| SDK version    | 1.3.0 (2026-02-13)          |
| Main firmware  | 1.8.0 (updated 2026-08-10)  |
| RFID module FW | RED4S_v2.2.2_K_SD           |
| HW version     | 1.0.2                       |
| VID / PID      | 0x339C / 0x271B             |
| COM port       | COM4 (dev laptop, Windows 11) |

## Known defect (FW 1.8.0 / SDK 1.3.0)

`CallBackCommandData` never fires for `ReadMemory()`.  ReadMemory results arrive
via `CallBackReadTagData` instead (`tagdata.data` or `tagdata.tid`).  Vendor
notified 2026-08-08.  See `vtccp/references/asr-p35u/docs/ASREADER_TID_DEFECT.md`.

## Build notes

The DLL targets .NET Framework 4.x and is referenced from `DeviceInterface.csproj`
via `<Reference>` with a relative `HintPath` — the same pattern used for the
Cognex DataMan SDK.  The DLL must be present at compile time and is copied to the
output directory (`<Private>true</Private>`).
