# AsReader P35U — Protocol Notes for C# Integration

**Source:** VCCS RFID FlexWedge Pro development, 2026
**Updated:** 2026-08-17 — vendor engineering response incorporated
**Purpose:** Engineering reference for implementing `IEpcReader` in vtccp / Command Pilot

---

## Connection Type

**USB → VCP (Virtual COM Port)**

The P35U enumerates as a standard COM port on Windows 10/11.  No FTDI driver is
required — the device uses the AsReader proprietary USB CDC driver (installed
automatically by Windows Update or bundled with the SDK zip).

| Parameter  | Value         |
|------------|---------------|
| VID        | 0x339C        |
| PID        | 0x271B        |
| Baud rate  | 115200        |
| Data bits  | 8             |
| Parity     | None          |
| Stop bits  | 1             |
| Flow ctrl  | None          |

**Important:** Never open the COM port directly.  The SDK (`AsReaderP3xU.dll`)
manages the serial port internally.  All communication goes through
`AsReader.ConnectWithVCP(comPort)`.  Bypassing the SDK with raw serial reads
leaves the device busy for ~5 s and corrupts subsequent callbacks.

---

## SDK Overview

- **DLL:** `AsReaderP3xU.dll` (SDK version 1.3.0, 2026-02-13)
- **Namespace:** `AsReaderP3xU`
- **Main class:** `AsReader`
- **Helper types:** `Types` (enums), `InventoryResult` (tag data struct)
- **Runtime:** .NET Framework 4.x (pre-installed on Windows 10/11)
- **Python access:** via `pythonnet` (`import clr; clr.AddReference('AsReaderP3xU')`)

---

## Initialisation Sequence

```
1. new AsReader()
2. SetDelegate(...)          — register all 6 callbacks before connecting
3. ConnectWithVCP("COM4")   — returns 0 on success
4. SetRegion(REGION_US)     — required; REGION_EU also available
5. SetTxPower(dBm)          — 13–27 dBm; 20 dBm is a safe working default
6. [optional] SetReadTime / SetIdelTime / SetSession / SetAntiCollisionMode
7. StartInventory(...)
```

---

## Key SDK Calls

### ConnectWithVCP
```csharp
uint ConnectWithVCP(string comPort)   // e.g. "COM4"
// Returns 0 = success, non-zero = failure
```

### SetDelegate
```csharp
void SetDelegate(
    CallBackReadTagData    cbTag,       // fires on every tag read AND on ReadMemory results
    CallBackErrorCode      cbError,     // hardware error
    CallBackSuccessCode    cbSuccess,   // command accepted
    CallBackCommandData    cbCommand,   // firmware update packets ONLY — never for ReadMemory
    CallBackReadComplete   cbComplete,  // inventory round finished
    CallBackTriggerHandler cbTrigger    // hardware trigger button pressed
)
```

All six delegates must be registered in a single `SetDelegate()` call.
Re-registration (e.g. to change one callback) requires passing all six again.

> ⚠️ **CallBackCommandData is NOT for ReadMemory.**  Per vendor engineering
> (confirmed 2026-08-17), `CallBackCommandData` is reserved exclusively for
> firmware update packets (address request / file transfer / reboot / RFID
> module update).  It will **never** fire for `ReadMemory()` results regardless
> of firmware version, mode, or configuration.  Always use `CallBackReadTagData`
> to receive ReadMemory results — see below.

### StartInventory
```csharp
uint StartInventory(
    bool   rssiEnabled,   // true = include RSSI in tag data
    int    maxTags,       // 0 = unlimited; 1 = stop after first
    int    maxSecs,       // 0 = unlimited
    int    maxCycles,     // 0 = unlimited
    int    antenna        // 1 = antenna 1 (only antenna on P35U desktop unit)
)
```

For a single-read mode (C# equivalent of FlexWedge "Stealth" preset):
```csharp
StartInventory(rssiEnabled: true, maxTags: 1, maxSecs: 0, maxCycles: 0, antenna: 1)
```
The DLL fires `cbComplete` when the tag limit is reached.

For continuous mode:
```csharp
StartInventory(rssiEnabled: true, maxTags: 0, maxSecs: 0, maxCycles: 0, antenna: 1)
// Call StopInventory() to halt.
```

### Tag Data Callback (`CallBackReadTagData`) — also delivers ReadMemory results

```csharp
void CallBackReadTagData(InventoryResult result)
```

This callback fires for **both** inventory tag reads and `ReadMemory()` results.
When fired for a ReadMemory response, the tag data fields are:

| Field              | Type   | Example                      | Notes                          |
|--------------------|--------|------------------------------|--------------------------------|
| `result.tagdata.epc`  | string | `"30342A7CC844C7D0F36A0676"` | Uppercase hex, no spaces      |
| `result.tagdata.pc`   | string | `"3000"`                     | Hex string — convert with int(pc, 16) |
| `result.tagdata.tid`  | string | `"E28011920008C7C"`          | TID memory bank contents       |
| `result.tagdata.data` | string | raw memory bytes             | Also carries ReadMemory result |
| `result.rssi`         | float  | `-35.0`                      | May arrive as unsigned byte 128–255 = two's complement |
| `result.antenna`      | int    | `1`                          |                                |

Check `tagdata.tid` first; fall back to `tagdata.data` if `tid` is null/empty.

**RSSI note:** The DLL sometimes delivers RSSI as an unsigned byte (0–255).
Values 128–255 represent negative dBm via two's complement: `rssi_dbm = raw - 256`.
Values 0–127 are returned as-is.

### ReadMemory (TID)
```csharp
uint ReadMemory(
    Types.MemBankType membank,   // MEM_TID = 0x02
    uint offset,                 // 0 = start of bank
    uint length,                 // word count (4 words = 8 bytes = 64-bit TID)
    uint password,               // 0 = no access password
    byte[] epc                   // EPC bytes of the target tag
)
// Returns 0 = command accepted.
```

**Result delivery:** always via `CallBackReadTagData` — check `tagdata.tid` / `tagdata.data`.

Confirmed working sequence for single-tag TID read:
```
1. Run inventory until tag is detected (cbTag fires with EPC)
2. Stop inventory (StopInventory → wait for cbComplete)
3. Call ReadMemory(MEM_TID, 0, 4, 0, epcBytes)
4. Next cbTag call delivers TID in tagdata.tid (or tagdata.data)
5. Parse TID hex string for manufacturer and model info
```

### StopInventory
```csharp
uint StopInventory()
// Fires cbComplete asynchronously — do not assume inventory has stopped
// until the cbComplete callback fires.
```

### CheckTagStatus (Lock Check)
```csharp
uint CheckTagStatus(byte[] epc)
// Asynchronous — result arrives via cbSuccess / cbError.
// cbSuccess code values observed:
//   40 = PermaLock (all memory banks permanently locked)
//   41 = Lock      (write-protected, not permanently)
//   42 = Unlock    (no lock; tag is writable)
```

### Power Control
```csharp
uint SetTxPower(int dBm)   // valid range: 13–27
uint GetTxPower(ref int dBm)
```

---

## Callback Purpose Reference

| Callback               | Fires for                                      |
|------------------------|------------------------------------------------|
| CallBackReadTagData    | Every inventory tag read + every ReadMemory result |
| CallBackErrorCode      | Hardware errors                                |
| CallBackSuccessCode    | Commands accepted (e.g. CheckTagStatus result) |
| **CallBackCommandData**| **Firmware update packets only** — address request, file transfer, reboot, RFID module update |
| CallBackReadComplete   | Inventory round complete                       |
| CallBackTriggerHandler | Hardware trigger button pressed                |

---

## Enum Values

### Types.MemBankType
```
MEM_RESERVED = 0x00
MEM_EPC      = 0x01
MEM_TID      = 0x02
MEM_USER     = 0x03
```

### Types.RegionType
```
REGION_US = (default for North American unit)
REGION_EU
```

### Types.SessionType
```
SESSION_S0 = 0
SESSION_S1 = 1   ← FlexWedge default
SESSION_S2 = 2
SESSION_S3 = 3
```

### Types.AntiCollisionMode
```
FixedQ   — fixed Q value (use SetQuery to set Q)
DynamicQ — dynamic Q algorithm (recommended for most use cases)
```

---

## Error Codes

| Code | Meaning                                   |
|------|-------------------------------------------|
| 0    | Success / command accepted                |
| 1    | Command failed (returned by ReadMemory)   |
| 4    | Device conflict (CheckTagStatus while busy) |

cbError codes observed: any non-zero value arriving during active inventory
indicates a hardware disconnect. Arriving while idle = spurious; log and ignore.

---

## Full DLL Public Member List (SDK 1.3.0)

```
CallBackCommandData, CallBackErrorCode, CallBackReadComplete,
CallBackReadTagData, CallBackSuccessCode, CallBackTriggerHandler,
CheckTagStatus, ConnectWithVCP, DefaultSetting, DisConnect,
Equals, Finalize, GetAntiCollisionMode, GetBasicTarget,
GetBuzzer, GetChannel, GetFH_LBT, GetFrequencyAutomatic,
GetFwVersion, GetHIDInventoryMode, GetHIDWorkParams, GetHashCode,
GetHwVersion, GetIdelTime, GetProductSN, GetQuery,
GetRFIDFwVersion, GetRSSIThreshold, GetReadTime, GetRegion,
GetSdkVersion, GetSelectMask, GetSelectionEnable, GetSession,
GetTxPower, GetType, InventoryResult, Kill, LockMemory,
MemberwiseClone, Overloads, ReadMemory, ReferenceEquals,
SendCommand, SetAntiCollisionMode, SetBasicTarget, SetBuzzer,
SetChannel, SetDelegate, SetFH_LBT, SetFrequencyAutomatic,
SetHIDInventoryMode, SetHIDWorkParams, SetIdelTime, SetQuery,
SetRSSIThreshold, SetReadTime, SetRegion, SetSelectMask,
SetSelectionEnable, SetSession, SetTxPower, StartInventory,
StopInventory, TagAction, TagData, TagMask, ToString,
WriteMemory
```

**Untested/unsupported from Python (pythonnet binding crash):**
- `SetHIDInventoryMode` — parameter type unknown; vendor notes SDK is C#-only officially
- `GetHIDWorkParams` — same issue
- `Kill` — not tested; Gen2 Kill command
- `WriteMemory` — not tested

---

## Hardware Specs (Unit KE00048)

| Field          | Value                   |
|----------------|-------------------------|
| S/N            | KE00048                 |
| Main FW        | 1.8.0 (updated 2026-08-10) |
| RFID module FW | RED4S_v2.2.2_K_SD       |
| HW Version     | 1.0.2                   |
| SDK Version    | 1.3.0 (2026-02-13)      |
| COM port (dev) | COM4 (Windows 11 laptop)|
| TX Power range | 13–27 dBm               |
| Region         | REGION_US               |

---

## Disconnect Handling

The DLL fires `cbError` on unexpected cable pull while inventory is running.
It does NOT fire any callback if the cable is pulled while the reader is idle.
After disconnect: call `DisConnect()`, then `ConnectWithVCP()` to reconnect.
Do NOT call `StartInventory()` on a disconnected device.
