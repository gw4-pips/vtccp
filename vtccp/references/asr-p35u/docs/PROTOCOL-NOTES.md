# AsReader P35U — Protocol Notes for C# Integration

**Source:** VCCS RFID FlexWedge Pro development, 2026
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
    CallBackReadTagData    cbTag,       // fires on every tag read
    CallBackErrorCode      cbError,     // hardware error
    CallBackSuccessCode    cbSuccess,   // command accepted
    CallBackCommandData    cbCommand,   // raw command response bytes (see TID note)
    CallBackReadComplete   cbComplete,  // inventory round finished
    CallBackTriggerHandler cbTrigger    // hardware trigger button pressed
)
```

All six delegates must be registered in a single `SetDelegate()` call.
Re-registration (e.g. to change one callback) requires passing all six again.

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

### Tag Data Callback (`CallBackReadTagData`)

```csharp
void CallBackReadTagData(InventoryResult result)
```

`InventoryResult` fields observed in Python (all are nullable):

| Field              | Type   | Example                      | Notes                          |
|--------------------|--------|------------------------------|--------------------------------|
| `result.tagdata.epc`  | string | `"30342A7CC844C7D0F36A0676"` | Uppercase hex, no spaces      |
| `result.tagdata.pc`   | string | `"3000"`                     | Hex string — convert with int(pc, 16) |
| `result.tagdata.tid`  | string | `"E28011920008C7C"`          | Populated by ReadMemory (see below) |
| `result.tagdata.data` | string | raw memory bytes             | ReadMemory result arrives here |
| `result.rssi`         | float  | `-35.0`                      | May arrive as unsigned byte 128–255 = two's complement |
| `result.antenna`      | int    | `1`                          |                                |

**RSSI note:** The DLL sometimes delivers RSSI as an unsigned byte (0–255).
Values 128–255 represent negative dBm via two's complement: `rssi_dbm = raw - 256`.
Values 0–127 are returned as-is.

### StopInventory
```csharp
uint StopInventory()
// Fires cbComplete asynchronously — do not assume inventory has stopped
// until the cbComplete callback fires.
```

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

**Critical behaviour difference between firmware versions:**

| Firmware | ReadMemory result callback |
|----------|---------------------------|
| ≤ 1.2.0  | `CallBackCommandData` (never actually observed to fire — possible DLL bug) |
| 1.8.0    | `CallBackReadTagData` — result arrives in `tagdata.data` or `tagdata.tid` |

On firmware 1.8.0 (current), always hook `CallBackReadTagData` for ReadMemory
results. Do NOT wait only on `CallBackCommandData` — it will not fire.

Confirmed working sequence for single-tag TID read (firmware 1.8.0):
```
1. Run inventory until tag is detected (cbTag fires with EPC)
2. Stop inventory (StopInventory → wait for cbComplete)
3. Call ReadMemory(MEM_TID, 0, 4, 0, epcBytes)
4. Next cbTag call delivers TID in tagdata.data / tagdata.tid
5. Parse TID hex string for manufacturer and model info
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
REGION_US = (observed as default for North American unit)
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

**Untested methods (calling from Python/pythonnet crashes the process):**
- `SetHIDInventoryMode` — parameter type unknown; causes pythonnet MethodBinder crash
- `GetHIDWorkParams` — same issue
- `Kill` — not tested; expected to kill tag (Gen2 Kill command)
- `WriteMemory` — not tested; expected to write EPC/User bank

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
A robust implementation polls `IsConnected()` (or equivalent) at 1–2 s intervals
to detect idle disconnects.

After disconnect: call `DisConnect()`, then `ConnectWithVCP()` to reconnect.
Do NOT call `StartInventory()` on a disconnected device — undefined behaviour.
