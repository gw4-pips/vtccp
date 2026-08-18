---
name: AsReader ASR-P35U SDK Architecture
description: Key SDK facts, callback model, FW defects, and C# integration decisions for AsReaderP35UEpcReader.
---

## SDK basics
- DLL: `AsReaderP3xU.dll` v1.3.0 (place in `vtccp\lib\asreader-p3xu-sdk-1.3.0\`)
- Namespace: `AsReaderP3xU`; main class: `AsReader`; enums in `Types`
- Connection: VCP (no FTDI needed). VID=0x339C / PID=0x271B. COM4 on dev laptop.
- SDK manages the serial port — never open the port directly (leaves device locked ~5s)

## Init sequence (must follow this order)
1. `new AsReader()`
2. `dev.SetDelegate(...)` — all six callbacks in one call, BEFORE ConnectWithVCP
3. `dev.ConnectWithVCP("COM4")` → returns 0 on success
4. `dev.SetRegion(Types.RegionType.REGION_US)` — required before StartInventory
5. `dev.SetTxPower(dBm)` — valid range 13–27 dBm (REGION_US)

## StartInventory signature (confirmed from compiler errors against real DLL)
`StartInventory(bool rssiEnabled, byte maxTags, byte maxSecs, ushort maxCycles, bool an1)`
- Named params NOT supported — positional only
- an1 = antenna 1 enable (bool true)
- maxTags=1 → hardware auto-stops after first tag (eliminates timer race)
- Call: `StartInventory(true, 1, 0, 0, true)`

## SetTxPower signature
`SetTxPower(uint dBm)` — takes uint, not int; cast at call site

## ReadMemory signature
`ReadMemory(MemBankType, startAddr, length, password, epc)` — positional only; named param "offset" does not exist

## TagData type
`AsReader.TagData` is a **struct** (value type) — do NOT use `?.` null-conditional operator on it.
Access fields directly: `td.epc`, `td.data`, `td.tid`, `td.pc`

## Callbacks (six, all mandatory in SetDelegate)
| Delegate | Param | Notes |
|---|---|---|
| CallBackReadTagData | InventoryResult result | result.tagdata (TagData struct — not nullable), .epc/.pc/.tid/.data; result.rssi (float) |
| CallBackErrorCode | **uint** errorCode | Non-zero during active inventory = hardware disconnect |
| CallBackSuccessCode | **uint** code | 40=PermaLock, 41=Lock, 42=Unlock (CheckTagStatus) |
| CallBackCommandData | byte[] data | **NEVER fires for ReadMemory on FW 1.8.0** (confirmed DLL defect) |
| CallBackReadComplete | bool completeStatus | true + _hwStopExpected = clean auto-stop; false = unexpected disconnect |
| CallBackTriggerHandler | int state | 1=button pressed, 0=released |

## Critical FW 1.8.0 defect: CallBackCommandData never fires for ReadMemory
- **ReadMemory result arrives via CallBackReadTagData** (tagdata.data or tagdata.tid)
- Workaround: register a one-shot `_pendingTidCb` on the next cbTag before calling ReadMemory
- TID sequence: StartInventory(maxTags=1) → cbTag(EPC) → cbComplete → ReadMemory → next cbTag(TID)
- Vendor notified 2026-08-08; unresolved as of 2026-08-11
- Defect report: `vtccp/references/asr-p35u/docs/ASREADER_TID_DEFECT.md`

## RSSI correction
Values 128–255 from SDK = negative dBm via two's complement: `rssi_dbm = raw - 256`

## C# implementation
- `AsReaderP35UEpcReader.cs` — implements IEpcReader
- Callback-to-async bridge: `TaskCompletionSource<bool>` + `_stateLock` + `_pendingResults` list
- `_pendingTidCb`: volatile Action field for one-shot ReadMemory result hook
- `ReadTidAsync(byte[] epcBytes, TimeSpan timeout)`: explicit TID read with defect workaround
- `EpcReadResult.Tid`: nullable string field added for TID hex string
- `EpcReaderFactory.CreateAsReaderP35U()`: primary factory; E310/MTI marked [Obsolete]

## Build requirement
DLL must be placed at `vtccp\lib\asreader-p3xu-sdk-1.3.0\AsReaderP3xU.dll` before building.
Same pattern as Cognex SDK reference in DeviceInterface.csproj.

**Why:** DLL is .NET Framework 4.x vendor binary; not on NuGet; not redistributable.

## Reference files
- Protocol notes: `vtccp/references/asr-p35u/docs/PROTOCOL-NOTES.md`
- Python reference: `vtccp/references/asr-p35u/source/reader.py`
- Test vectors: `vtccp/references/asr-p35u/test-vectors/epc-decode-vectors.json`
- TID defect: `vtccp/references/asr-p35u/docs/ASREADER_TID_DEFECT.md`

## CheckTagStatus lock-check hazard (FW 1.8.0)
A TIMED-OUT TID ReadMemory emits a delayed stray cbSuccess 41 once the hardware finishes the RF op. CheckTagStatus results also arrive as cbSuccess 40/41/42, so a lock check armed right after a timed-out TID read can mis-read the stray 41 as "Locked". Rule: correlate QC callbacks to their command — expect/drain the stray ack only after a TID timeout (a successful TID read via cbTag needs no drain, and delaying it risks the tag leaving RF range); treat cbError 4 as "device busy, retry" not a status.
