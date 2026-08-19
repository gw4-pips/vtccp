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
| CallBackSuccessCode | **uint** code | Do not use for `CheckTagStatus`; a delayed 41 can acknowledge a timed-out TID read |
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

## CheckTagStatus direct-return contract (confirmed)

`CheckTagStatus(epc)` returns the tag lock status directly; it is not a command-accepted result that must await a status callback:

| Return | Meaning |
|---:|---|
| 0 | Unlocked |
| 1 | Locked |
| 2 | Permalocked |
| 3 | Unknown |
| 4 | Error |

The standalone RFID Wedge reference maps these direct return values and reports the known test tag as Permalocked. A VTCCP Windows trace showed `CheckTagStatus returned 2`, proving the app discarded the correct Permalocked status by treating every non-zero return as a rejected command.

**Why:** A timed-out TID `ReadMemory` can still emit a delayed `cbSuccess 41`; that callback is a stale acknowledgment, not the subsequent tag-lock result. The prior callback-correlator assumption conflated the two SDK behaviors.

**How to apply:** Map `CheckTagStatus`'s direct return on a worker task (the SDK can block while the tag leaves the RF field), with a timeout that preserves `Unknown` only for a genuine timeout/error. Do not await `cbSuccess 40/41/42` for this operation.

## Known permanent-lock discrepancy

A known test tag reports **Permalocked** through the standalone RFID Wedge decoder, while VTCCP reported **Unknown** for the same tag. The Windows trace resolved the discrepancy: VTCCP received direct return `2` (Permalocked) and incorrectly treated it as a rejected command.

**Why:** The method contract had been implemented as asynchronous callback delivery, but the working standalone reference demonstrates that its direct return is the lock-status enum. The TID read's delayed `cbSuccess 41` added misleading evidence.

**How to apply:** Replace the callback-correlator lock-status path with a direct return mapping, then test the known tag on the Windows workstation. The expected PDF value is **Permalocked**.

## Application-exit disconnect race (SDK 1.3.0)

**Rule:** On final application exit, stop inventory but do not call the vendor SDK's `DisConnect()` method; use the dedicated shutdown path and let Windows release the VCP handle when the process terminates. Keep explicit disconnect for an in-app disconnect/reconnect.

**Why:** The SDK can clear `RcpProtocolHandler.RxRspParsed` while its receive worker is still dispatching a response, producing a `NullReferenceException` inside `AsReaderP3xU.dll` during application close.

**How to apply:** Only use the shutdown-only path from the application exit handler. Do not substitute it for the normal Disconnect button, which must release the reader for an in-process reconnect.
