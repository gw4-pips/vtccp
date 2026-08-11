# AsReader P35U — TID Read Defect Report

**Product:** AsReader P35U UHF RFID Reader  
**Interface:** VCP (Virtual COM Port) via USB  
**SDK:** AsReader C# DLL, loaded via pythonnet (Python ↔ .NET interop)  
**Platform:** Windows 10/11, Python 3.x, pythonnet  
**Date:** 2026-08-08  
**Status:** Unresolved — submitted for vendor support

---

## Summary

The `CallBackCommandData` delegate registered via `SetDelegate()` **never fires** in response to `ReadMemory()` calls. As a result, TID memory bank data cannot be retrieved from Gen2 tags using the provided SDK. All other callbacks (`CallBackReadTagData`, `CallBackSuccessCode`, `CallBackErrorCode`, `CallBackReadComplete`, `CallBackTriggerHandler`) fire correctly.

---

## Environment

| Item | Detail |
|---|---|
| Reader model | AsReader P35U |
| Connection | USB → VCP (COM4), 115200 baud |
| SDK language | C# DLL loaded via pythonnet |
| Python version | 3.x (Windows) |
| Available DLL methods | See Appendix A |

---

## Expected Behaviour

Per SDK documentation, calling `ReadMemory(MemBankType.MEM_TID, 0, 4, 0, epcBytes)` should:

1. Transmit a Gen2 Read command to the tag targeting the TID memory bank  
2. Deliver the response asynchronously via the `CallBackCommandData` delegate  
3. The delegate receives a `byte[]` containing the raw TID bytes

---

## Observed Behaviour

`ReadMemory()` returns `0` (accepted), but `CallBackCommandData` is **never invoked**. The call times out silently after 2 seconds. No error code or exception is raised. The hardware is confirmed busy for approximately 2 seconds after the call (subsequent `CheckTagStatus()` succeeds immediately after the timeout, confirming the hardware does complete the RF operation — the result simply never reaches the callback).

---

## Investigation Steps Taken

### 1. Confirmed delegate registration is correct

All six delegates are registered via `SetDelegate()` in a single call:

```python
dev.SetDelegate(
    AsReaderCls.CallBackReadTagData(cb_tag),       # fires correctly ✓
    AsReaderCls.CallBackErrorCode(cb_error),        # fires correctly ✓
    AsReaderCls.CallBackSuccessCode(cb_success),    # fires correctly ✓
    AsReaderCls.CallBackCommandData(cb_command),    # NEVER fires ✗
    AsReaderCls.CallBackReadComplete(cb_complete),  # fires correctly ✓
    AsReaderCls.CallBackTriggerHandler(cb_trigger), # fires correctly ✓
)
```

`CallBackCommandData` is listed as a valid delegate type in `dir(device)`.

### 2. Confirmed ReadMemory is accepted by hardware

`ReadMemory(MEM_TID, 0, 4, 0, epcBytes)` returns `0`. The hardware executes the RF command (it is busy for ~2 s). `DLL success callback: 41` fires after inventory completion. The RF field remains active — `CheckTagStatus()` succeeds on the same tag immediately after the timeout.

### 3. Tested alternative: SendCommand with raw YRM100 packet

Constructed a raw YRM100 Read Memory packet (command `0x39`, TID bank `0x02`, 4 words) and called `SendCommand(packet)`. The DLL returns `True` (accepted) and the hardware executes the command (device is busy for ~5 s). `CallBackCommandData` still **never fires**. Additionally, `CheckTagStatus()` returns `Error (raw=4)` while hardware is busy, indicating a device conflict — this approach was abandoned.

### 4. Confirmed SetInventoryType does not exist on this DLL

`SetInventoryType` is not present in `dir(device)`. The `InventoryType` enum exists in the Types namespace with members including `PC_EPC_TID`, `PC_EPC_RSSI`, and `ONLY_PC_EPC` — but no SDK method exposes it.

### 5. SetHIDInventoryMode — unusable without correct type signature

`SetHIDInventoryMode` is present in `dir(device)`. However, any attempt to call it from Python/pythonnet with plausible argument types causes `System.ArgumentException: We should never receive instances of other managed types` in pythonnet's `MethodBinder` at the binding layer — **before Python can catch the exception**, terminating the process. The correct C# parameter type is unknown and is not documented.

### 6. GetHIDWorkParams — unusable without correct argument type

`GetHIDWorkParams` requires at least one argument (calling with zero arguments raises `No method matches given arguments`). The expected argument type is unknown. Probing with primitive Python types (`None`, `0`, `True`, `b''`) causes the same pythonnet binding crash as above.

---

## Questions for ASReader Tech Support

1. **Why does `CallBackCommandData` never fire after `ReadMemory()`?** Is there a prerequisite call, mode setting, or firmware version requirement for this callback to be delivered?

2. **What is the correct C# signature for `SetHIDInventoryMode()`?** Specifically, what enum or type should be passed as its argument?

3. **What is the correct C# signature for `GetHIDWorkParams()`?** What argument type does it require?

4. **Is there an alternative SDK method to read TID memory bank data** that does not rely on `CallBackCommandData`? For example, a synchronous overload, a different callback, or a specific inventory mode that includes TID in `CallBackReadTagData` results?

5. **Is there a minimum firmware version** required for `ReadMemory` callback support? Can you provide the version that introduced working `CallBackCommandData` delivery?

6. **Is `SendCommand` the intended path for raw YRM100 commands?** If so, why does the hardware remain busy for ~5 s (vs ~2 s for `ReadMemory`), and does `CallBackCommandData` apply to `SendCommand` responses?

---

## Appendix A — DLL Public Members

```
['CallBackCommandData', 'CallBackErrorCode', 'CallBackReadComplete',
 'CallBackReadTagData', 'CallBackSuccessCode', 'CallBackTriggerHandler',
 'CheckTagStatus', 'ConnectWithVCP', 'DefaultSetting', 'DisConnect',
 'Equals', 'Finalize', 'GetAntiCollisionMode', 'GetBasicTarget',
 'GetBuzzer', 'GetChannel', 'GetFH_LBT', 'GetFrequencyAutomatic',
 'GetFwVersion', 'GetHIDInventoryMode', 'GetHIDWorkParams', 'GetHashCode',
 'GetHwVersion', 'GetIdelTime', 'GetProductSN', 'GetQuery',
 'GetRFIDFwVersion', 'GetRSSIThreshold', 'GetReadTime', 'GetRegion',
 'GetSdkVersion', 'GetSelectMask', 'GetSelectionEnable', 'GetSession',
 'GetTxPower', 'GetType', 'InventoryResult', 'Kill', 'LockMemory',
 'MemberwiseClone', 'Overloads', 'ReadMemory', 'ReferenceEquals',
 'SendCommand', 'SetAntiCollisionMode', 'SetBasicTarget', 'SetBuzzer',
 'SetChannel', 'SetDelegate', 'SetFH_LBT', 'SetFrequencyAutomatic',
 'SetHIDInventoryMode', 'SetHIDWorkParams', 'SetIdelTime', 'SetQuery',
 'SetRSSIThreshold', 'SetReadTime', 'SetRegion', 'SetSelectMask',
 'SetSelectionEnable', 'SetSession', 'SetTxPower', 'StartInventory',
 'StopInventory', 'TagAction', 'TagData', 'TagMask', 'ToString',
 'WriteMemory']
```

---

## Appendix B — Minimal Reproduction (Python / pythonnet)

```python
import clr
clr.AddReference('AsReader')
from AsReader import AsReader as AsReaderCls, Types

dev = AsReaderCls()

fired = []

def cb_command(data):
    fired.append(bytes(data) if data else b'')

dev.SetDelegate(
    AsReaderCls.CallBackReadTagData(lambda r: None),
    AsReaderCls.CallBackErrorCode(lambda c: None),
    AsReaderCls.CallBackSuccessCode(lambda c: None),
    AsReaderCls.CallBackCommandData(cb_command),      # never fires
    AsReaderCls.CallBackReadComplete(lambda c, i: None),
    AsReaderCls.CallBackTriggerHandler(lambda t: None),
)

dev.ConnectWithVCP('COM4')
dev.SetRegion(Types.RegionType.REGION_US)

# Scan a tag, then call:
epc = bytes.fromhex('30342BF92851DD10F36A0483')
ret = dev.ReadMemory(Types.MemBankType.MEM_TID, 0, 4, 0, epc)
print(f'ReadMemory returned: {ret}')   # prints 0

import time; time.sleep(3)
print(f'cb_command fired: {len(fired)} times')   # prints 0
print(f'TID data: {fired}')                       # prints []
```

**Expected:** `cb_command fired: 1 times`, `TID data: [b'\xe2\x00...']`  
**Actual:** `cb_command fired: 0 times`, `TID data: []`
