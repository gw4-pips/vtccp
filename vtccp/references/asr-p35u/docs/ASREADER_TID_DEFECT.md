# AsReader P35U — TID Read — Investigation & Resolution

**Product:** AsReader P35U UHF RFID Reader  
**Interface:** VCP (Virtual COM Port) via USB  
**SDK:** AsReader C# DLL, loaded via pythonnet (Python ↔ .NET interop)  
**Platform:** Windows 10/11, Python 3.x, pythonnet  
**Opened:** 2026-08-08  
**Resolved:** 2026-08-17 — vendor confirmed expected SDK behaviour  
**Status:** ✅ RESOLVED — working correctly as of firmware 1.8.0 + FlexWedge Path A fix

---

## Summary

`CallBackCommandData` was expected to deliver `ReadMemory()` results but never fired.
**Root cause: by design.** `CallBackCommandData` is reserved exclusively for firmware
update packets. `ReadMemory()` results are delivered via `CallBackReadTagData` — the
same callback used for inventory tag reads.

Our Path A fix (intercepting `_cb_tag` for ReadMemory results) is the correct and
officially documented behaviour. TID reads are now working on firmware 1.8.0.

---

## Vendor Response (AsReader Engineering, Japan — 2026-08-17)

> **Q1. Why does CallBackCommandData never fire after ReadMemory()?**
> This is by design, not a bug. CallBackCommandData is reserved exclusively for
> firmware update command/response packets (address request, file transfer, transfer
> complete, device reboot, RFID module firmware update). It is never invoked for
> ReadMemory() results, regardless of mode, firmware version, or configuration —
> there is no prerequisite setting that would make it fire for memory reads.
>
> **Q4. Is there an alternative SDK method to read TID memory bank data?**
> Yes. Data read via ReadMemory() (including the TID memory bank) is delivered
> through the CallBackReadTagData callback, as part of the InventoryResult object
> (the tag data field corresponding to TID). This is delivered while an inventory
> session is running — it is not a separate synchronous call, but the existing
> tag-report channel you are already using.
>
> Since you have already confirmed that CallBackReadTagData fires correctly in your
> environment, we would like to ask you to check the TID value there, instead of
> waiting on CallBackCommandData. We expect this will resolve the issue you're seeing.
>
> **Q2/Q3. Correct C# signature for SetHIDInventoryMode() / GetHIDWorkParams()?**
> Our SDK officially supports and is verified for C#/.NET environments only. We are
> not able to guarantee behaviour when the SDK is accessed via pythonnet, since type
> marshalling between Python and .NET can behave differently than in native C#.
>
> **Q5/Q6. Minimum firmware version / SendCommand for raw YRM100 commands?**
> Once TID is read via CallBackReadTagData as described above, these should no longer
> be necessary.

---

## Original Investigation

### What we observed
`ReadMemory()` returns `0` (accepted), but `CallBackCommandData` was never invoked.
The call timed out silently after 2 seconds. No error code or exception was raised.

### What we tried
1. Confirmed delegate registration was correct (all six callbacks registered)
2. Confirmed ReadMemory was accepted by hardware (returns 0)
3. Tested raw YRM100 packet via `SendCommand` — also never triggered `CallBackCommandData`, and left device busy for ~5 s
4. Confirmed `SetInventoryType` not present on this DLL
5. Attempted `SetHIDInventoryMode` — pythonnet binding crash before Python could catch it

### What fixed it
Added a `_pending_memory_cb` one-shot hook in `_cb_tag` (our `CallBackReadTagData`
handler). When `read_tid()` sets this hook before calling `ReadMemory()`, the next
`_cb_tag` invocation is intercepted, `tagdata.tid` (or `tagdata.data`) is read, and
the result is returned. This is now confirmed correct per vendor.

---

## Appendix A — DLL Public Members (SDK 1.3.0)

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

### CallBackCommandData — correct usage (firmware update only)

This callback is used internally by the SDK during firmware updates for these
packet types only:
- `AsReaderP3xUFirmwareTypeAddress` (0x58) — firmware address request
- `AsReaderP3xUFirmwareTypeTransferFile` (0x59) — firmware file transfer

Never register `CallBackCommandData` expecting ReadMemory or other command results.
