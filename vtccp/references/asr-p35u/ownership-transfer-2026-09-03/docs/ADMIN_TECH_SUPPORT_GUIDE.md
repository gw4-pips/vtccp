# RFID FlexWedge Pro — Administrator and Technical Support Guide

**Version:** 1.0  
**Validated baseline:** 03 September 2026  
**Transferred authority:** VTCCP Command Pilot

> This document supports the transferred Python reference implementation.
> VTCCP is now authoritative for FlexWedge RFID and VeriWedge RFID. Verify
> VTCCP-specific behavior before applying these instructions to a later native
> or integrated implementation.

## 1. Supported baseline

The transferred baseline was validated with:

- Windows 10
- 64-bit Python 3.11
- AsReader P35U Desktop UHF reader
- USB VCP connection
- `AsReaderP3xU.dll` from P35U C# SDK 1.3.0
- main firmware 1.8.0 on the development unit
- RFID module firmware `RED4S_v2.2.2_K_SD`

The fresh-machine test confirmed connection, SGTIN-96 inventory and decoding,
GTIN-14 and serial extraction, RSSI, antenna, lock state, separate TID read, and
TID update in the UI read log.

## 2. Component map

| Component | Responsibility |
|---|---|
| `main.py` | tkinter UI, presets, workflow, session log, QC orchestration |
| `reader.py` | pythonnet loading, P35U SDK delegates, connection and RFID operations |
| `decoder.py` | EPC decoding and output formatting |
| `injector.py` | Windows keyboard injection |
| `config.py` | defaults, JSON loading, persistence, migration |
| `requirements.txt` | Python dependencies |
| `build_exe.bat` | PyInstaller onedir build |
| `AsReaderP3xU.dll` | proprietary AsReader SDK dependency |

## 3. Clean source installation

Use 64-bit Python 3.11 for the validated support baseline.

```bat
cd C:\dev\RFID-FlexWedge-Pro
py -3.11 -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

Required Python packages:

- pythonnet
- pyserial
- pynput
- pystray
- Pillow

`AsReaderP3xU.dll` must be beside `main.py`.

## 4. Building the Windows application

Place the authorized DLL beside `build_exe.bat`, then run:

```bat
build_exe.bat
```

Output:

```text
dist\RFIDWedgePro\RFIDWedgePro.exe
```

Deploy the entire `dist\RFIDWedgePro\` folder. Never copy only the EXE. The DLL
and PyInstaller support files must remain with it.

The build is `--onedir`, not a single-file executable. This avoids repeatedly
extracting the SDK and makes dependency inspection easier.

## 5. DLL and .NET loading

The application locates the DLL beside the running script or packaged EXE and
loads it through:

```text
pythonnet → clr.AddReference("AsReaderP3xU") → AsReaderP3xU.AsReader
```

### Downloaded-file blocking

If .NET reports:

```text
An attempt was made to load an assembly from a network location
```

or:

```text
Operation is not supported
```

close the application and run from its folder:

```bat
powershell -NoProfile -Command "Unblock-File -LiteralPath '.\AsReaderP3xU.dll'"
```

For a packaged folder received as a ZIP, unblocking the ZIP before extraction
can prevent Windows from propagating the web-origin mark to its contents.

### Architecture

The validated configuration uses 64-bit Python. Use matching process and SDK
architecture. A load error before connection generally indicates blocking,
architecture, missing dependency, or wrong DLL—not a USB problem.

## 6. Hardware and connection architecture

The P35U appears as a Windows COM port, but the application does not open that
port through pyserial. `AsReaderP3xU.dll` owns the connection through
`ConnectWithVCP`.

Do not:

- open the P35U COM port with a terminal
- run raw serial diagnostics while FlexWedge is connected
- run the AsReader sample application simultaneously
- run two FlexWedge instances on the same reader

The validated unit enumerated with VID `0x339C` and PID `0x271B`. Confirm future
hardware revisions in Device Manager rather than treating those values as
universal.

Port discovery prioritizes known VID/PID, then descriptive keywords, then other
available COM ports. A listed COM port is not proof that it is the reader.

## 7. SDK initialization

The operational sequence is:

1. Instantiate `AsReader`.
2. Register all six delegates in one `SetDelegate` call.
3. Call `ConnectWithVCP`.
4. Configure the US region.
5. Apply power, timing, session, anti-collision, buzzer, filter, and threshold.
6. Start inventory when requested.

Registered callbacks cover:

- tag data
- SDK errors
- success/status codes
- command data
- read completion
- hardware trigger

`CallBackCommandData` is firmware-update-only. It never returns `ReadMemory`
results.

## 8. Connection lifecycle

Normal states include:

- Not connected
- Connecting
- Connected
- Reading
- Paused
- Armed
- Disarmed
- Disconnected

Disconnect should stop inventory before calling `DisConnect`. Unexpected SDK
errors must clear connected/reading state and return the UI to a safe state.

If a cable is pulled while idle, the SDK may not report it until another command
is attempted. Do not infer physical presence only from the last UI state.

## 9. Advanced Settings

### RF / Power

- **TX Power:** 13–27 dBm
- **RSSI Threshold:** −99 disables threshold filtering
- **Session:** S0, S1, S2, or S3
- **Anti-Collision:** DynamicQ or FixedQ
- **Start Q / Min Q / Max Q**

### Scan Timing

- read time per cycle
- idle time between cycles
- stop after first tag
- buzzer OFF, LOW, or HIGH
- Read TID after each scan
- Verify tag lock status after each scan

### Tag Filter

The packaging-level filter uses the three EPC filter bits at EPC bank offset
21. Supported selections are:

| Value | Meaning |
|---:|---|
| 0 | All others / unclassified |
| 1 | POS trade item / individual retail unit |
| 2 | Full case for transport / carton |
| 3 | Reserved |
| 4 | Inner pack trade item / shipper |
| 5 | Reserved |
| 6 | Unit load / pallet |
| 7 | Unit or component inside product |

Reserved values 3 and 5 are disabled in the UI. Multiple selected values are
applied as sequential Select commands with OR behavior where supported.

If the installed SDK does not expose the expected filter method, the driver
skips unsupported filtering rather than attempting unsafe reflection calls.

### Device

The Device tab reports available:

- main firmware
- hardware version
- RFID module firmware
- serial number
- SDK version
- COM port
- regional frequency-band summary

Per-read channel frequency is not exposed by the current SDK.

## 10. Configuration changes while reading

The SDK may reject setting changes during inventory. FlexWedge therefore:

1. stops inventory
2. waits briefly for stop completion
3. applies settings
4. restarts inventory when appropriate

Avoid rapidly applying settings while a TID or lock check is in progress.

## 11. Preset administration

Factory presets establish tested combinations of:

- transmit power
- buzzer level
- read and idle timing
- stop-after-first behavior
- Gen2 session
- anti-collision behavior
- dedupe interval
- TID and lock checks
- trigger mode

Factory-controlled settings are intentionally constrained while a factory
preset is active.

Two custom presets are stored in the configuration JSON. Custom preset capture
does not include every basic output-format field. When supporting a reported
preset problem, compare:

1. active preset name
2. advanced reader/QC fields
3. separately persisted output-format fields
4. startup-default preset

**Restore factory defaults** changes application configuration. A device
**Factory Reset** changes reader hardware settings. Do not confuse them.

## 12. TID technical behavior

TID is not delivered through `CallBackCommandData`.

The current driver attempts:

1. combined EPC/TID inventory if a safe supported SDK method exists; otherwise
2. `ReadMemory(MEM_TID, offset 0, length 4 words, password 0, EPC)`.

For the validated DLL, the result arrives asynchronously through
`CallBackReadTagData`, in the TID or data field of the inventory result.

The fallback workflow requires:

1. a target EPC
2. inventory stopped
3. the same physical tag still in range
4. a wait of up to approximately two seconds

The separate operation is the SDK's intended behavior, not a FlexWedge defect.

Support implications:

- An EPC success followed by empty TID is possible.
- Multiple tags can make target retention ambiguous.
- A moving or removed tag can time out.
- Do not “fix” this by waiting on `CallBackCommandData`.
- Do not call raw YRM100 commands while the SDK owns the device.

## 13. Lock-status technical behavior

`CheckTagStatus` runs asynchronously after EPC/TID processing. Observed success
codes map to:

- 40 — Permalock
- 41 — Lock
- 42 — Unlock

The support timeout is approximately three seconds. A timeout should return
control to the UI rather than freeze it.

Lock checking is unavailable in Trigger - Continuous mode because continuous
inventory conflicts with the required isolated QC sequence.

## 14. EPC decoding

Recognized header families include:

- SGTIN
- SSCC
- SGLN
- GRAI
- GIAI
- GID
- GDTI
- GSRN
- GSRNP
- CPI

GTIN-14 and item serial extraction apply to SGTIN where the required partition
data is valid. Other schemes can still identify their EPC scheme and URI but do
not necessarily produce a GTIN.

Decoder responses such as **Unknown**, **Too short**, or **Invalid hex** indicate
input/decode limitations, not reader transport errors.

The reference application displays EPC URI. It does not itself generate or
resolve the complete GS1 Digital Link verification URL for every EPC scheme.

## 15. Output injection troubleshooting

Keyboard injection uses the currently focused Windows control.

If no output is injected:

1. Confirm a read appears in FlexWedge.
2. Confirm the destination accepts keyboard input.
3. Check whether another window stole focus.
4. Test with Notepad.
5. Disable prefix, suffix, and translation to test raw HEX.
6. Check append-key behavior.

If Notepad works but the target application does not, investigate target-field
focus, browser security, application shortcut handling, or input timing rather
than RFID transport.

## 16. Runtime files

Files are created beside the script or packaged executable:

| File | Purpose |
|---|---|
| `rfid_wedge_config.json` | persisted application settings and custom presets |
| `TagLog.csv` | optional automatic read log |
| `debug.log` | SDK and diagnostic log |
| `debug.log.1` | previous rotated diagnostic log |

The debug log rotates at approximately 500 KB.

Do not collect or transfer logs without reviewing them for customer/tag data.
EPC and TID values can be operationally sensitive even when they are not user
credentials.

## 17. Automatic log versus Export CSV

The automatic `TagLog.csv` row is appended when the initial EPC event is
processed. A fallback TID can arrive afterward and update the in-memory UI row.

Consequently:

- the UI can show TID
- a later **Export CSV** can show TID
- the already-written automatic `TagLog.csv` row can retain a blank TID

This is expected in the transferred baseline and should be considered when
comparing logs.

## 18. Diagnostic procedure

Use this order to avoid conflating unrelated layers.

### Layer 1 — Application launch

- Does the window open?
- Is Connect enabled?
- Is pythonnet installed?
- Is the DLL present and unblocked?

### Layer 2 — Windows device

- Does the P35U appear in Device Manager?
- Which COM port is assigned?
- Does unplug/replug remove and restore that port?

### Layer 3 — SDK connection

- Is the correct COM port selected?
- Is another process using it?
- Does Connect return an SDK code?
- What does `debug.log` record?

### Layer 4 — Inventory

- Does Start return success?
- Is the region configured?
- Is power sufficient?
- Is RSSI threshold disabled for testing?
- Is an EPC filter excluding the test tag?

### Layer 5 — Decode and injection

- Does raw EPC appear?
- Does the scheme decode?
- Does Notepad receive raw HEX?
- Is translation applicable to this EPC scheme?

### Layer 6 — QC

- Is TID or lock checking enabled?
- Is only one tag present?
- Does the tag remain stationary?
- Is the workflow one-shot rather than continuous?

## 19. Common errors

### DLL not found

Place `AsReaderP3xU.dll` beside `main.py` or the packaged EXE. Do not rename it.

### Pythonnet not installed

Activate the intended virtual environment and run:

```bat
python -m pip install -r requirements.txt
```

Confirm the prompt begins with `(.venv)` and verify:

```bat
python --version
```

### Network-location assembly error

Unblock the DLL as described in Section 5.

### Connection failed with an SDK code

- verify COM port
- close competing programs
- reconnect USB
- restart the application
- inspect `debug.log`

### Start failed

The device may be disconnected, busy, or left in an incomplete prior command.
Disconnect cleanly, close the application, reconnect USB, and retry before
changing code.

### Connected but SDK information is unavailable

A wrong COM port may open without representing the P35U. Verify the port by
unplug/replug observation in Device Manager.

### No tags

- test one known-good tag
- remove filters
- use RSSI threshold −99
- use Standard at 20 dBm
- place tag close and correctly oriented
- verify the correct regulatory region

### TID unavailable

- enable TID
- use Ghost or Stealth
- keep one tag in the field
- retry close to the antenna
- review callback timing in `debug.log`

### Lock timeout

Keep the tag present and use a one-shot/QC-capable mode.

### Incorrect or missing translated identifier

Capture the raw EPC first. Confirm the EPC scheme supports the requested GTIN,
EAN, UPC, or serial representation. Preserve leading zeros.

## 20. Support-data collection

For a reproducible support case, collect:

- Windows version
- Python version and architecture, if running source
- application version
- SDK DLL filename and hash
- P35U COM port
- hardware, main firmware, RFID firmware, serial, and SDK versions
- selected preset
- advanced settings
- raw EPC
- expected and actual output
- whether one or multiple tags were present
- relevant `debug.log` excerpt
- exact steps and timestamps

Do not request customer credentials or unrelated personal data.

## 21. Known baseline limitations

- AsReader officially validates the SDK for C#/.NET, not pythonnet.
- TID requires a separate asynchronous request on the validated DLL.
- TID and lock checks depend on the tag remaining in the RF field.
- Lock status is unavailable in Trigger - Continuous.
- Per-read RF frequency is unavailable.
- Automatic CSV can precede fallback TID completion.
- The reference application does not implement the complete Command Pilot
  WebScan/DataMan/RFID correlation workflow.
- VeriWedge RFID and FlexWedge RFID may intentionally diverge.
- VTCCP changes after transfer supersede this document where implementations
  differ.

## 22. Escalation boundary

Escalate to the VTCCP owner when:

- behavior differs from this transferred Python baseline
- the issue concerns WebScan or DataMan event coordination
- barcode/RFID correlation or verification fails
- FlexWedge and VeriWedge product requirements conflict
- a native VTCCP implementation differs from pythonnet behavior

Escalate to AsReader when:

- the SDK or firmware returns undocumented codes
- a supported C# sample reproduces the hardware problem
- firmware update or recovery is required
- SDK licensing or redistribution terms need clarification
