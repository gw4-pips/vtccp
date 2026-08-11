# RFID Wedge Pro

Keyboard-wedge replacement for the BAIT RFID Wedge / MTI RFID ME tool.
Targets the **AsReader P35U Desktop UHF Reader**.
Works with **Excel**, **VTCCP**, **Notepad**, and any other Windows application.

---

## Features

| Feature | Details |
|---|---|
| Keyboard injection | Types EPC (or GTIN-14) into whatever window has focus |
| Deduplicate reads | Configurable cooldown — same tag only fires once per window |
| Prefix / Suffix | Prepend or append fixed text to every read |
| Append key | Nothing / Tab / Enter after each tag |
| Display Spaces | `AABBCC…` → `AA BB CC …` |
| Timestamp | Appends date-time to injected string, with configurable delimiter |
| UPC/EAN | Injects GTIN-14 instead of raw EPC (SGTIN tags only) |
| Include Filter | Prepends EPC filter digit |
| Power control | Set reader TX power in dBm |
| Live decode | Scheme, GTIN-14, EPC URI, RSSI, antenna, frequency for last read |
| Read log | Scrolling table of all reads in session |
| CSV export | Save session log as CSV |
| Auto file log | Appends every read to `TagLog.csv` next to the .exe |
| Settings persist | `rfid_wedge_config.json` saved on exit / Settings button |

---

## Hardware

**AsReader P35U** — USB Desktop UHF RFID Reader  
- USB VCP (Virtual COM Port) connection, standard COM port  
- TX power: 13–27 dBm (US/EU region)  
- SDK: `AsReaderP3xU.dll` v1.3.0 (ships with the .exe)  

Previous hardware (GoToTags E310) was retired — its VCP interface was non-functional.

---

## SDK Notes

The AsReader P35U uses a proprietary C# SDK (`AsReaderP3xU.dll`).  
Python calls it via **pythonnet** (`import clr`), which allows direct .NET DLL calls from Python.  
.NET Framework 4.x is pre-installed on Windows 10/11 — no extra runtime needed.

Key SDK calls used:
- `ConnectWithVCP(comPort)` — connect by COM port name  
- `StartInventory(rssiEn, maxTags, maxSecs, maxCycles, antenna1)` — begin continuous read  
- `StopInventory()` — halt reading  
- `SetRegion(REGION_US)` — required on init  
- `SetTxPower(dBm)` — set antenna power  
- `SetDelegate(...)` — register Python callbacks for tag data, errors, completion  

Tag data callback delivers: `epc`, `pc`, `tid`, `rssi`, `channel`, `antenna`, `phase`

---

## Building the .exe (run on Windows)

**Requirements:** Python 3.11 or later installed and on PATH.

```bat
cd tools\rfid-wedge
build_exe.bat
```

Output: `dist\RFIDWedgePro\RFIDWedgePro.exe`  
Copy the **entire** `dist\RFIDWedgePro\` folder — do not move the .exe alone.  
`AsReaderP3xU.dll` is copied automatically by the build script.

---

## Running from source (development)

```bat
pip install -r requirements.txt
python main.py
```

---

## Usage

1. Plug in the P35U reader via USB.
2. Launch `RFIDWedgePro.exe`.
3. Select the COM port from the dropdown (auto-detected).
4. Click **Connect**.
5. Configure formatting options as needed; click **Save Settings**.
6. Click **Set** to apply the power level to the reader.
7. Click **▶ Start Reading**.
8. Click on the Excel cell / VTCCP field you want to receive data.
9. Hold an RFID tag near the reader — the EPC (or GTIN-14) is typed into the focused field.
10. Click **■ Stop** when done.

---

## File layout

```
tools/rfid-wedge/
├── main.py              — UI (tkinter)
├── reader.py            — AsReader P35U driver (pythonnet → AsReaderP3xU.dll)
├── decoder.py           — EPC decode + GTIN extraction
├── injector.py          — keyboard injection (pynput)
├── config.py            — settings persistence (JSON)
├── requirements.txt     — pyserial, pynput, pythonnet
├── build_exe.bat        — PyInstaller build script (copies DLL automatically)
├── AsReaderP3xU.dll     — AsReader SDK DLL (place here before building)
└── README.md
```

---

## SDK Files (not committed — obtain from AsReader)

| File | Source |
|---|---|
| `AsReaderP3xU.dll` | `AsReader_P35U_SDK_c#_1_3_0` zip → root folder |
| Reference guide | `ASR-P3xU C# SDK Reference Guide V1.1` (PDF) |
| Sample app | `Sample_AsReader-P35U_c#_1_3_0` zip (C# WinForms demo) |

---

## Licence

Internal VCCS tool. Not for distribution.
