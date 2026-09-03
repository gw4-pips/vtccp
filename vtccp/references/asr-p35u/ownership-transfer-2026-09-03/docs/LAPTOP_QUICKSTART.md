# RFID FlexWedge Pro — Laptop Quick Start

## Required environment

- Windows 10 or 11
- AsReader P35U connected by USB
- Python 3.11 (64-bit recommended)
- The authorized `AsReaderP3xU.dll` from AsReader's P35U C# SDK 1.3.0

Python 3.11 is recommended for this hardware test because it is the version used
for the packaged build path and has established compatibility with pythonnet.

## Prepare the folder

1. Extract the supplied FlexWedge source ZIP to a normal local folder.
2. Copy `AsReaderP3xU.dll` into that extracted folder, beside `main.py`.
3. Do not launch the AsReader sample application at the same time; only one
   application should control the reader's COM port.

## Create the Python environment

Open Command Prompt in the extracted folder and run:

```bat
py -3.11 -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If `py -3.11` reports that Python 3.11 is unavailable, install the 64-bit
Python 3.11 release from python.org with the Python Launcher enabled.

## Connect and run

1. Connect the P35U by USB and wait for Windows to finish recognizing it.
2. Open Device Manager → Ports (COM & LPT) and note the P35U COM port.
3. In the activated Command Prompt, run:

```bat
python main.py
```

4. Select the P35U COM port.
5. Click **Connect**.
6. Place only one sample tag near the reader for the first test.
7. Start reading and confirm that EPC, decode information, and RSSI appear.
8. Run the TID/quality check while keeping the same tag in the field.

## Stop and restart

Close the application normally before unplugging the reader. To run it again:

```bat
cd path\to\the\extracted\folder
.venv\Scripts\activate
python main.py
```

## Files created during use

- `rfid_wedge_config.json` — saved settings
- `TagLog.csv` — automatic read log
- `debug.log` — diagnostic output, when enabled by the application

These are local runtime files and should not be returned as source code.

## If connection fails

Check these items before changing code:

1. `AsReaderP3xU.dll` is beside `main.py`.
2. The selected COM port matches Device Manager.
3. No AsReader sample program or second FlexWedge instance is open.
4. Python and the SDK DLL use compatible architecture; use 64-bit Python 3.11
   for this package.
5. Disconnect/reconnect the USB cable, reopen the application, and try once.
