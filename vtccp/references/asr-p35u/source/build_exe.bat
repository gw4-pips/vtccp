@echo off
REM ─────────────────────────────────────────────────────────────────────────
REM  RFID Wedge Pro — Windows .exe builder
REM  Run this on your Windows machine (Python 3.11+ required).
REM  Output: dist\RFIDWedgePro\RFIDWedgePro.exe
REM
REM  Before building, copy AsReaderP3xU.dll into this folder.
REM  (Extract it from AsReader_P35U_SDK_cs_1_3_0.zip.)
REM ─────────────────────────────────────────────────────────────────────────

echo Checking for AsReaderP3xU.dll...
if not exist "AsReaderP3xU.dll" (
    echo.
    echo  ERROR: AsReaderP3xU.dll not found in the current folder.
    echo  Copy it from the SDK zip ^(AsReader_P35U_SDK_cs_1_3_0.zip^) first.
    echo.
    pause
    exit /b 1
)

echo Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building executable...
pyinstaller ^
  --onedir ^
  --windowed ^
  --name "RFIDWedgePro" ^
  --icon NONE ^
  --add-data "." ^
  --add-binary "AsReaderP3xU.dll;." ^
  --hidden-import serial ^
  --hidden-import serial.tools ^
  --hidden-import serial.tools.list_ports ^
  --hidden-import pynput.keyboard._win32 ^
  --hidden-import pynput.mouse._win32 ^
  --hidden-import clr ^
  --hidden-import pythonnet ^
  main.py

echo.
echo Copying AsReaderP3xU.dll to dist folder...
xcopy /Y "AsReaderP3xU.dll" "dist\RFIDWedgePro\"

echo.
echo ─────────────────────────────────────────────────────────────────────────
echo  Done.  Executable is at:  dist\RFIDWedgePro\RFIDWedgePro.exe
echo  Copy the entire dist\RFIDWedgePro\ folder to run on another machine.
echo  AsReaderP3xU.dll must remain in the same folder as the .exe.
echo  Config is saved as rfid_wedge_config.json next to the .exe.
echo  Tag log is written as TagLog.csv next to the .exe.
echo ─────────────────────────────────────────────────────────────────────────
pause
