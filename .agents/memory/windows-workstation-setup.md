---
name: Windows workstation setup
description: Steps to set up a fresh Windows machine to build and run VtccpApp from GitHub.
---

# Windows workstation setup

## Clone path convention
- `C:\dev\vtccp` — repo root (workspace). Run `git pull` here.
- `C:\dev\vtccp\vtccp` — C# source root (where VTCCP.sln lives). Run `dotnet build` here.
- The workspace repo (`gw4-pips/vtccp`) contains the full Replit workspace at its root; the C# project is the `vtccp/` subdirectory.

## One-time setup commands
```powershell
git clone https://github.com/gw4-pips/vtccp.git C:\dev\vtccp
New-Item -ItemType Directory -Path "C:\dev\vtccp\vtccp\lib\asreader-p3xu-sdk-1.3.0" -Force
copy "<SDK zip location>\AsReaderP3xU.dll" "C:\dev\vtccp\vtccp\lib\asreader-p3xu-sdk-1.3.0\AsReaderP3xU.dll"
cd C:\dev\vtccp\vtccp
dotnet build VtccpWindows.sln -c Release
```

## AsReader DLL location (lab network)
`Q:\VendorDOC\AsReader (Asterisk)\AsReader_P35U_SDK_cs_1_3_0\AsReader_P35U_SDK_c#_1_3_0\AsReaderP3xU.dll`
Place at: `C:\dev\vtccp\vtccp\lib\asreader-p3xu-sdk-1.3.0\AsReaderP3xU.dll`

**Why:** The DLL is not committed to the repo (vendor binary). The Windows projects resolve `..\lib` from the inner C# source root, not from the outer Git workspace root. Without it, `EpcReaderFactory` is excluded from DeviceInterface compilation and the `ASREADER_SDK` compile symbol is not defined in VtccpApp — RFID scanning is unavailable at runtime but the build succeeds cleanly.

## Day-to-day update
Double-click `C:\dev\vtccp\vtccp\tools\update-and-build.bat` — pulls from GitHub and rebuilds.

## Known harmless warnings
12 × NU1701 (PDFsharp-GDI / PDFsharp-MigraDoc-GDI targeting .NET Framework instead of net8.0). These are permanent and do not affect runtime.

## Windows-only build notes
Task agents run on Linux (Replit) and cannot compile the WPF project. Windows-specific errors to watch for:
- CS0103 `Path`/`File` not found → missing `using System.IO;` (wpftmp project doesn't get implicit SDK global usings)
- CS0103 `EpcReaderFactory` not found → AsReader DLL absent at compile time (expected; RFID unavailable)
- CS8999 raw string literal whitespace → closing `"""` indentation > minimum content indentation
- CS1503 wrong argument type → task agent calling wrong overload of a method
