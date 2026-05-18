# GS1 Syntax Engine — Version Pin

| Item | Value |
|---|---|
| **Pinned version** | 1.4.0 |
| **Published** | 2026-05-18 |
| **Downloaded** | 2026-05-18 |
| **Next check due** | 2026-06-18 |
| **Source** | https://github.com/gs1/gs1-syntax-engine/releases |
| **License** | Apache 2.0 |

## Files included

| Path | Description |
|---|---|
| `src/GS1Encoder.cs` | C# .NET wrapper class (P/Invoke — the file to reference from VTCCP) |
| `src/GS1EncoderTest.cs` | Official test suite for the C# wrapper |
| `src/gs1encoders-dotnet-lib.csproj` | Project file for the dotnet-lib package |
| `src/README.md` | Dotnet binding integration notes |
| `dotnet-lib-release/gs1encoders-dotnet.dll` | Pre-built managed .NET wrapper DLL |
| `dotnet-lib-release/runtimes/win-x64/native/gs1encoders.dll` | Pre-built native x64 DLL (bundled in dotnet release) |
| `dotnet-lib-release/runtimes/win-x86/native/gs1encoders.dll` | Pre-built native x86 DLL (bundled in dotnet release) |
| `native/x64/gs1encoders.dll` | Native x64 DLL (standalone libs release) |
| `native/x64/gs1encoders.h` | C header for native library |
| `native/x64/gs1encoders.lib` | MSVC import library for x64 |
| `native/x86/gs1encoders.dll` | Native x86 DLL (standalone libs release) |
| `native/x86/gs1encoders.h` | C header for native library (x86) |
| `native/x86/gs1encoders.lib` | MSVC import library for x86 |
| `LICENSE` | Apache 2.0 license text |
| `README.md` | Top-level engine README |

Also: `vtccp/lib/gs1-syntax-dictionary/gs1-syntax-dictionary.txt` (344 lines) —
the AI rules dictionary that the engine loads at startup.

## Integration path into VTCCP (when E1 is scheduled)

For a WPF/C# project:

1. **For build-time source integration** (preferred — no extra DLL management):
   - Add `src/GS1Encoder.cs` directly to the `DeviceInterface` or a new
     `Validation` C# project (do not reference the csproj — just include the .cs).
   - Copy `dotnet-lib-release/gs1encoders-dotnet.dll` +
     `dotnet-lib-release/runtimes/win-x64/native/gs1encoders.dll` (or x86)
     into the VTCCP output directory (set `Copy to Output Directory = Always`).
   - `GS1Encoder.cs` uses P/Invoke to load `gs1encoders.dll` by filename —
     it must be in the same directory as the `.exe` at runtime.

2. **Syntax dictionary**: copy `gs1-syntax-dictionary.txt` to the output
   directory alongside the DLL (or the engine falls back to its built-in
   static table — either works, but the file gives you the latest AI rules).

3. **Quick smoke test** (from `GS1EncoderTest.cs`):
   ```csharp
   using (var gs1 = new GS1Encoders.GS1Encoder()) {
       gs1.ScanData = "]d1" + "[)>\x1E06\x1D18VD89536\x1D1P8902A\x1DS3122A02965\x1E\x04";
       Console.WriteLine(gs1.HasValidScanData ? "PASS" : "FAIL: " + gs1.ErrMarkup);
   }
   ```

## Monthly update check procedure

Check for new releases at: https://github.com/gs1/gs1-syntax-engine/releases

Steps:
1. Compare the latest release tag against `1.4.0` (current pin above).
2. If newer: read the release notes. If the C# API surface changed,
   update `src/GS1Encoder.cs` and check for breaking changes against any
   VTCCP code that calls it.
3. Download the new `gs1encoders-windows-dotnet-lib.zip` and
   `gs1encoders-windows-libs-x64.zip` / `x86.zip` from the Releases page.
4. Replace the files in this directory tree.
5. Also re-download `gs1-syntax-dictionary.txt` from:
   https://raw.githubusercontent.com/gs1/gs1-syntax-dictionary/main/gs1-syntax-dictionary.txt
6. Update the pin table at the top of this file (version, dates, next-check).

Also check the syntax dictionary separately — it updates independently of the engine:
https://github.com/gs1/gs1-syntax-dictionary/blob/main/CHANGES

**Note for the LLM assistant**: this update check is a standing instruction.
At the start of any session where the date is on or after the "Next check due"
date in the table above, flag to the user that an update check is due for
`gs1/gs1-syntax-engine` before proceeding with other work.
