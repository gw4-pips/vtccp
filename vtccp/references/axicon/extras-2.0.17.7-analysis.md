# Axicon Extras 2.0.17.7 — Static Analysis

**Document version:** 1.0  
**Revision date:** 5 September 2026

## Scope

Static, non-executing inspection of:

- `Extras-2.0.17.7.zip`
- Loose `c39ascii`, `c39pharm`, `cip39`, `EuroCoupon`, `GenSpec`, `HIBC`,
  `InStore`, and `isbn` `.avp` files

The archive was extracted under `/tmp` for inspection. No legacy executable or
plug-in was run, installed, or copied into the application.

## What the `.avp` files are

The `.avp` files are native 32-bit Windows PE DLLs, not text configuration or
saved-scan data. The package README calls them optional **data content analysis
plug-ins** for Axicon Verifier. Axicon installed them by copying selected files
beside `Verifier.exe`.

The loose plug-ins supplied separately are byte-for-byte SHA-256 matches for
the same named files in the ZIP. They are not newer variants.

Examples of the optional analyzers include:

- Code 39 ASCII
- Code 39 pharmaceutical formats
- CIP 39
- European and US coupon formats
- General specification checks
- HIBC
- In-store data
- ISBN
- AIAG, PZN, SISAC, UPC, UK Coupon, User Data, and other legacy schemes

The archive also has an `obsolete` plug-in directory containing superseded
EAN-128/CIP-128 components. Those should not be treated as current standards
implementations.

## Archive contents

The package contains three useful groups:

### Plugins

Optional native verifier analysis modules and several language-specific
resource DLLs. These appear intended to extend interpretation and application
checks inside the old Axicon Verifier software.

### System

Legacy Windows runtime/update files, including MFC and C runtime components.
These are from the Windows 9x/early-XP era and should not be installed on a
production workstation without a controlled compatibility assessment.

### Utils

The most relevant utilities are:

- **PluginStub.exe** — Axicon's stand-alone host for testing and using verifier
  plug-ins outside the main Verifier application.
- **SetupScanDB.exe** — installer for the “Scan File Data Extractor.”
- **ScanFix.exe** — repairs damaged saved scan files.
- **Console.exe** — text-based interface to the PC Verifier software.
- **DecodeConfig.exe** — low-level decoder configuration utility.
- **Online.exe** — contains Borland Database Engine references and may be tied
  to a legacy local database workflow.
- **Translator.exe** — purpose needs controlled runtime inspection.

The README explicitly says `SetupScanDB.exe` installs a Scan File Data
Extractor. That is more relevant to VeriWedge integration than the content
analysis plug-ins themselves because it may expose the structure of Axicon
saved-scan files or provide a supported export path.

## Integration implications

The plug-ins do **not** presently look like a direct live-result API for
VeriWedge. They are application-standard/data-content analyzers loaded by the
legacy Axicon verifier.

The strongest integration leads in this package are:

1. Determine what the Scan File Data Extractor reads and emits.
2. Determine whether `Console.exe` can initiate scans or emit machine-readable
   results.
3. Use `PluginStub.exe` to learn the old plug-in contract only if application
   data checks from these legacy modules are still required.
4. Keep the preferred VeriWedge boundary at exported/saved verifier evidence
   rather than loading untrusted 32-bit plug-ins into the .NET 8 process.

## Safety and compatibility

Most archive timestamps are from 2003–2004. The binaries are 32-bit native
Windows programs built with legacy Delphi/C++Builder/MFC-era dependencies.

Do not install or execute them directly on the normal development or verifier
workstation as an exploratory first step. Use an isolated Windows virtual
machine or disposable test system, record filesystem/registry changes, and
scan the binaries before execution.

## Recommended controlled follow-up

On an isolated 32-bit-compatible Windows environment:

1. Snapshot the system.
2. Run and monitor `SetupScanDB.exe`.
3. Record installed files, registry keys, file associations, and supported
   input/output formats.
4. Test the extractor against a copy of a real Axicon saved scan.
5. Inspect `Console.exe` help and output without connecting production
   hardware first.
6. Run `PluginStub.exe` with one non-obsolete plug-in and record its required
   inputs and outputs.
7. Restore the snapshot after testing.

The decisive missing artifact remains a current Axicon saved-scan/export file
from the actual verifier installation.
