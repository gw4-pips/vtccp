# GS1 Barcode Syntax Resource — Catalog Entry

**Date cataloged**: 2026-05-18
**Status**: DOWNLOADED AND INCLUDED — `vtccp/lib/gs1-syntax-engine/` (v1.4.0, 2026-05-18)
**Next update check due**: 2026-06-18
**Triggered by**: v1.24 probe — `DebugGS1` returned all-undefined, confirming
DMST's `r.validation.gs1` is NOT an AI-property bag. Application-syntax
validation for GS1 / MIL-STD / ISO 15434 must live in VTCCP itself, not
delegated to the device.

---

## The Suite (three repos, one publisher)

All three are published under the **GS1 GitHub organization**:
<https://github.com/gs1>

All are licensed **Apache 2.0** — permissive, compatible with proprietary
closed-source distribution (VTCCP ships as a VCCS internal tool, so this
is fine).

---

### 1. gs1-syntax-dictionary
<https://github.com/gs1/gs1-syntax-dictionary>

A single machine-readable text file (`gs1-syntax-dictionary.txt`) that is
the authoritative source of truth for every assigned GS1 Application
Identifier: format spec, allowed character sets, length constraints,
mutually-exclusive AI pairs, requisite AI pairs, and Digital Link key
definitions.

Examples of what each entry encodes:
- `01` → `N14,csum,gcppos2` | `ex=255,37 dlpkey=22,10,21|235` → GTIN
- `10` → `X..20` | `req=01,02,03,8006,8026` → BATCH/LOT
- `11` → `N6,yymmd0` | `req=01,02,03,8006,8026` → PROD DATE

The gs1-syntax-engine (below) loads this file at startup or embeds a
compiled-in static table derived from it.

### 2. gs1-syntax-engine ← **THE ONE TO USE IN VTCCP**
<https://github.com/gs1/gs1-syntax-engine>
Docs: <https://gs1.github.io/gs1-syntax-engine/>

A native C library implementing the full GS1 AI validation stack, with
official language bindings including **C# .NET (P/Invoke wrapper)**.

**Input formats it accepts:**

| Format | Example |
|---|---|
| Bracketed AI element string | `(01)09780345418951(10)ABC123` |
| Unbracketed AI element string | `0109780345418951` + GS + `10ABC123` |
| Scan data (AIM ID prefix) | `]d10109780345418951` + GS + `10ABC123` |
| GS1 Digital Link URI | `https://example.com/01/09780345418951` |

**For VTCCP** the relevant input format is **scan data**, because our
`DecodedData` arrives already prefixed with the AIM ID (`]d1` for Data
Matrix GS1). The engine's `setScanData()` call accepts exactly this.

**What it validates:**
- AI format constraints (field type, length, check digits)
- Mutually-exclusive AI combinations
- Requisite AI dependencies (e.g. batch number requires GTIN)
- Repeated AIs
- Character set conformance (cset39, cset64, cset82, etc.)
- Date field plausibility (yymmd0 format)

**C# / .NET integration:**

Not published on NuGet. Distributed as:
- Pre-built `gs1encoders.dll` (Windows portable, from GitHub Releases)
- C# wrapper class in `src/dotnet/` (P/Invoke, idiomatic managed interface)

Integration path for VTCCP:
1. Download the release DLL + C# wrapper from the Releases page
2. Add the C# wrapper `.cs` file to the `DeviceInterface` (or new
   `Validation`) project
3. Bundle `gs1encoders.dll` alongside the `.exe` (or embed as a resource
   and extract on first run — simpler for a single-machine install)
4. Call from the GS1 validation checkbox workflow:

```csharp
// Pseudocode — actual class name TBD after reading src/dotnet/
using GS1.Encoders;

var gs1 = new GS1Encoder();
gs1.ScanData = scanDataFromDmst;  // e.g. "]d1[)>\x1E06\x1D18VD89536..."
bool valid = gs1.HasValidScanData;
string errors = gs1.ErrMarkup;    // marked-up string showing which AIs failed
var ais = gs1.HRI;                // Human Readable Interpretation
gs1.Dispose();
```

### 3. gs1-barcode-engine (not needed for VTCCP)
<https://github.com/gs1/gs1-barcode-engine>

Barcode *symbol generation* (rendering Data Matrix, GS1-128, etc.).
VTCCP does not generate symbols — disregard for now.

---

## ISO 15434 / MIL-STD-129/130 — Situation

**ISO 15434** defines the transport-layer envelope syntax used by both GS1
Format 06 and DoD MIL-STD data:

```
[)>  RS  06  GS  18VD89536  GS  1P8902A  GS  S3122A02965  RS  EOT
```
where RS = 0x1E, GS = 0x1D, EOT = 0x04.

**No open-source ISO 15434 validator exists.** The user's assessment is
confirmed by search. The standard is a DoD/supply-chain niche; commercial
verifiers (Webscan TruCheck, Cognex DMST) implement it internally.

**What VTCCP must implement itself:**

| Layer | Standard | What to parse | Open-source help |
|---|---|---|---|
| Envelope framing | ISO 15434 | Strip `[)>RS`, identify Format DI (`06`=GS1, `05`=ANSI), split on GS/RS/EOT | None — implement in-house (~50 lines) |
| GS1 AI content | GS1 GenSpecs | Validate each AI after stripping the envelope | `gs1-syntax-engine` (above) |
| DoD UID / DI content | ANSI MH10.8.2 + MIL-STD-130N | Validate Data Identifiers (`18`, `1P`, `S`, `25S`, etc.) | None — implement in-house or skip to format-only check |
| MIL-STD-129R labels | MIL-STD-129R | Format 06 envelope, label field presence checks | None — implement in-house |

**Implementation plan (for when this feature reaches D-phase):**

1. Write a small `Iso15434Parser.cs` (~50 lines): detect envelope header,
   extract Format DI, split record on GS delimiter.
2. For GS1 Format 06 (`[)>RS06GS...`): feed the AI element string to
   `gs1-syntax-engine` C# binding.
3. For DoD UID Format 05 (`[)>RS05GS...`): validate DIs against a local
   lookup table built from ANSI MH10.8.2 DI list (no OSS library; table
   is small enough to hand-code from the standard).
4. For raw GS1 (no 15434 envelope, e.g. a plain `]d1...` GS1 scan):
   feed directly to `gs1-syntax-engine`.
5. Report: per-AI/per-DI pass/fail, with the AI/DI key, parsed value,
   and failure reason from the engine's error markup.

The checkbox UI in VTCCP would offer three modes:
- **GS1 syntax check** (uses `gs1-syntax-engine`)
- **MIL-STD-130N / DoD UID check** (uses `Iso15434Parser` + DI table)
- **ISO 15434 framing only** (just checks envelope structure, no content validation)

---

## Data from v1.24 probe — why DMST's `r.validation.gs1` is not the answer

The `DebugGS1` probe in v1.24 enumerated 187 AI keys directly on
`r.validation.gs1` — all returned `undefined`. This means:
- DMST's JS scope does NOT expose AIs as named properties on the
  validation object.
- The GS1 application-syntax checkbox in DMST is a separate engine-level
  function, not accessible from the push script.
- **VTCCP must do its own GS1 validation** — cannot piggyback on the
  device's check.

The v1.25 probe for `r.validation.gs1` will try method-style access
(`.getAI()`, `.parsed`, `.fields`) to understand what the object *does*
expose, but even if something comes back, the above conclusion stands:
application-syntax validation in VTCCP is independent of the device.

---

## What's physically in the repo (v1.4.0)

```
vtccp/lib/gs1-syntax-engine/
  src/GS1Encoder.cs                           ← C# P/Invoke wrapper (1006 lines)
  src/GS1EncoderTest.cs                       ← Official test suite (306 lines)
  src/gs1encoders-dotnet-lib.csproj
  src/README.md
  dotnet-lib-release/gs1encoders-dotnet.dll   ← managed wrapper DLL
  dotnet-lib-release/runtimes/win-x64/native/gs1encoders.dll
  dotnet-lib-release/runtimes/win-x86/native/gs1encoders.dll
  native/x64/gs1encoders.{dll,h,lib}          ← native x64 (C interop)
  native/x86/gs1encoders.{dll,h,lib}          ← native x86
  LICENSE  README.md
  VERSION-PIN.md                              ← version pin + update-check procedure

vtccp/lib/gs1-syntax-dictionary/
  gs1-syntax-dictionary.txt                   ← 344-line AI rules dictionary
  CHANGES
```

See `vtccp/lib/gs1-syntax-engine/VERSION-PIN.md` for the integration path
into VTCCP and the monthly update check procedure.

## Action items (filed, not yet scheduled)

- [x] ~~Download latest `gs1-syntax-engine` release DLL + C# wrapper~~ — done (v1.4.0)
- [ ] Write `Iso15434Parser.cs` (envelope framing only — ~50 lines)
- [ ] Wire into a new `ApplicationSyntaxResult` model in `DeviceInterface`
      or a new `Validation` project
- [ ] Surface as a checkbox in `SessionView` (next to the grading-standard
      selector)
- [ ] Schedule after D1 (report layout) is settled — the validation result
      needs to appear on the report
