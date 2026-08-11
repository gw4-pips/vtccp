# Transfer Brief — AsReader P35U FlexWedge → Command Pilot (vtccp)

**Date:** 2026-08-06
**From project:** GS1 Resolver (Replit)
**To project:** vtccp / Command Pilot (Replit)
**Prepared by:** Command Pilot agent

---

## Context

Command Pilot (vtccp) is a .NET 8 Windows application for barcode verification with
RFID cross-validation. It needs to integrate the AsReader ASR-P35U UHF RFID reader
via a C# `IEpcReader` implementation. The Python FlexWedge code in GS1 Resolver is
the authoritative protocol reference for that implementation.

The GoToTags E310 reader has been rejected. The ASR-P35U is the canonical hardware
going forward.

---

## What is needed — please export the following

### 1. Full source of VCCS RFID FlexWedge Pro

All Python source files for the FlexWedge tool, including:
- The main entry point / app shell
- The AsReader P35U serial/USB communication layer (everything that talks to the reader)
- The EPC parsing and parameter capture/decoding logic
- Any GS1 cross-validation or GCP lookup code
- Configuration / settings handling
- Any output / report generation code

**Preferred format:** individual `.py` files or a single `.zip`

### 2. Protocol notes / reverse-engineering findings

Any markdown, text, or comment blocks that document:
- The AsReader P35U USB connection type (VCP COM port? HID? other?)
- Baud rate, framing, byte order
- Command packet format (how to trigger an inventory / single read)
- Response packet format (how to extract the EPC from the response)
- Any error/no-tag response codes observed
- CRC algorithm and coverage if applicable
- Any driver requirements (FTDI? AsReader proprietary? None?)

If this is only in code comments, the source files above cover it.

### 3. Sample captured EPC data

Any raw captures (hex strings, byte dumps, or log files) from real AsReader P35U
reads, especially:
- A successful SGTIN-96 read (96-bit EPC)
- A successful SGTIN-198 read (198-bit) if available
- A no-tag / timeout response
- Any error responses

**Format:** text or log files

### 4. Full parameter capture output example

An example of the "full parameter capture and parsing" output — whatever the
FlexWedge currently produces for a successful read. Could be a console dump,
a JSON record, a report snippet — anything that shows all the fields it extracts.

### 5. Any GS1 / EPC validation test cases

If the project has known-good test vectors (input EPC hex → expected GTIN + serial),
export those too. They will become unit tests in the C# implementation.

---

## What will be done with this in vtccp

1. Source goes into `vtccp/references/asr-p35u/` as a read-only protocol reference
2. A C# `AsReaderP35UEpcReader` class implementing the `IEpcReader` interface will be
   written using the Python code as the specification
3. The EPC parsing and GCP validation logic will be compared against vtccp's own
   designed (but not yet coded) C# parser to identify any gaps
4. Sample captures become unit test fixtures
5. The Python code is NOT directly imported or run — it is a reference only

---

## Packaging instruction for the GS1 Resolver agent

Please zip all of the above into a single file named:

`VCCS-RFID-FlexWedge-transfer-2026-08-06.zip`

and make it available for download from the GS1 Resolver project's file panel.
The user will download it there and upload it into the vtccp project here.
