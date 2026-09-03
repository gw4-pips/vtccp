# FlexWedge RFID / VeriWedge RFID Transfer Manifest

**Prepared:** 03 September 2026  
**Destination:** VTCCP Command Pilot  
**Package status:** Formal ownership and stewardship transfer

## Package structure

### `source/`

- `main.py` — validated tkinter application and workflow
- `reader.py` — AsReader P35U/pythonnet integration
- `decoder.py` — EPC and GTIN decoding
- `injector.py` — Windows keyboard injection
- `config.py` — persisted application configuration
- `requirements.txt` — Python runtime dependencies
- `build_exe.bat` — Windows PyInstaller build
- `tests/test_config_buzzer.py` — configuration/buzzer tests
- `tests/test_preset_tid.py` — preset and TID behavior tests

### `sdk/`

- `AsReaderP3xU.dll` — VCCS-supplied proprietary AsReader dependency
- `THIRD_PARTY_NOTICE.md` — ownership and redistribution notice

### `docs/`

- `OWNERSHIP_TRANSFER.md` — authority and responsibility transfer
- `TRANSFER_MANIFEST.md` — this inventory
- `README.md` — application overview and operation
- `LAPTOP_QUICKSTART.md` — clean-machine setup
- `PROTOCOL-NOTES.md` — P35U integration reference
- `ASREADER_TID_DEFECT.md` — resolved callback investigation
- `FIRMWARE_LOG.md` — development-unit provenance
- `SAMPLE-EPC-DATA.md` — sample EPC reference
- `PARAMETER-CAPTURE-EXAMPLE.md` — parameter capture reference

### `test-vectors/`

- `epc-decode-vectors.json` — decoder interoperability vectors

### Package root

- `SHA256SUMS.txt` — checksums for transferred files
- `VALIDATION.txt` — preparation and fresh-laptop validation record

## Source-selection note

The packaged `main.py` is the last clean compiling implementation containing
the TID read-log update and is the exact application version used in the
03 September 2026 laptop validation. The packaged supporting modules are the
latest validated versions paired with that application during the same test.

This deliberate validated snapshot takes precedence over any newer but
unvalidated or malformed historical copy in the transferring repository.

## Runtime files intentionally omitted

- `.venv/`
- `__pycache__/`
- `*.pyc`
- `rfid_wedge_config.json`
- `TagLog.csv`
- `debug.log`
- `build/`
- `dist/`
- PyInstaller specification files
