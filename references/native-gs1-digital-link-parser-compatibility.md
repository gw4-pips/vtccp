# Native GS1 Digital Link Parser Compatibility

Rev 1.0 — 2026-08-21

## Purpose

This is the reference record for known native-verifier GS1 Digital Link parser
boundaries. It is deliberately kept out of the per-scan VCCS PDF report, which
shows only the concise device-specific note when the condition is observed.

## Recorded boundaries

| Native verifier | Version boundary | Status | Evidence |
|---|---|---|---|
| DataMan / DM TC | 6.1.16 and earlier | Unsupported | The released verifier-line record identifies `6.1.16_sr4` as the latest released artifact. A valid Digital Link scan produced native DFC `FAIL` while the independent VeriWedge GS1 parser passed. |
| DataMan / DM TC | Later than 6.1.16 | Not verified | No later released verifier-line version or passing native Digital Link result is recorded. |
| Webscan TruCheck | 3.03.74 and earlier | Unsupported | The observed TruCheck application version is `3.03.74`; available application and report evidence shows no native Digital Link parser support. |
| Webscan TruCheck | Later than 3.03.74 | Not verified | No later TruCheck version or passing native Digital Link result is recorded. |

## Report behavior

When a native Data Format Check fails but the independent VeriWedge GS1 parser
validates the same Digital Link, the native report column uses `FAIL*` and adds
only this short, device-specific note at the bottom of that column:

> Firmware 6.1.16_tc9 does not support GS1 Digital Link parsing.

The note is not a barcode-data or RFID result and must not be displayed inside,
above, or below the VeriWedge parser block.

## Version interpretation

The implementation compares the numeric version through `6.1.16`. A pre-release
suffix such as `_tc9` therefore remains within that recorded boundary until a
newer released native parser result is verified.