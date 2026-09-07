# FlexWedge regression checklist

**Document version:** 1.0.0  
**Revision date:** 2026-09-06

Use the existing `DeviceInterface.Tests` FlexWedge/EPC vectors as the authoritative EPC inputs.

- [ ] Each SGTIN-96 and SGTIN-198 vector retains raw hex and expected scheme, GTIN, serial, filter, partition, prefix, URI, and bit ranges.
- [ ] Valid, not-found, mismatch, invalid, and no-table GCP cases are visibly distinguished; no table yields `NotChecked`.
- [ ] Manual EPC processing adds one session entry and CSV/JSONL preserve the canonical fields.
- [ ] Raw EPC, GTIN + serial, and mapped-delimited profiles honor prefix, suffix, delimiter, and append key.
- [ ] Clipboard and previously-focused-application delivery contain the selected profile text.
- [ ] A GS1 element string and Digital Link barcode exercise pass/fail serial and GTIN comparison through `RfidValidator`.
- [ ] With the vendor DLL absent, hardware controls explicitly report unavailable and manual mode still works.