# FlexWedge / VeriWedge 1.0 scope and boundary

**Document version:** 1.0.0  
**Revision date:** 2026-09-06

This Windows WPF executable is the focused ASR-P35U RFID wedge release: one COM-connected reader, one manual/demo EPC path, canonical EPC decode, GCP lookup, output profiles, clipboard/focused-application delivery, and in-memory export.

VeriWedge is intentionally limited to one barcode GS1 payload compared with one RFID EPC by `RfidValidator`. It has no verifier-brand automation, no multi-reader or multi-antenna workflow, and no ASR 8-port/tub or bug/demo screen. Future Access, API, SQL, Excel mapping, and richer destination adapters remain additive at the output boundary.

The owned EPC Translator workbook/add-in is pending supplied materials. Its exact behavior and function names are not claimed or invented by this release.