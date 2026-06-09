---
name: GS1 Digital Link
description: DL URI syntax, how the GS1 library handles it, and what VTCCP needs to do when DecodedData is a URI
---

Full detail: `vtccp/architecture/gs1-digital-link.md`  
Standard on file: `attached_assets/GS1_Digital_Link_Standard_-_URI_Syntax,_Rel._1.6.0,_March_2025_*.pdf`

**What it is:** Same GS1 AI data encoded as a web URI. Semantically identical to bracketed AI string.
```
(01)09521234000015(10)ABC123(11)210630
https://id.gs1.org/01/09521234000015/10/ABC123?11=210630
```

**Library support:** `encoder.DataStr = "https://..."` → library parses as GS1 DL URI → populates AIdataStr, HRI[] identically to AI input. Bidirectional: `GetDLuri(stem)` goes the other way.

**VTCCP gap (not yet implemented):** When `DecodedData` starts with `https://` or `http://`, route to `encoder.DataStr` instead of current AI parser path. All downstream AI-derived fields then populate correctly.

**Probe needed:** Confirm DM475V push XML delivers raw URI string in `<DecodedData>` for a DL-encoded label (vs pre-converting to AI format).

**DataMan CT split:** CT uses GS1 lib for DL, own parser for AI strings. VTCCP should use library for both — no legacy debt.

**How to apply:** Before any decoded-data parsing code, check `StartsWith("https://") || StartsWith("http://")` and branch to encoder.DataStr path.
