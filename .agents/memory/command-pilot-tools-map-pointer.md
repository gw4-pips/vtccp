---
name: Command Pilot tools map
description: Which of the 5 VTCCP tools provides which data; which fields have no alternative to the HTML scraper
---

Full detail: `vtccp/architecture/command-pilot-tools-map.md`

**5 tools, each owns distinct data:**

1. **Cognex SDK** (port 44444) — connect/session/device metadata only. Cannot trigger, cannot run arbitrary DMCC, XmlResultArrived dead for external triggers, SetResultTypes() forbidden.
2. **Raw TCP DMCC** (port 23) — software trigger, IMAGE.SEND live view, GET/SET DMCC keys. Every command needs `||>` prefix. Default response mode = 0 (silent); must SET 2 first.
3. **HTTP pub/sub** (port 44444, codes.xml) — PRIMARY result delivery. All ISO grades, decoded data, JPEG, timestamps. General Characteristics block has CORRECT EncodedCharacters/DataCodewords/ECBudget.
4. **DMST HTML Scraper** (pcm_report.html) — ONLY source for: ImagePolarity ✓, ECLevel (QR pending), DataMaskPattern (QR pending), ECI (pending). Also: correct EncodedCharacters, DataCodewords, ErrorCorrectionBudget.
5. **GS1 Syntax Engine** (local lib) — AI validation, HRI, Digital Link ↔ AI conversion.

**Permanently unresolvable from push XML on fw 6.1.16_sr4 (HTML only):**
ImagePolarity, ECLevel (QR), DataMaskPattern (QR), ECI value.

**Key invariants:** TRIGGER.TYPE=0 always; SetResultTypes() never; COM.DMCC-SAVE never without intent; LIVEIMG.MODE=0; `||>` on every port-23 command.
