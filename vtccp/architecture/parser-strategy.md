# VTCCP Parser Strategy — Architecture Decision

**Logged**: 2026-06-09  
**Decision maker**: George  
**Status**: DECIDED — implementation blocked pending Control Panel UI (see §4)

---

## 1. Decision

**Webscan (Cognex) parser is the default and authoritative parser for all barcode
content parsing in VTCCP** — for GS1 Application Identifiers, GS1 Digital Link
URIs, and ISO 15434 / MH10.8.2 Data Identifiers alike.

User-configurable settings will control whether VTCCP uses:
- DMST parser output only (scraped from pcm_report.html)
- CP (VTCCP-native, gs1-syntax-engine) parser only
- Both parsers simultaneously, with discrepancy flagging

This applies to GS1 AI parsing and GS1 Digital Link. A control panel UI must
be designed before the per-format and per-parser user settings are implemented.

---

## 2. What "Webscan/Cognex parser" means in VTCCP's context

Two distinct channels deliver Webscan/Cognex parser output to VTCCP.
These must not be conflated:

| Channel | What it is | How VTCCP accesses it |
|---|---|---|
| **DMST scrape** | pcm_report.html parsed by DmstHtmlScraper | Already implemented — `TryMergeAsync()` on each scan |
| **gs1-syntax-engine** (E1) | The GS1 reference C# library — the same engine DataMan CT uses for Digital Link | Already integrated — `vtccp/lib/gs1-syntax-engine/` |

For GS1 Digital Link: the gs1-syntax-engine IS the Webscan/Cognex implementation.
For GS1 AI element strings: gs1-syntax-engine is the GS1 reference parser, functionally
equivalent to Webscan's — minor divergence only in edge cases and permissiveness flags.
For ISO 15434 / MH10.8.2 DI data: DMST scrape is the primary Webscan parser output
(the "Data Format Check" table in pcm_report.html). The CP-side ISO15434Parser.cs
currently handles only batch/lot extraction; full DI parsing is not yet built.

---

## 3. Parser routing by format type

| Barcode content type | Detection | Default parser | CP-side parser |
|---|---|---|---|
| GS1 AI element string | AIM ID `]d2`/`]Q3`/`]e0` or `^` prefix or `(NN)` pattern | DMST scrape (HRI from pcm_report.html) | gs1-syntax-engine `AIdataStr` |
| GS1 Digital Link URI | `DecodedData.StartsWith("https://")` or `"http://"` | gs1-syntax-engine `DataStr` setter | gs1-syntax-engine (same) |
| ISO 15434 Format 06 (DI) | `[)>` envelope + `\x1E06` | DMST scrape ("Data Format Check" table) | ISO15434Parser.ParseAllFields() — NOT YET BUILT |
| Plain text / non-GS1 | All other | Pass-through — no AI/DI parse | n/a |

For GS1 Digital Link: DMST scrape and CP parser are the same implementation.
Running "both" on DL input is therefore a configuration option but produces
no additional information in practice — log this as a no-op case.

---

## 4. Control Panel — required before implementation

**The per-format, per-parser user settings SHALL NOT be implemented until the
Control Panel UI is designed and agreed.**

The Control Panel will expose, per format type (GS1 AI / GS1 DLink / ISO 15434):

| Setting | Options |
|---|---|
| Active parser | DMST only / CP only / Both (comparison) |
| On discrepancy (comparison mode) | Log only / Warn operator / Flag in report / Block record |
| Permissiveness | Strict (GS1-compliant only) / Permissive (accept unknown AIs) |

Until the Control Panel is designed:
- **GS1 AI**: use DMST scrape as primary; gs1-syntax-engine as secondary/validation
- **GS1 Digital Link**: use gs1-syntax-engine (already correct)
- **ISO 15434 DI**: use DMST scrape (only complete source); ISO15434Parser for batch/lot only

---

## 5. Dual-parser comparison mode

Running DMST and CP parsers against the same scan and comparing field-by-field
is a high-value QC feature — it catches parser discrepancies, validates data
independently, and provides a cross-check for compliance records.

**Pre-conditions for dual-parser mode:**
1. Control Panel UI defined and implemented
2. A discrepancy schema defined (`ParsedAI[]` vs `ParsedAI[]` diff record)
3. A display surface for discrepancy results in the session window and/or report

Dual-parser mode on GS1 Digital Link is a no-op (same library both sides) —
the Control Panel should grey out this option for DL input or note "parsers identical."

---

## 6. Relationship to existing architecture docs

- `gs1-digital-link.md` §4: DataMan CT DL/AI parser split — VTCCP does NOT replicate
  this split; gs1-syntax-engine handles both. This decision reinforces that.
- `iso15434-mil-std-parser-plan.md` §6: Full DI parser implementation plan — still
  blocked on completing the DI table; DMST scrape is the interim source.
- `vtccp-vs-dmst-feature-matrix.md`: GS1 AI parsing parity items — this decision
  sets the default behaviour as DMST-equivalent until the Control Panel is built.
