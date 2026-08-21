---
name: GS1 parser report layout
description: Approved presentation rule for the VCCS GS1 parser block in TruCheck-style reports.
---

The VeriWedge GS1 parser must be presented as an independent Field / Data / Check block, not as values borrowed from or limited by the DataMan TruCheck rows.

For a GS1 Digital Link, the original URI is always the first parser row. Parsed Application Identifiers follow using GS1 field names. The canonical GS1 Element String belongs in the final parser row and may wrap only between complete AI elements. Both Digital Links and Element Strings use this comparison layout whenever VCCS has an applicable parser result.

**Why:** Each DMST-connected trigger provides native TruCheck output as well as VCCS parser data. The parser can contain more GS1 detail than the native verifier table, and the URI is the source artifact for a Digital Link rather than a parser-derived field.

**How to apply:** Preserve this order for future GS1 PDF/report designs. Do not replace the native DataMan panel with a fallback because the source is Digital Link; keep it beside the VCCS parser panel. Do not make parser rows contingent on matching native verifier rows or place the canonical AI string ahead of the URI.