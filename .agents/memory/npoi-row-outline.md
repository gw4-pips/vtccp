---
name: NPOI IRow.OutlineLevel read-only
description: NPOI IRow.OutlineLevel has no setter in the IRow interface for this NPOI build; workaround and implication for XlsAdapter
---

## Rule
`NPOI.SS.UserModel.IRow.OutlineLevel` is **read-only** (getter only) in this NPOI build.
Attempting `row.OutlineLevel = (short)level;` produces CS0200 at compile time.

**Why:** NPOI's `IRow` interface exposes the property for reading (for consumers that
query outline depth) but does not expose a setter. The concrete `HSSFRow` class may
have a setter, but casting through the interface fails.

**How to apply:**
- In `XlsAdapter.SetRowOutlineLevel()`: implement as a **no-op**. Outline grouping for
  parse-detail rows is a progressive-enhancement UX feature; its absence on the legacy
  `.xls` path is acceptable — rows are still written and visible without the `+/-` button.
- If outline grouping ever becomes required for `.xls`, cast to `NPOI.HSSF.UserModel.HSSFRow`
  and set directly, or use `ISheet.GroupRow(rowFrom, rowTo)` which sets level via grouping API.
- For `.xlsx` (EPPlus), `ExcelRow.OutlineLevel` has a proper setter — no workaround needed.
