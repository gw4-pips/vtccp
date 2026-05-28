---
name: GS1 DataMatrix detection
description: How to distinguish GS1 DataMatrix from plain DataMatrix in push XML on fw 6.1.16_sr4
---

## Rule

Firmware always emits `<SymbologyName>Data Matrix</SymbologyName>` regardless of GS1 content.
Use `<SymbologyId>` to distinguish:

- `]d1` = plain DataMatrix ECC200
- `]d2` = GS1 DataMatrix (FNC1 in first position — confirmed on fw 6.1.16_sr4, 2026-05-28)

## Fix location

`DmstResultParser.cs` — read `symbIdEarly = Str(map.SymbologyId)` immediately after symbName,
then override `symbology = "GS1 DataMatrix"` and `symbFamily = SymbologyFamily.GS1DataMatrix`
when `symbology == "DataMatrix" && symbIdEarly == "]d2"`.

## GS1 Group Separator (0x1D)

0x1D (ASCII GS, GS1 FNC1 separator between variable-length AIs) is absolutely forbidden in XML 1.0.
`CheckCharacters=false` does NOT help — that flag only covers "not recommended" chars, not strictly-illegal ones.

**Fix**: `xml = xml.Replace('\x1D', '|')` before `XDocument.Load()` in `DmstResultParser.Parse()`.
Pipe is the conventional human-readable GS1 AI separator and is safe in XML.
