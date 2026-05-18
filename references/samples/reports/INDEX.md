# PDF Report Samples Index

Real DMST and Webscan PDF reports. Two uses:
1. **Parser/reverse-report regression** — these are the formats VTCCP must
   eventually be able to either reproduce or distinguish from.
2. **Trade-dress reference** — visual comparison surface for designing the
   VCCS-distinctive VTCCP report (we want to be obviously different from
   DMST's current look, while incorporating Webscan-lineage column semantics).

## Folders

| Folder | Contents |
|---|---|
| `dmst-2D/` | DMST PDF reports for 2D symbols (Data Matrix, GS1 Data Matrix, fixed-pattern grading tests, UEC tests) — March sprint vintage (2025.4.1.1 DMST) |
| `dmst-1D/` | DMST PDF reports for 1D symbols (UPC, EAN, defects, decodability) — `OMNI_Wide_Angle_*` family from March |
| `dmst-GS1/` | DMST PDF reports for GS1-encoded symbols — important for GS1 element-string parsing reference |
| `dmst-current-sprint/` | DMST PDF reports captured during current sprint (May 17-18, 2026) — useful for comparing 2025-release report format against any 2026-release output we generate later |
| `Webscan-format-2025-09-26.pdf` | **Webscan-format report** — the trade-dress reference; predates Cognex's DMST report design, owned by the lineage VCCS DMV came from |

## Trade-dress strategy

The presence of the Webscan-format report is significant: it predates Cognex's
DMST report design and originates from the Webscan acquisition lineage. When
designing the VCCS-distinctive VTCCP report, we have two reference points:

- **Webscan-format** — the original schema/look; we may legitimately echo
  elements of this format because it's an ancestor of the schema we're
  implementing
- **DMST current** — the format we want to be obviously *different* from
  for trade-dress safety

Cross-referenced in `architecture/vtccp-vs-dmst-feature-matrix.md` (pending).
