# Excel Samples Index

Real TruCheck XLS output — the canonical 167-column schema source.

| File | Source | Notes |
|---|---|---|
| `CalCardProd-2026-03-17.xls` | Production cal-card multi-scan session, March 17, 2026 | Full 167-column schema with real values. Reference for `ExcelEngine` column ordering, multi-scan structure, and 1D averaged-profile column population (`Avg*` fields). |

## More available

`attached_assets/` may contain additional XLS files not yet filed here. Run:
```
ls attached_assets/*.xls attached_assets/*.xlsx
```
to enumerate. File any newly-discovered samples here with a meaningful name
(date + scan-type + symbol-family).

## Canonical schema reference

The 167-column schema implemented by `vtccp/ExcelEngine/` should match the
column ordering and naming in `CalCardProd-2026-03-17.xls`. Any divergence
between the file and the engine is a parser bug.

Cross-reference: `references/samples/scripts/cognex-trucheck-csv-reference.js`
is Cognex's own push-script that produces matching CSV output — useful for
confirming column-name spellings.
