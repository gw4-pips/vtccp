---
name: SGTIN-96 GTIN-14 reconstruction formula
description: Exact formula for reconstructing GTIN-14 from SGTIN-96 EPC fields — confirmed from GS1 TDS and verified with worked example
---

**Partition table** (GS1 TDS Table 14-1) — M = GCP bits, L = GCP digits, N = ItemRef bits, K = ItemRef digits:

| P | M  | L  | N  | K |
|---|----|----|----|----|
| 0 | 40 | 12 |  4 | 1  |
| 1 | 37 | 11 |  7 | 2  |
| 2 | 34 | 10 | 10 | 3  |
| 3 | 30 |  9 | 14 | 4  |
| 4 | 27 |  8 | 17 | 5  |
| 5 | 24 |  7 | 20 | 6  |
| 6 | 20 |  6 | 24 | 7  |

L + K = 13 always. M + N = 44 always.

**Key insight:** The Item Reference field (K digits) in SGTIN-96 encoding INCLUDES the GTIN indicator/packaging digit as its leading digit. It is NOT a separate field. So GTIN-14 needs no separate indicator term.

**GTIN-14 reconstruction:**
```
payload13 = GCP_value.ToString().PadLeft(L, '0') + ItemRef_value.ToString().PadLeft(K, '0')
GTIN-14   = payload13 + GS1CheckDigit(payload13)
```

**GS1 Check digit formula** (for 13-char input, 0-indexed from left):
```
weight_i = ((12 - i) % 2 == 0) ? 3 : 1
sum = Σ (digit_i × weight_i)
check = (10 - (sum % 10)) % 10
```
Verified: payload "0001234567890" → sum=85 → check=5 → GTIN-14 "00012345678905" ✓

**Why:** GS1 EPC TDS §6.3.1 states "The Item Reference includes the leading Indicator Digit from the GTIN." Naively the table's K values seem to give L+K=13 digits (not 12), which confused the analysis — but the indicator IS those extra digits.

**SGTIN-198 serial:** 140 bits of packed 7-bit ASCII (20 chars max), null-terminated. Extract with `GetBits(data, startBit + c*7, 7)` for each character, stop at `'\0'`.
