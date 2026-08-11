---
name: SGTIN-96 GTIN-14 formula
description: Correct GS1 GTIN-14 reconstruction from SGTIN-96 EPC fields; previous version had wrong field order.
---

## Rule

GTIN-14 = indicator(1) + GCP.PadLeft(L, '0') + item_body.PadLeft(K-1, '0') + GS1CheckDigit(payload13)

Where:
- indicator = ItemRef / 10^(K-1)  (MSB decimal digit of the ItemReference field)
- item_body = ItemRef % 10^(K-1)  (remaining K-1 digits)
- payload13 = indicator.ToString() + GCP.PadLeft(L) + item_body.PadLeft(K-1)

The indicator digit precedes the GCP — this is the common mistake.

**Wrong (old):** payload13 = GCP.PadLeft(L) + ItemRef.PadLeft(K)

**Why:** GS1 GTIN-14 structure is [indicator][GCP][item reference body][check digit]. The ItemReference field in the EPC encodes (indicator × 10^(K-1) + body), so the indicator must be extracted and placed before the GCP in the output, not left concatenated after it.

**How to apply:** Both Sgtin96Decoder and Sgtin198Decoder use this formula. Any future scheme decoder that reconstructs a GTIN must split ItemRef into indicator + body before building the payload string.

## Check digit weight rule

For a 13-digit payload: weight = 3 if (12 - i) % 2 == 0, else 1, for i = 0..12 (left to right).

## Partition table (L+K=13, L+K always=13)

| P | M  | L  | N  | K |
|---|----|----|----|----|
| 0 | 40 | 12 |  4 |  1 |
| 1 | 37 | 11 |  7 |  2 |
| 2 | 34 | 10 | 10 |  3 |
| 3 | 30 |  9 | 14 |  4 |
| 4 | 27 |  8 | 17 |  5 |
| 5 | 24 |  7 | 20 |  6 |
| 6 | 20 |  6 | 24 |  7 |

M = GCP bits, L = GCP decimal digits, N = ItemRef bits, K = ItemRef decimal digits.

## Test vector note

The JSON test-vector file (references/asr-p35u/test-vectors/epc-decode-vectors.json) has incorrect serial values for live-B and live-C, and an incorrect GCP for the defect-repro vector. These were created from an earlier Python decoder with its own bug. The EpcParserTests assertions use bit-correct values (verified by hand from the EPC bytes).
