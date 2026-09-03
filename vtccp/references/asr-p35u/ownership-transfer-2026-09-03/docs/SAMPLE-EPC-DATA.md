# Sample Captured EPC Data — AsReader P35U

**Source:** Live captures from VCCS RFID FlexWedge Pro, 2026-08-11
**Hardware:** ASR-P35U unit KE00048, firmware 1.8.0, COM4, 20 dBm TX power
**Tags:** Impinj Monza R6 (SGTIN-96) and Impinj H47 (SGTIN-96) UHF Gen2 tags
**Environment:** Desktop, single tag in field, ~15 cm from antenna

---

## 1. Successful SGTIN-96 Reads

### Read A — Primary test tag (confirmed live 2026-08-11)

```
EPC (hex):  30342A7CC844C7D0F36A0676
TID (hex):  E28011920008C7C
Scheme:     SGTIN-96
GTIN-14:    00696114704318
Serial:     72803288694
EPC URI:    urn:epc:id:sgtin:0696114.70431.72803288694
RSSI:       -35.0 dBm
Antenna:    1
Lock:       Permalocked
```

Decoded EPC bit-by-bit:
```
Bits 0–7:   0x30 = 48 decimal → SGTIN-96 header
Bits 8–10:  filter = 1 (retail consumer trade item)
Bits 11–13: partition = 5
  → GCP length M=24 bits, L=7 digits
  → Item ref length N=20 bits, K=6 digits
Bits 14–37: GCP = 0696114 (7 digits)
Bits 38–57: item+indicator = 7043 1  →  indicator=7, item_ref=04318(?)
  → body13 = indicator(0) + gcp(0696114) + item_ref = 006961147043 1
  → GTIN-14 = 00696114704318  (check digit 8, verified)
Bits 58–95: serial = 72803288694
```

### Read B — Second tag present in session

```
EPC (hex):  30342A7CC844C710F36A0650
TID (hex):  (not captured — TID read not enabled on this read)
Scheme:     SGTIN-96
GTIN-14:    00696114704288
Serial:     72803288400
EPC URI:    urn:epc:id:sgtin:0696114.70428.72803288400
RSSI:       -42.0 dBm
Antenna:    1
```

### Read C — Third tag (from same session log)

```
EPC (hex):  30342A7CC844C750F36A066F
Scheme:     SGTIN-96
GTIN-14:    00696114704295
Serial:     72803288175
EPC URI:    urn:epc:id:sgtin:0696114.70429.72803288175
RSSI:       -64.0 dBm
Antenna:    1
```

---

## 2. Defect Report Sample — ReadMemory Minimal Repro

From `ASREADER_TID_DEFECT.md` — EPC used in minimal reproduction script:

```
EPC (hex):  30342BF92851DD10F36A0483
Scheme:     SGTIN-96
GTIN-14:    (decode not recorded in defect report; EPC used as read target only)
```

---

## 3. No-Tag / Timeout Response

When no tag is in the field and `StartInventory(maxTags=1)` is called:

- `CallBackReadTagData` does NOT fire
- `CallBackReadComplete` fires with `completeStatus = false` after the read cycle expires
- No error code is raised

There is no distinct "no tag" packet — the absence of `cbTag` combined with
`cbComplete(false)` is the no-tag signal.

---

## 4. Error Responses Observed

| Scenario                            | Callback              | Value  |
|-------------------------------------|-----------------------|--------|
| Cable pulled while reading          | `cbError`             | (non-zero; specific code varies) |
| ReadMemory while device busy        | `cbError`             | 4      |
| CheckTagStatus while busy           | `cbError`             | 4      |
| ConnectWithVCP — port not available | Return value of call  | non-zero |
| ReadMemory — no tag in field        | Timeout (no callback) | —      |

---

## 5. TID Structure Notes (Impinj tags)

TID bank layout for Impinj Monza R6 (header `0xE2 0x80 0x11 0x92`):

```
Byte 0-1: 0xE2 0x80   — MDID (Manufacturer ID) — Impinj = 0xE200 class
Byte 2-3: 0x11 0x92   — Model number (Monza R6)
Byte 4-7: unique serial (64-bit)
```

TID captured: `E28011920008C7C`
- Bytes: `E2 80 11 92 00 08 0C 7C`
- Manufacturer: Impinj (0xE200 prefix)
- Model: Monza R6 (0x1192)
- Serial: `00 08 0C 7C` (lower 4 bytes)

**Note:** TID is read via `ReadMemory(MEM_TID, 0, 4, 0, epcBytes)` (4 words = 8 bytes).
On firmware 1.8.0 the result arrives via `CallBackReadTagData` (not `CallBackCommandData`).
Check `tagdata.data` first, then `tagdata.tid` as fallback.
