# AIM Barcode Symbology Identifiers

**Source**: Honeywell SPS Support Knowledge Base, Article 000063426
**URL**: https://sps-support.honeywell.com/s/article/List-of-barcode-symbology-AIM-Identifiers
**Scraped**: 2026-05-18
**Applies to**: Honeywell scanners and imagers (also used as reference standard for Cognex DataMan output — AIM IDs are a cross-vendor standard)

---

## Format

Each AIM Code Identifier is a three-character string `]cm` where:
- `]` — Flag character (literal right-bracket)
- `c` — Code character (identifies the symbology family)
- `m` — Modifier character (encodes options: check digit handling, ECI, FNC1 position, etc.)

**Example**: Full ASCII Code 39 with check character W, A+I+MI+DW → `]A7` where modifier 7 = (3+4).

---

## Complete Table

### Code 39 — `]A`

| AIM ID | Description |
|---|---|
| `]A0` | No check character, no Full ASCII processing |
| `]A1` | Mod-43 check performed and transmitted |
| `]A2` | Mod-43 check performed and stripped |
| `]A4` | Full ASCII conversion performed; no check digit validation |
| `]A5` | Full ASCII conversion + check character verified and transmitted |
| `]A7` | Full ASCII conversion + check character verified and stripped |

### Telepen — `]B`

| AIM ID | Description |
|---|---|
| `]B0` | Full ASCII mode |
| `]B1` | Double density numeric mode |
| `]B2` | Double density numeric followed by full ASCII |
| `]B4` | Full ASCII followed by double density numeric |

### Code 128 — `]C`

| AIM ID | Description |
|---|---|
| `]C0` | Standard. No FNC1 in first or second symbol character position after start character |
| `]C1` | FNC1 in first character position — GS1-128 / GS1 DataBar Expanded |
| `]C2` | FNC1 in second character position — ISBT concatenation |
| `]C4` | ISBT (International Society for Blood Transfusion) concatenation performed |

### UPC / EAN — `]E`

| AIM ID | Description |
|---|---|
| `]E0` | Standard full EAN country code format (13 digits for UPC-A / UPC-E, or Bookland EAN) |
| `]E1` | Two-digit supplement data only |
| `]E2` | Five-digit supplement data only |
| `]E3` | EAN-13 with 2/5-digit add-on, or UPC-E with 2/5-digit add-on / Extended Coupon Code / EAN-8 with add-on |
| `]E4` | EAN-8 |

### Codabar — `]F`

| AIM ID | Description |
|---|---|
| `]F0` | No check digit processing |
| `]F1` | Check digit validated and transmitted; ABC Codabar concatenate/message append performed |
| `]F3` | Check digit validated and stripped |

### Code 93 — `]G`

| AIM ID | Description |
|---|---|
| `]G0` | Standard (modifiers 0–9, A–Z, a–m for various options) |

### Code 11 — `]H`

| AIM ID | Description |
|---|---|
| `]H0` | Single check digit validated and transmitted |
| `]H1` | Two check digits validated and transmitted |
| `]H3` | Check characters validated but not transmitted |

### Interleaved 2 of 5 — `]I`

| AIM ID | Description |
|---|---|
| `]I0` | No check digit processing |
| `]I1` | Check digit validated |
| `]I3` | Check digit validated and stripped |

### PDF417 — `]L`

| AIM ID | Description |
|---|---|
| `]L0` | 1994 PDF417 spec protocol; ECI invocation status uncertain; byte 92 not doubled |
| `]L1` | ENV 12925 Extended Channel Interpretation; all byte-92 characters doubled |
| `]L2` | Basic channel operation; byte-92 not doubled; structured append and ECI not available |
| `]L3` | Micro PDF417 — Code 128 emulation, FNC1 implied in first position |
| `]L4` | Micro PDF417 — Code 128 emulation, FNC1 implied after initial letter or pair of digits |
| `]L5` | Micro PDF417 — Code 128 emulation, no implied FNC1 |

### MSI Plessey — `]M`

| AIM ID | Description |
|---|---|
| `]M0` | Single check digit verified and transmitted |
| `]M1` | Two check digits verified; check digit not transmitted (single) |
| `]M2` | Single check digit verified and stripped |
| `]M3` | Two check digits verified and stripped |

### Codablock — `]O`

| AIM ID | Description |
|---|---|
| `]O1` | Codablock 256, FNC1 in first data character; subsequent FNC1 → ASCII 29 (GS) |
| `]O4` | Codablock F, no FNC1 |
| `]O5` | Codablock F, FNC1 in first data character; subsequent FNC1 → ASCII 29 |
| `]O6` | Codablock A |
| `]Om` | Codablock 256, no FNC1 |

### Standard Plessey — `]P`

| AIM ID | Description |
|---|---|
| `]P0` | No options specified (always modifier 0) |

### QR Code — `]Q`

| AIM ID | Description |
|---|---|
| `]Q0` | Model 1 |
| `]Q1` | Model 2 (QR Code 2005) — **ECI protocol NOT implemented** |
| `]Q2` | Model 2 (QR Code 2005) — **ECI protocol implemented** |
| `]Q3` | Model 2, ECI not implemented, FNC1 implied in first position (GS1 QR) |
| `]Q4` | Model 2, ECI implemented, FNC1 implied in first position |
| `]Q5` | Model 2, ECI not implemented, FNC1 implied in second position |
| `]Q6` | Model 2, ECI implemented, FNC1 implied in second position |

### Standard 2 of 5 — `]R`

| AIM ID | Description |
|---|---|
| `]R0` | No check digit verification |
| `]R1` | Check digit verified but not transmitted |
| `]R2` | Check digit verified and transmitted |

### Discrete 2 of 5 — `]S`

| AIM ID | Description |
|---|---|
| `]S0` | Straight 2 of 5 Industrial / IATA 2 of 5; no options specified |

### Code 49 — `]T`

| AIM ID | Description |
|---|---|
| `]Tm` | Modifiers 0, 1, 2, 4 (varies) |

### Maxicode — `]U`

| AIM ID | Description |
|---|---|
| `]U0` | Mode 4 or 5 |
| `]U1` | Mode 2 or 3 |
| `]U2` | Mode 4 or 5, ECI implemented |
| `]U3` | Mode 2 or 3, ECI implemented in secondary message |

### Miscellaneous — `]X`

| AIM ID | Description |
|---|---|
| `]X0` | Trioptic Code 39 / Bookland EAN / Code 32 Pharmaceutical (PARAF) / China Post / Matrix 2 of 5 / NEC 2 of 5 / Chinese Sensible Code (Han Xin) / any Postal symbologies |

### No Barcode — `]Z`

| AIM ID | Description |
|---|---|
| `]Z` | No barcode data |

### Data Matrix — `]d`

| AIM ID | Description |
|---|---|
| `]d0` | ECC 000–140 (legacy, pre-ECC200) |
| `]d1` | **ECC 200** (standard; most common) |
| `]d2` | ECC 200, FNC1 in first or fifth position (GS1 Data Matrix) |
| `]d3` | ECC 200, FNC1 in second or sixth position |
| `]d4` | ECC 200, ECI protocol implemented |
| `]d5` | ECC 200, FNC1 in first or fifth position, ECI protocol implemented |
| `]d6` | ECC 200, FNC1 in second or sixth position, ECI protocol implemented |

### GS1 DataBar — `]e`

| AIM ID | Description |
|---|---|
| `]e0` | GS1 DataBar / GS1 DataBar Limited / GS1 DataBar Expanded |
| `]e1` | Data following an encoded symbol separator character |
| `]e2` | Data following an escape mechanism character; ECI protocol NOT supported |
| `]e3` | Data following an escape mechanism character; ECI protocol supported |

### Aztec — `]z`

| AIM ID | Description |
|---|---|
| `]z0` | No options |
| `]z1` | FNC1 preceding first message character |
| `]z2` | FNC1 following an initial letter or pair of digits |
| `]z3` | ECI protocol implemented |
| `]z4` | FNC1 preceding first message character + ECI implemented |
| `]z5` | FNC1 following letter or digit pair + ECI implemented |
| `]z6` | Structured append header included |
| `]z7` | Structured append + FNC1 preceding first message character |
| `]z8` | Structured append + FNC1 following letter or digit pair |
| `]z9` | Structured append + ECI implemented |
| `]zA` | Structured append + FNC1 preceding first message character + ECI |
| `]zB` | Structured append + FNC1 following letter or digit pair + ECI |
| `]zC` | "Rune" decoded |

---

## VTCCP-specific notes

### Confirmed on Cognex DM475V fw 6.1.16_sr4

| AIM ID | Observed | Notes |
|---|---|---|
| `]d1` | v1.23 + v1.24 live DM scans | ECC 200, standard Data Matrix — always emitted for this device |
| `]d2` | (not yet captured) | GS1 Data Matrix — would be seen if GS1 DM scanned with FNC1 |
| `]Q1` | v1.24 email QR loaded-image | QR Code 2005, no ECI — byte-mode payload, no `\000026` prefix stripping needed |
| `]Q2` | v1.24 URL QR loaded-image | QR Code 2005, ECI implemented — `\000026` (ECI 26 = UTF-8) prefix must be stripped from `<DecodedData>` |

### Parser branching rules (DmstResultParser)

```
SymbologyId[0] == ']'       → always true (all AIM IDs start with ])
SymbologyId[1] == 'd'       → Data Matrix family
SymbologyId[1] == 'Q'       → QR Code family
  SymbologyId[2] == '2'     → ECI present → strip \x00\x00\x26 prefix from DecodedData
  SymbologyId[2] == '4'     → ECI + GS1 FNC1 in first position
  SymbologyId[2] == '6'     → ECI + GS1 FNC1 in second position
  otherwise                 → no ECI, no stripping
SymbologyId[1] == 'C'       → Code 128 family
  SymbologyId[2] == '1'     → GS1-128 (FNC1 in first position)
SymbologyId[1] == 'e'       → GS1 DataBar family
SymbologyId[1] == 'L'       → PDF417 family
SymbologyId[1] == 'z'       → Aztec family
```

### GS1 FNC1 detection (for VTCCP gs1-syntax-engine trigger)

GS1 mode is indicated by FNC1 in the data stream. AIM IDs that indicate GS1 content:
- `]C1` — GS1-128
- `]d2`, `]d5` — GS1 Data Matrix (FNC1 in first/fifth position)
- `]d3`, `]d6` — GS1 Data Matrix (FNC1 in second/sixth position)
- `]Q3`, `]Q4` — GS1 QR Code (FNC1 in first position)
- `]Q5`, `]Q6` — GS1 QR Code (FNC1 in second position)
- `]e0`, `]e1`, `]e2`, `]e3` — GS1 DataBar family (always GS1)

When `SymbologyId` matches any of the above, VTCCP should pass `<DecodedData>` to `gs1-syntax-engine`
for AI-level validation regardless of the device's Application Settings `Data Format Check` setting.

### ECI stripping rule

ECI mode 26 = UTF-8 encoding. When `]Q2` (or `]Q4`, `]Q6`, `]d4`, `]d5`, `]d6`, `]z3`+ variants):
- The `<DecodedData>` field begins with an ECI designator prefix (`\000026` for ECI 26 / UTF-8)
- Strip this prefix before display or storage of the human-readable payload
- Store the raw (unstripped) value as `<DecodedDataRaw>` for audit purposes

---

## Limitations of this source

This is a Honeywell scanner KB article — it documents Honeywell device behavior. AIM identifiers
are a cross-vendor standard (published by AIM International as "Guidelines for Direct Part Marking"),
but specific modifier assignments and which modifiers a given device emits may vary by manufacturer.
The Cognex DM475V emits `]d1` for ECC200 DM (confirmed) and `]Q1`/`]Q2` for QR (confirmed).
Other modifier values on the Cognex device should be treated as hypothetical until captured in a
live scan XML.

**Supersede with**: ISO/IEC 15424:2008 (Data Carrier Identifiers, including Symbology Identifiers)
when the user acquires it. That is the normative standard; this KB article is a working reference.
