# GS1 Digital Link — Architecture Reference for VTCCP
**Logged**: 2026-06-09  
**Standard**: GS1 Digital Link Standard: URI Syntax, Release 1.6.0, Ratified Mar 2025  
**Reference doc on file**: `attached_assets/GS1_Digital_Link_Standard_-_URI_Syntax,_Rel._1.6.0,_March_2025_1781040754021.pdf` (2897 lines, partially read)  
**Library**: GS1 Barcode Syntax Engine v1.4.0 — `vtccp/lib/gs1-syntax-engine/` — full Digital Link support confirmed in C# wrapper

---

## 1. What GS1 Digital Link Is

GS1 Digital Link is a URI syntax standard (formally GS1 Digital Link Standard: URI Syntax, ISO/IEC 18975) that encodes the same GS1 Application Identifier data as a standard web URI. It is not new content — it is the same AIs already used in traditional barcode symbologies, expressed in a web-native form that can be resolved by a browser.

**Traditional GS1 AI element string:**
```
(01)09521234000015(10)ABC123(11)210630
```

**GS1 Digital Link URI carrying identical data:**
```
https://id.gs1.org/01/09521234000015/10/ABC123?11=210630
```

The two formats are **semantically equivalent**. Every GS1 Digital Link URI can be losslessly round-tripped to/from an AI element string.

---

## 2. URI Structure

```
https://{stem}/{primary_AI}/{primary_value}/{qualifier_AI}/{qualifier_value}?{attr_AI}={value}&{attr_AI}={value}
```

| Segment | Content | Example |
|---|---|---|
| Stem | Any domain + optional path; `https://id.gs1.org/` is the GS1 canonical resolver | `https://id.gs1.org/` |
| Primary key path | One mandatory primary identifier AI (almost always `01` = GTIN) as `/{AI}/{value}` | `/01/09521234000015` |
| Qualifier path segments | Qualifier AIs (`10`=batch, `21`=serial) in defined hierarchy, also in path | `/10/ABC123` |
| Data attribute query params | Remaining AIs as `?{AI}={value}` pairs | `?11=210630` |
| Non-AI query params | Ignored by the library; reported via `DLignoredQueryParams` | `&name=...` |

Key rules:
- AI (01) GTIN must be **14 digits** (GTIN-14). Zero-suppressed GTINs (GTIN-12/13/8) are deprecated but still present in the wild — handled by `PermitZeroSuppressedGTINinDLuris`.
- The stem can be any brand/retailer domain — the label owner is not required to use `id.gs1.org`.
- Unrecognised AIs in the path or query string: controlled by `PermitUnknownAIs`.
- The URI can be resolved by a web browser — the GS1 resolver returns product information. This is the strategic value: one barcode serves both machine-readable GS1 data AND a live web URL.

---

## 3. How the GS1 Syntax Engine Handles Digital Link

Digital Link is a **first-class input format** in the library — not bolted on. The entry point is `DataStr`.

### Parsing (DL URI → AI data)

```csharp
encoder.DataStr = "https://id.gs1.org/01/09521234000015/10/ABC123?11=210630";
// Library detects https:// prefix → parses as GS1 DL URI
// → internally extracts AI element string
// → populates AIdataStr, HRI[] identically to AI-syntax input
string aiStr = encoder.AIdataStr;  // → "(01)09521234000015(10)ABC123(11)210630"
string[] hri = encoder.HRI;        // → ["(01) 09521234000015", "(10) ABC123", "(11) 210630"]
```

### Generation (AI data → DL URI)

```csharp
encoder.AIdataStr = "(01)09521234000015(10)ABC123(11)210630";
string dlUri = encoder.GetDLuri(null);
// → "https://id.gs1.org/01/09521234000015/10/ABC123?11=210630"  (canonical stem)

string branded = encoder.GetDLuri("https://brand.example.com");
// → "https://brand.example.com/01/09521234000015/10/ABC123?11=210630"
```

### Key API surface for Digital Link

| Member | Purpose |
|---|---|
| `DataStr` setter (https:// prefix) | Parse a GS1 DL URI → extract AI data |
| `GetDLuri(stem)` | Convert AI data → GS1 DL URI (null = GS1 canonical stem) |
| `DLignoredQueryParams` | Returns non-AI query params that were present but ignored |
| `PermitZeroSuppressedGTINinDLuris` | Accept GTIN-12/13/8 in path (legacy DL URIs) |
| `PermitUnknownAIs` | Accept AIs not in the library's table (applies to DL URI input) |
| `Validation.UnknownAInotDLattr` | Reject unknown AIs used as DL data attributes |

---

## 4. How DataMan CT Uses Digital Link (and What It Means for VTCCP)

DataMan CT uses the GS1 Syntax Engine **only for Digital Link URI parsing**. For traditional GS1 AI element strings, CT uses its own internally developed parser. This split is architectural legacy — CT's internal AI parser predates the GS1 Syntax Engine and is deeply integrated into the grading pipeline.

**VTCCP should NOT replicate this split.** Our internal AI handling is shallow, with no legacy debt to protect. The correct VTCCP architecture is:

| Input type | Detection | Routing |
|---|---|---|
| GS1 Digital Link URI | `DecodedData.StartsWith("https://") \|\| .StartsWith("http://")` | `encoder.DataStr = DecodedData` → then read `AIdataStr`, `HRI[]` |
| Traditional GS1 AI string | Starts with `(01)` or FNC1-encoded (`^`) | Current AI parsing path (unchanged) |
| Non-GS1 data | Neither of the above | Pass through as raw string; no AI parsing |

One parser surface (`GS1Encoder`) for both input types — consistent validation, consistent HRI output, no divergence.

---

## 5. VTCCP Exposure to Digital Link — What Needs to Change

**Current state**: `DmstResultParser` and `BarcodeDataFormatter` assume `DecodedData` is always a traditional AI string or plain text. A DL URI would pass through the control-character substitution path, get `<F1>` prepended if the AIM ID is `]d2`/`]Q3` etc., and be stored as a raw URI string in the Excel/record — not parsed into AI fields.

**What happens on a DL-encoded label today**: The grading result is captured correctly (grading is independent of decoded content). But `ParsedAI`, `HRI`, and any AI-derived fields (lot number, GTIN, expiry date) would be empty or garbage.

**Required change (not yet implemented)**:
1. In `BarcodeDataFormatter` or the record assembly step: detect DL URI prefix in `DecodedData`
2. Route through `encoder.DataStr` → read `encoder.AIdataStr` and `encoder.HRI[]`
3. Populate AI-derived fields from the parsed output exactly as for traditional AI input

**Probe needed**: Confirm what the DM475V push XML delivers in `<DecodedData>` when it scans a DL-encoded QR Code or Data Matrix. Expected: the raw URI string. But verify — some firmware may pre-process it to AI format. Use a real GS1 DL label on the bench.

---

## 6. Practical Prevalence — When You Will See Digital Link

Digital Link encoding is increasingly common in:
- **Retail 2D** (QR Code on consumer packaged goods) — major grocery/retail chains moving to 2D barcodes with DL URIs by 2027 per GS1 "Sunrise 2027" initiative
- **Pharmaceutical serialisation** — DSCSA-compliant labels often use DL URIs in GS1 DataMatrix
- **Healthcare/medical device** — UDI labels may encode DL URIs

For the DM475V environment (label printing verification, not retail scanning), you are more likely to encounter it in pharmaceutical and medical device customers than in traditional industrial label applications. But the "Sunrise 2027" transition means DL will become the standard 2D label format broadly — VTCCP must handle it before that wave arrives.

---

## 7. Reference Standard

- **GS1 Digital Link Standard: URI Syntax, Release 1.6.0, Ratified Mar 2025**
  Full text on file: `attached_assets/GS1_Digital_Link_Standard_-_URI_Syntax,_Rel._1.6.0,_March_2025_1781040754021.pdf`
  2,897 lines total; contributors list + full specification.
- The GS1 Barcode Syntax Engine (v1.4.0, `vtccp/lib/gs1-syntax-engine/`) is the GS1 reference implementation of this standard. The library's test suite in `src/GS1EncoderTest.cs` contains extensive DL URI test cases that serve as a secondary specification reference.
