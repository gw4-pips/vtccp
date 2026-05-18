# DMV-8072V Reference Manual Digest
Source: `reference-manual-25.4.1.1.pdf`  
Digest date: 2026-05-18

---

## 1. Product Overview

The **DMV-8072V** is a high-performance fixed-mount barcode verifier and the direct successor/higher-tier descendant of the **Webscan TruCheck** verifier line (acquired by Cognex). This manual covers firmware **25.4.1.1**.

**Relation to DM475V:**
- DM475V is a modular verifier requiring separate attachments for Label vs. DPM lighting.
- DMV-8072V is an integrated system with a larger field of view and native multi-angle lighting (30°/45°/90°).
- Market segment: precision industrial verification (automotive, aerospace, medical) for DPM on challenging surfaces (curved, shiny, matte) alongside high-speed label verification.

---

## 2. Webscan Lineage — Parameter Mapping

The DMV-8072V uses the legacy Webscan 167-column CSV schema. These fields are output via FTP/CSV or mapped in the XML `r.metrics` object.

| Field Name (XML/Push) | Data Type | Unit/Scale | ISO/AIM Standard Section | Description |
|---|---|---|---|---|
| `LLS` | Grade | A–F | ISO 15415 FPD | Left L-Side of DataMatrix finder pattern |
| `BLS` | Grade | A–F | ISO 15415 FPD | Bottom L-Side of DataMatrix finder pattern |
| `TCT` | Grade | A–F | ISO 15415 FPD | Top Clock Track (alternating modules) |
| `RCT` | Grade | A–F | ISO 15415 FPD | Right Clock Track (alternating modules) |
| `TTR` | Percent | % (0–100) | ISO 15415 FPD | Top Transition Ratio (width vs module size) |
| `RTR` | Percent | % (0–100) | ISO 15415 FPD | Right Transition Ratio |
| `LQZ` / `BQZ` | Grade | A–F | ISO 15415 QZ | Left / Bottom Quiet Zone grades |
| `TQZ` / `RQZ` | Grade | A–F | ISO 15415 QZ | Top / Right Quiet Zone grades |
| `HQZ` | Grade | A–F | N/A | Horizontal Quiet Zone (average of L+R) |
| `BWG` | Percent | % | ISO 15415:2024 | Bar Width Growth (Print Growth) |
| `MRD` | Percent | % | N/A | Module Reflectance Distribution (Uniformity) |
| `SCRlRd` | String | % (Rmax/Rmin) | ISO 15415 SC | Symbol Contrast with min/max reflectance |

---

## 3. DPM (Direct Part Mark) Mode

DPM mode (ISO/IEC 29158 / AIM-DPM) uses different grading logic than standard label scanning (ISO/IEC 15415).

**Reflectance Scaling:** Uses "Optimized Brightness" (L₂₅₅ and L₀) rather than absolute 0–100% reflectance to account for low-contrast metallic surfaces.

**Metric substitutions in DPM mode:**

| Standard (ISO 15415) | DPM Equivalent (ISO 29158) |
|---|---|
| `symbolContrast` (SC) | `cellContrast` (CC) |
| `modulation` (MOD) | `cellModulation` (CMOD) |
| `fixedPatternDamage` | `finderPatternDefects` |
| `axial/gridNonuniformity` | `horizontal/verticalMarkMisplacement` |

---

## 4. Geometry Differences from DM475V

| Feature | DM475V | DMV-8072V |
|---|---|---|
| Illumination | Modular (Label Light or DPM attachment) | Native 30°, 45°, 90° all-in-one |
| Aperture | Fixed set | Dynamic (auto-select by X-dimension) |
| Field of View | Standard | ~2× larger |
| Push-Script API | `DmstPushScript_v1.js` | **Identical — 100% API compatible** |
| DMCC API | Standard | Supports `WBU` (Write Buffer) for image playback |

> **Key finding:** Push-script JS API is identical between DM475V and DMV-8072V. Scripts written for DM475V transfer without modification.

---

## 5. Verification Grading Parameters (Full List)

| Parameter | Scale | Symbology | Standard Section |
|---|---|---|---|
| Overall Grade | 0.0–4.0 (A–F) | All | ISO 15415 / 15416 |
| Unused Error Correction (UEC) | 0–100% | DM, QR | ISO 15415 |
| Axial Non-Uniformity (ANU) | Grade A–F | DM, QR | ISO 15415 |
| Grid Non-Uniformity (GNU) | Grade A–F | DM, QR | ISO 15415 |
| Modulation (MOD) | Grade A–F | All | ISO 15415 / 15416 |
| Reflectance Margin (RM) | Grade A–F | 2D | ISO 15415:2024 |
| Symbol Contrast (SC) | Grade A–F | All | ISO 15415 / 15416 |
| Fixed Pattern Damage (FPD) | Grade A–F | DM | ISO 15415 |
| Cell Contrast (CC) | Grade A–F | DM DPM | ISO 29158 |
| Cell Modulation (CMOD) | Grade A–F | DM DPM | ISO 29158 |
| Quiet Zone (QZ) | Grade A–F | All | ISO 15415 |

---

## 6. Quiet Zone Grading — Key Finding

The DMV-8072V provides **per-side** quiet zone reporting (LQZ, BQZ, TQZ, RQZ).

- Reports aggregate: `HQZ` (Horizontal average of L+R) and `MinQZ` (lowest of the four sides).
- **Per-quadrant for large symbols:** For DataMatrix symbols ≥32×32, the device surfaces **ULQZ, URQZ, RUQZ, RLQZ** in the XML output and CSV. This allows pinpointing where the quiet zone violation occurs on high-density codes.

> **Implication for DM475V push channel:** The `q.topQuietZone` / `q.rightQuietZone` structural limit (flat `{grade:X, numericGrade:8.8}` objects with no per-quadrant breakdown) confirmed in v1.27 may be firmware-generation specific. The DMV-8072V on newer firmware exposes individual ULQZ/URQZ/RUQZ/RLQZ. DM475V firmware 6.1.16_sr4 does not.

---

## 7. Modulation Values / Codeword Values

The manual documents internal arrays for Reflectance Grids (per-module reflectance) and Codeword Correctability (per-codeword UEC).

- **Access method for DMV-8072V:** These are NOT exposed in the standard `r.metrics` JS object. They require a DMCC `GET REPORT` command or FTP transfer of the `.xml` / `.pdf` full report, which contains `<ModulationTable>` and `<CodewordTable>` blocks.

> **Note vs DM475V firmware 6.1.16_sr4:** v1.27 probes confirmed that `q.modulationArray`, `q.codewordArray`, and `q.encodationAnalysisArray` ARE directly accessible in the push-script JS object on DM475V firmware 6.1.16_sr4 — no separate DMCC `GET REPORT` needed. This appears to be a firmware-generation difference: newer firmware (DM475V 6.1.16_sr4) exposes these arrays directly in push; older DMV-8072V manual (25.4.1.1 pre-dates this) may describe an older access method.

---

## 8. Configuration

- **Network:** Supports static IP or DHCP; recommended 1000 Mbps for image transfer.
- **Triggering:** Single-button manual trigger or External Pulse (M12 I/O).
- **Push Output:** Configurable via "Network Client" in DMST. Supports formatted strings or the JavaScript `onResult` pipeline.
- **DMCC `WBU`:** Unique to DMV-8072V — "Write Buffer" command for image playback (not present on DM475V).
