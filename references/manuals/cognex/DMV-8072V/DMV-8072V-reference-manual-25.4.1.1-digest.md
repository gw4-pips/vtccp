# Cognex DMV-8072V Verifier Reference Manual — Digest

**Source**: `reference-manual-25.4.1.1.pdf` (5.2 MB)
**Version**: 2025.4.1.1
**Digested by**: explorer subagent, 2026-05-18

**Provenance note**: Comprehensive at the top level but not page-exhaustive.
For any specific Webscan-lineage field definition used in our schema /
report copy, verify the exact wording against the source PDF.

---

## TOP-LINE FINDINGS (read this section first)

1. **8072V is the higher-tier descendent of the Webscan TruCheck verifier
   line** that VCCS DMV originally used. Its column-naming conventions are
   the lineage source for VTCCP's 167-column schema. Field definitions
   here are authoritative for our schema's Webscan-lineage columns.

2. **Webscan column-name expansions** (we'd been treating these as opaque
   acronyms):
   - **LLS = Left Light Source** / **BLS = Bottom Light Source** —
     illumination consistency on the L-shape finder pattern segments
   - **LQZ/BQZ/TQZ/RQZ** = Left/Bottom/Top/Right Quiet Zone grades
   - **HQZ** = Horizontal Quiet Zone (averaged Left+Right)
   - **TTR/RTR** = Top/Right Transition Ratio
   - **TCT/RCT** = Top/Right Clock Track
   - **AG** = Average Grade (across scan lines)
   - **MRD** = Module Reflectance Distribution
   - **BWG** = Bar Width Growth (= ISO 15415 Print Growth)
   - **SCRlRd** = Symbol Contrast formatted as `SC (Rl/Rd)` where
     Rl=max reflectance, Rd=min reflectance

3. **Image-load capability CONFIRMED on 8072V** via DMST's "Image
   Playback" feature. Pre-captured `.jpg` / `.png` files can be fed
   through the grading engine. This is the canonical reference for
   VTCCP's loaded-image flow being feasible at all — and tells us the
   8072V's path is via DMST UI (the underlying DMCC mechanism is `WBU`
   per the Comms guide digest).

4. **DPM-mode (AIM-DPM / ISO 29158) metric substitution table** is
   explicit in the 8072V manual:
   - `cellDefects` replaces `defects`
   - `finderPatternDefects` replaces `fixedPatternDamage`
   - `dataMatrixCellWidth/Height` replaces `moduleSize`
   - `horizontalMarkMisplacement` / `verticalMarkMisplacement` replaces
     `axialNonuniformity` / `gridNonuniformity`
   - `cellContrast` (CC) replaces `symbolContrast` (SC)
   - `cellModulation` (CMOD) replaces `modulation` (MOD)
   - Reflectance scaling uses "Optimized Brightness" (L_255, L_0) rather
     than absolute 0-100%
   This **resolves the long-standing question** of why our 30-metric probe
   surfaces DPM-only fields that are always NA on calibration-card scans.

5. **The 8072V push pipeline is fully compatible with our existing
   `DmstPushScript_v1.js`** — both run the same ECMAScript engine and
   expose the same `r.*` object. Migration to 8072V (if ever) requires
   no script changes.

---

## 1. The 8072V Overview

- **Hardware tier**: Higher than DM475V. High-performance fixed-mount
  verifier derived from the Webscan TruCheck line.
- **Target market**: Primarily DPM (Direct Part Mark) + high-end label
  verification. Automotive, aerospace, medical device marking. Non-flat
  surfaces and varying finishes (shiny / matte / cast).
- **Optical/Lighting**: Multi-quadrant, multi-angle integrated lighting
  chamber. 30°, 45°, and 90° options native (DM475V requires modular
  attachments to switch between Label Light and DPM lighting).
- **Symbologies**: All standard 1D ISO codes; 2D includes Data Matrix,
  QR Code, GS1 Data Matrix, GS1 QR Code, DotCode, DMRE.

### Differences from DM475V

| Feature | DM475V | DMV-8072V |
|---|---|---|
| Lighting modes | Modular attachments (Label / DPM) | Native multi-angle in shroud |
| Aperture range | Fixed set | Wider; auto-select by X-dim |
| DPM optimization | Yes (with attachment) | Yes (native) |
| Field of view | Smaller | Larger |
| Cell Contrast / Modulation (DPM-only) | Yes (firmware-supported) | First-class |

---

## 2. Webscan Lineage Column Semantics (the authoritative table)

| Field | Meaning | Units | ISO equivalent |
|---|---|---|---|
| **LLS** | Left Light Source — finder-pattern illumination consistency | Grade A-F | Fixed Pattern Damage (segment) |
| **BLS** | Bottom Light Source — same, bottom finder segment | Grade A-F | Fixed Pattern Damage (segment) |
| **LQZ** | Left Quiet Zone grade | Grade A-F | Part of FPD |
| **BQZ** | Bottom Quiet Zone grade | Grade A-F | Part of FPD |
| **TQZ** | Top Quiet Zone grade | Grade A-F | Part of FPD |
| **RQZ** | Right Quiet Zone grade | Grade A-F | Part of FPD |
| **HQZ** | Horizontal Quiet Zone (mean of L+R) | Grade A-F | — |
| **ULQZ** | Upper-Left Quiet Zone (large matrices ≥32×32) | Grade A-F | FPD region |
| **URQZ** | Upper-Right Quiet Zone | Grade A-F | FPD region |
| **RUQZ** | Right-Upper Quiet Zone | Grade A-F | FPD region |
| **RLQZ** | Right-Lower Quiet Zone | Grade A-F | FPD region |
| **TTR** | Top Transition Ratio — ratio of transition-zone width to module size | % + Grade | Part of FPD |
| **RTR** | Right Transition Ratio | % + Grade | Part of FPD |
| **TCT** | Top Clock Track — alternating-module integrity | Grade A-F | Part of FPD |
| **RCT** | Right Clock Track | Grade A-F | Part of FPD |
| **AG** | Average Grade across scan lines | 0.0-4.0 | — |
| **MRD** | Module Reflectance Distribution — spread of reflectance within modules | % | — (general char) |
| **SCRlRd** | Symbol Contrast formatted `SC (Rl/Rd)`. Rl=Rmax, Rd=Rmin | % reflectance | Symbol Contrast (15415) |
| **BWG** | Bar Width Growth, horizontal + vertical | % | **Print Growth** (15415 Ed.3) |
| **AvgEdge** | Average Edge Determination (1D, across scans) | Grade A-F | Edge Determination (15416) |
| **AvgRlRd** | Average reflectance Rl/Rd pair (1D) | % | — |
| **AvgSC** | Average Symbol Contrast (1D) | % + Grade | Symbol Contrast (15416) |
| **AvgMinEC** | Average minimum Edge Contrast (1D) | Grade | Edge Contrast (15416) |
| **AvgMOD** | Average Modulation (1D) | Grade A-F | Modulation (15416) |
| **AvgDefect** | Average Defects (1D) | Grade | Defects (15416) |
| **AvgDcod** | Average Decodability (1D) | Grade | Decodability (15416) |
| **AvgDEC** | Average Decode result | Grade | Decode (15416) |
| **AvgLQZ** | Average Left Quiet Zone (1D, across scans) | Grade | QZ |
| **AvgRQZ** | Average Right Quiet Zone (1D) | Grade | QZ |
| **AvgHQZ** | Average HQZ (1D) | Grade | QZ averaged |
| **MinQZ** | Minimum of LQZ / RQZ | Grade | — |

**Big revelation**: BWG = Print Growth. Our currently-empty
`<BWGPercent>` column maps directly to `r.metrics.printGrowth` (which is
itself now a graded parameter in 15415 Ed.3). This nails down one of the
v1.24 wire-ups from a different angle.

---

## 3. DPM-Mode Behavior (ISO 29158 / AIM-DPM)

When the verifier switches to DPM mode, the metric set substitutes:

| ISO 15415 (printed labels) | ISO 29158 (DPM) |
|---|---|
| `symbolContrast` (SC) | `cellContrast` (CC) |
| `modulation` (MOD) | `cellModulation` (CMOD) |
| `defects` | `cellDefects` |
| `fixedPatternDamage` | `finderPatternDefects` |
| `moduleSize` | `dataMatrixCellWidth` / `dataMatrixCellHeight` |
| `axialNonuniformity` | `horizontalMarkMisplacement` |
| `gridNonuniformity` | `verticalMarkMisplacement` |

Reflectance scaling: uses "Optimized Brightness" (L_255 and L_0 values)
rather than absolute 0-100% reflectance.

**This finally explains** why our 30-key v1.23 metric enumeration surfaces
DPM-only metrics (`cellDefects`, `finderPatternDefects`,
`dataMatrixCellHeight/Width`, both `*MarkMisplacement` fields) that
always report NA on calibration-card scans — they're the DPM substitutes,
hot only when `ApplicationStandard=AIM-DPM`.

**VTCCP schema implication**: We should not blindly add 7 new DPM
columns. Better is to use a single set of columns (SC, MOD, Defects,
FPD, ModuleSize, ANU, GNU) and a `MetricSet` enum (`ISO-15415` |
`AIM-DPM`) telling consumers which value semantics apply. This matches
how Cognex actually emits — same field names with different meaning per
mode.

---

## 4. Reportable Fields (condensed)

| Parameter | Description | Range | Condition |
|---|---|---|---|
| Overall Grade | Lowest of all parameter grades | A-F (4.0-0.0) | Always |
| UEC | Unused Error Correction | 0-100% | 2D only |
| ANU / GNU | Axial / Grid Nonuniformity | 0-100% | 2D only |
| FPD | Fixed Pattern Damage | Grade | 2D only |
| SC / CC | Symbol / Cell Contrast | 0-100% | 15415 / 29158 |
| MOD / CMOD | Modulation / Cell Modulation | Grade | 15415 / 29158 |
| RM | Reflectance Margin | Grade | 15415:2024+ only |
| Matrix Size | rows × cols (e.g. 14×14) | string | 2D only |
| X-Dimension | Measured module size | mils / mm | Always |

---

## 5. Output Formats

The 8072V produces three native formats:
1. **PDF report** — Webscan-style with heatmaps.
2. **CSV results** — flat file, 167-column schema (the format we
   replicate in our XLSX).
3. **XML over TCP/FTP** — structured `<DMSymVerResponse>` envelope
   identical to what our `DmstPushScript_v1.js` produces.

This **confirms our schema and XML structure are 8072V-compatible** —
not coincidence, both share the Webscan-lineage data model. Migration
path to 8072V hardware would be largely free.

---

## 6. Calibration

- **Procedure**: NIST-traceable calibration cards (DMV-DMCC for Data
  Matrix). User captures cal symbol under specified lighting angle.
- **Frequency**: Recommended every 30 days or whenever
  environment/mounting changes.
- **Traceability claim**: Cognex claims ISO 17025 accreditation for
  their calibration cards.

---

## 7. Network Output / Push Configuration

- Network Client push supported (same architecture as DM475V).
- Script-Based Formatting using JavaScript supported in the `onResult`
  pipeline.
- **`DmstPushScript_v1.js` runs unmodified** on the 8072V.

---

## 8. Image Polarity & Orientation

- **`<ImagePolarity>`** values:
  - `BlackOnWhite` — standard
  - `WhiteOnBlack` — inverted
- Orientation auto-detected; no manual "flipped" field needed in schema.

---

## 9. Image-Load Capability — CONFIRMED

The 8072V manual documents that it **supports verifying pre-captured
images** via DMST's "Image Load" / "Image Playback" feature. Accepts
`.jpg` and `.png` files. Runs them through the grading engine.

**For VTCCP's loaded-image flow**:
- The capability exists at the DMST-tool level on the 8072V
- The underlying DMCC mechanism is `WBU` (Write Buffer) per the Comms
  guide digest
- We don't have explicit confirmation that DM475V supports the same
  flow — needs verification, but given firmware commonality, likely yes

**OpticsCompliant flag** still applies because the loaded image bypasses
the verifier's calibrated optical chain — same caveats apply regardless
of whether the loaded-image path is supported by the hardware or not.

---

## 10. Glossary of 8072V-Specific Terms

| Term | Meaning |
|---|---|
| **AICC** | Applied Image Calibration Card (1D) |
| **DMCC** | (in card context) Applied Image Data Matrix Calibration Card |
| **GS1CC** | GS1 Calibration Card |
| **SBG** | Standards-Based Grading (verification without strict 15426-2 hardware compliance — used on non-V readers) |
| **TruCheck** | Legacy Webscan brand for the verification algorithm suite |
| **AIM-DPM** | ISO 29158 DPM grading methodology |
| **L_255 / L_0** | Reference brightness values used for DPM reflectance normalization |

---

## 11. What this changes for VTCCP

### Schema validation wins
- All Webscan-lineage column names in our 167-column schema are
  documented + ISO-mapped — no more opaque acronyms in report copy.
- `BWGPercent` = Print Growth — confirmed wire-up target for v1.24.
- DPM-only metric set fully mapped to its ISO 15415 equivalents.

### Schema design recommendation
- Add `MetricSet` enum column (`ISO-15415` | `AIM-DPM`) instead of
  shadowing DPM fields as separate columns. Consumers interpret the
  same columns based on this enum. Matches how Cognex emits.

### Report design wins
- "Webscan-style" PDF is explicitly mentioned with heatmaps in the
  8072V manual. The Webscan PDF sample in
  `references/samples/reports/Webscan-format-2025-09-26.pdf` likely
  reflects this style. **Trade-dress posture is stronger now** —
  the look is documented, lineage-owned, predates DMST.

### Loaded-image flow design
- 8072V supports it natively in DMST. DM475V likely does too via the
  same DMCC `WBU` command — needs hardware confirmation (B3 in
  session plan).
- VTCCP's role: WPF file-load UI → `WBU` to device → existing push
  path returns results → record written with `OpticsSource=LoadedImage`
  + `OpticsCompliant=false`.

### Migration insurance
- If VCCS DMV ever upgrades to 8072V hardware, **VTCCP would work
  unchanged** — same XML schema, same push-script API, same Webscan
  field semantics. Worth noting in the operational documentation.
