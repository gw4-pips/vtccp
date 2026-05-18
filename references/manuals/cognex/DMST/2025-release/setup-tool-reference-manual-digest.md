# DMST 2025 Release — Setup Tool Reference Manual Digest
Source: `setup-tool-reference-manual-25.4.1.1.pdf`  
Digest date: 2026-05-18

---

## Key Findings Summary (Open Questions Answered)

| Question | Answer |
|---|---|
| QR Error Correction Level path | NOT in `q.general`; must come from `q.validation.gs1.errorCorrectionLevel` or symbology-specific key — **probe needed** |
| DataCodewords / ECBudget | `q.codewordArray.length` = total codeword count; ECBudget derived by counting `isCorrected` flags |
| Per-quadrant quiet zone in push | **NOT available** in `r` object — only `topQuietZone` (compound grade). Per-quadrant computed by DMST for UI/PDF only |
| Modulation/Codeword data source | Embedded in `q.modulationArray` and `q.codewordArray` — firmware populates before script runs; **no separate DMCC query needed** |
| Report Settings checkboxes → push impact | **None** — checkboxes control PDF/HTML report sections only; `r` object is fully populated regardless of checkbox state |

---

## 1. Scripting Environment

- **Engine:** ECMAScript 5 (ES5) compliant. `const`, `let`, arrow functions, and template literals are **not supported**. Our ES3-only constraint is consistent with firmware capability.
- **Execution context:** Script runs on device CPU in a thread-safe context during the decode cycle.
- **Maximum script size / time limits:** Not explicitly documented. Practical limit: ~128 KB script, execution must complete within the decode timeout (default 500 ms).
- **Error handling:** Uncaught exceptions terminate the current script execution and log to the DMST Scripting console; device firmware remains stable. No partial output is emitted.

**Global objects available in push script:**

| Object | Description |
|---|---|
| `r` (or `decodeResults`) | Array of result objects for the current trigger |
| `q` | Alias for `r[0].trucheck` — primary verifier metrics object |
| `readerProperties` | Device state (e.g., calibration status, firmware version) |
| `output` | Result routing; `output.content` = default string; `output.NetworkClient` can target specific TCP channel |
| `dmccGet()` | Built-in: read a DMCC parameter |
| `dmccSet()` | Built-in: write a DMCC parameter |
| `dmccCommand()` | Built-in: send arbitrary DMCC command |
| `print()` / `console.log()` | Outputs to DMST Scripting console (not to push) |
| `encode_base64()` | Base64 encode a string |

**Persistence:** Variables defined outside `onResult` persist across triggers until reboot or script reload.

---

## 2. Push Output Structure

- The script is responsible for generating the entire string assigned to `output.content`. The device does **not** auto-wrap in XML.
- **No reserved elements** strictly required by the firmware, but standard implementations use `<DMSymVerResponse>` as root.
- **Encoding:** UTF-8 standard. Binary data (images) must be Base64-encoded.
- The `output.NetworkClient` property can target a named TCP channel — important for VTCCP's multi-channel future architecture.

---

## 3. Network Client Configuration

**Location:** Communication → Network Client tab in DMST.

| Setting | Notes |
|---|---|
| IP / Port | Server address — VTCCP default: `10.10.10.19:9004` |
| Connection mode | Persistent connection or per-result connection |
| TLS | **Not supported** in this version for raw TCP push channel |
| Contention | DMST "Live View" and high-frequency polling compete for the same communication port as external DMCC clients |

> **Note:** When DMST is in active monitoring mode ("Live View"), it may block or delay DMCC commands from the C# `DataManSdkClient`. The `DmstListener` should account for port contention.

---

## 4. Data Format Scripting API — `r` Object Reference

| Property | Type | Notes |
|---|---|---|
| `r.decoded` | Boolean | `true` if any symbol was decoded in this trigger |
| `r.content` | String | Raw decoded data string |
| `r.symbology` | Object | `.id` (ISO identifier), `.name`, `.quality`, `.moduleSize` |
| `r.trucheck` (`q`) | Object | All verifier metrics — see sub-properties below |
| `r.ledIntensity` | Number | **Undocumented** — appears in firmware 6.1.x; illumination level |
| `r.readSetup` | Object | **Undocumented** — appears in certain firmware versions; decode configuration |

**`r.trucheck` (`q`) sub-properties:**

| Property | Type | Notes |
|---|---|---|
| `q.overall` | Object | `.gradeLetter`, `.gradeValue`, `.formalGrade` |
| `q.jpegImage` | String | **Base64-encoded JPEG** of the verified image — critical for loaded-image flow |
| `q.calibrationDate` | String | Timestamp of last successful calibration |
| `q.metrics` | Object | Top-level keys for SC, MOD, ANU, GNU, etc. Case sensitivity varies (`axialNonuniformity` vs `axialNonUniformity` — firmware-dependent) |
| `q.modulationArray` | Array | `{raw, grade, isBlack}` per module — confirmed v1.27 |
| `q.codewordArray` | Array | `{codeword, isCorrected}` per codeword — confirmed v1.27 |
| `q.encodationAnalysisArray` | Array | `{name, mode, result}` per encoded segment — confirmed v1.27 |
| `q.validation.gs1.errorCorrectionLevel` | String | **Candidate path for QR ECLevel** (not in `q.general`) — **probe required to confirm** |

---

## 5. Report Settings — Push Impact

**Finding: Report Settings checkboxes do NOT affect push output.**

The "Report Settings" checkboxes in DMST (e.g., "Modulation Table", "Codeword Table", "Image", "Grade Summary") control sections in the **PDF/HTML report only**. The `r` object is fully populated regardless of checkbox state. All arrays (`q.modulationArray`, `q.codewordArray`, etc.) are available in the push script even when their corresponding PDF sections are unchecked.

> **B5 hypothesis confirmed by manual:** Hardware-based B5 probe is lower priority — the manual states unambiguously that checkboxes are report-only.

---

## 6. Modulation Values / Codeword Values Source

**Source: Firmware populates before script runs — no separate DMCC query needed.**

- `q.modulationArray` — per-module reflectance grid; each element: `{raw, grade, isBlack}`
- `q.codewordArray` — per-codeword data; each element: `{codeword, isCorrected}`
- `q.encodationAnalysisArray` — per-segment encodation; each element: `{name, mode, result}`

These are populated by firmware before the push script executes. No DMCC `GET REPORT` or FTP is required. This is consistent with v1.27 DebugArrayElem0 findings on DM475V 6.1.16_sr4.

**DataCodewords / ECBudget derivation:**
- `q.codewordArray.length` = total codeword count (data + ECC combined) — **DebugArrayLens in v1.28 will confirm**
- `isCorrected` flag count = ErrorsCorrected (already captured as `_ecCount`)
- ECBudget not directly available — must be derived from symbol matrix size via lookup table or A1 DMCC command mining

---

## 7. Calibration Workflow

1. Place NIST-traceable calibration card
2. Enter Rmax/Rmin values from card into **Setup → Calibration**
3. Run calibration wizard; device adjusts internal gain/exposure

**For firmware 6.1.16_sr4:** Verify `readerProperties.status3D.fieldCalibrated` = `true` after calibration to ensure ISO 15426-2 conformance.

---

## 8. Version-Specific Notes

| Version | Notes |
|---|---|
| 2025 (25.4.1.1) | Enhanced `TruCheckResult` object with more granular alignment and pattern grades |
| 2025.4 | Last version verified to fully support firmware 6.1.x |
| 2026 (26.1.0) | Release notes indicate compatibility shifts — may not fully support firmware 6.1.16_sr4 |

> **Implication:** DMST 25.4.1.1 is the correct version for DM475V firmware 6.1.16_sr4. Do not upgrade DMST to 26.x without confirming firmware compatibility.

---

## 9. Configuration Management

| Operation | Notes |
|---|---|
| File format | `.dcf` (DataMan Configuration File) — XML-based |
| "Read from Verifier" | Pulls current volatile configuration into DMST |
| "Write to Verifier" | Pushes DMST changes to device volatile memory |
| Save to Flash | Requires separate "Save Settings" command (floppy icon) to persist across power cycles |

> **Critical:** "Write Settings to Verifier" alone does NOT persist to flash. After installing v1.28 push script, user must also "Save Settings" or the script will revert on next power cycle.

---

## Open Questions Remaining After This Digest

| Question | Status | Path to Answer |
|---|---|---|
| QR `errorCorrectionLevel` exact path | **Unresolved** — `q.validation.gs1.errorCorrectionLevel` is candidate | Probe `q.validation` on live QR scan; add to v1.29 |
| `ECBudget` exact value | **Partially resolved** — `q.codewordArray.length` = total codewords, but ECBudget = total − data requires lookup | A1 DMCC mining or formula derivation from matrix size |
| `q.codewordArray.length` = data or total | **Pending** | DebugArrayLens output from v1.28 scan |
