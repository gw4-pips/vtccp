# Cross-Manual Synthesis
Manuals: DMCC Comms Guide + DMV-8072V Reference + DMST 2025 Release  
All version 25.4.1.1. Synthesis date: 2026-05-18.

---

## 1. Open Questions — Resolved / Partially Resolved

### QR Error Correction Level (L/M/Q/H)

| Source | Finding |
|---|---|
| DMST digest | Candidate: `q.validation.gs1.errorCorrectionLevel` |
| DMCC guide | Candidate: `r.trucheck.symbols[0].errorCorrectionLevel`; also AIM ID in `r.symbology.id` |
| v1.27 live scan | `q.general` confirmed 7 keys only — ECLevel absent |
| **Status** | **Unresolved.** Two probe targets identified. Add `DebugSymbols` + `DebugValidationGS1` to v1.29. |

### DataCodewords / ECBudget

| Source | Finding |
|---|---|
| DMCC guide | `r.trucheck.symbols[0].dataCodewords` and `.ecCodewords` — documented |
| DMST digest | `q.codewordArray.length` = total codewords (data+ECC combined); ECBudget = total − data |
| v1.27 live scan | `q.symbols` did NOT appear in 47-key enumeration — may be absent in firmware 6.1.16_sr4 |
| **Status** | **Partially resolved.** `DebugSymbols` probe (v1.29) will confirm whether `q.symbols` exists on this firmware. `DebugArrayLens` (v1.28) will give total array lengths. |

### Per-Quadrant Quiet Zone Grades (ULQZ/URQZ/RUQZ/RLQZ)

| Source | Finding |
|---|---|
| DMCC guide | Only in `GET RESULT.XML` (not in push `r` object) |
| DMV-8072V digest | DMV-8072V exposes ULQZ/URQZ/RUQZ/RLQZ for ≥32×32 DM — newer firmware generation |
| v1.27 probe | `q.topQuietZone` / `q.rightQuietZone` = flat `{grade:X, numericGrade:8.8}` — per-quadrant not in push |
| **Status** | **Resolved as firmware-generation limit.** DM475V 6.1.16_sr4 does not expose per-quadrant QZ in push. Workaround: `dmccGet("RESULT.XML")` inside push script — substring parse for `<ULQZ>` etc. This adds ~50 ms per scan; defer until VCCS report requirements confirmed. |

### Modulation Values / Codeword Values Source

| Source | Finding |
|---|---|
| DMCC guide | Available in `GET RESULT.XML` with Extended Verification; also `q.modulationArray` / `q.codewordArray` in push |
| DMST digest | Push arrays confirmed primary source; no separate DMCC query needed |
| v1.27 probe | Array element shapes confirmed: mod={raw,grade,isBlack}, cw={codeword,isCorrected}, ea={name,mode,result} |
| **Status** | **Resolved.** B7 implementation uses push arrays directly. DebugArrayLens (v1.28) will confirm lengths. |

---

## 2. Architecture Decisions Confirmed by Manuals

### Push vs DMCC for image capture
**Decision: Use push.** `q.jpegImage` (Base64 JPEG) is in the push `r` object and fires synchronously with the result. `GET IMAGE` via DMCC is a separate, potentially unsynchronized call. The loaded-image flow should use `IMAGE.LOAD` → `IMAGE.REPLAY` → push fires with `q.jpegImage`.

### Configuration persistence
**Decision: Always CONFIG.SAVE after script install.** "Write Settings to Verifier" writes to volatile memory only. Without `CONFIG.SAVE` (or the floppy icon in DMST), the v1.28 push script reverts on power cycle.

### Report Settings checkboxes → no push impact
**Confirmed.** DMST Report Settings checkboxes control PDF/HTML report sections only. The `r` object is fully populated regardless. B5 hardware probe is informational only.

### ES3-only script constraint confirmed
DMST manual confirms ES5 is the engine version. ES3 is a conservative subset of ES5. Our ES3-only constraint is correct — it will execute on the device without issues.

### DMST version pinned at 25.4.1.1
DMST 2025.4 is the last version confirmed to fully support firmware 6.1.16_sr4. DMST 26.x has compatibility shifts — do not upgrade.

---

## 3. New v1.29 Probe Targets

Priority-ordered based on impact to VTCCP output:

| Priority | Probe | Target | Answers |
|---|---|---|---|
| 1 | `DebugSymbols` | `q.symbols`, `q.symbols[0]`, `.dataCodewords`, `.ecCodewords`, `.errorCorrectionLevel` | DataCodewords, ECBudget, QR ECLevel — all three high-value wires |
| 2 | `DebugValidationGS1` | `r.validation`, `r.validation.gs1`, `.errorCorrectionLevel` | Alternate QR ECLevel path; also reveals full GS1 validation object shape |
| 3 | `DebugAimId` | `r.symbology.id` full string for QR | AIM identifier encodes EC level and version for QR; can derive ECLevel from byte 3 |

These three probes may resolve all three remaining unknowns (QR ECLevel + DataCodewords + ECBudget) in a single v1.29 scan.

---

## 4. IMAGE.LOAD / IMAGE.REPLAY — Full Integration Sequence

For VTCCP's loaded-image flow (`DataManSdkClient.LoadAndVerifyImageAsync`):

```
1. client.SendCommand("SET TRIGGER.TYPE 1")     // manual trigger
2. client.SendCommand($"IMAGE.LOAD {size} {b64}")  // transfer image to WBU
3. client.SendCommand("IMAGE.REPLAY")           // decode against buffered image
4. DmstListener receives push result            // r fires same as hardware trigger
   - q.jpegImage = Base64 of the loaded image (confirmed from DMST digest)
   - AverageGrade = "X" for loaded images       // confirmed v1.26/v1.27 scans
   - All optical measurement grades = sentinels // confirmed v1.27 QR loaded scan
```

The `OpticsSource = "LoadedImage"` flag in `VerificationRecord` should be set whenever `AverageGrade == "X"`.

---

## 5. Webscan Lineage Field Mapping (DMV-8072V → VTCCP)

From the DMV-8072V digest — Webscan 167-column schema vs. current VTCCP push XML field names:

| Webscan CSV Column | Push XML Field (v1.28) | ISO Standard |
|---|---|---|
| LLS | `<LLSGrade>` | ISO 15415 FPD |
| BLS | `<BLSGrade>` | ISO 15415 FPD |
| TCT | `<TCTGrade>` | ISO 15415 FPD |
| RCT | `<RCTGrade>` | ISO 15415 FPD |
| TTR | `<TTRGrade>` | ISO 15415 FPD |
| RTR | `<RTRGrade>` | ISO 15415 FPD |
| TQZ / RQZ | `<ULQZGrade>` / `<URQZGrade>` (aggregate via topQuietZone) | ISO 15415 QZ |
| LQZ / BQZ | `<RUQZGrade>` / `<RLQZGrade>` (aggregate via rightQuietZone) | ISO 15415 QZ |
| BWG | `<BWGrade>` | ISO 15415:2024 |
| MRD | `<MRDGrade>` | Internal |
| SC / Rd / Rl | `<SCGrade>` / `<MinReflectance>` / `<MaxReflectance>` | ISO 15415 SC |

---

## 6. DPM Mode Impact on VTCCP

When the DM475V operates in DPM mode (ISO 29158 / AIM-DPM), the following substitutions apply in push output:

| Standard Field | DPM Replacement | Notes |
|---|---|---|
| `symbolContrast` (SC) | `cellContrast` (CC) | Different grade formula |
| `modulation` (MOD) | `cellModulation` (CMOD) | DPM-specific |
| `fixedPatternDamage` (FPD) | `finderPatternDefects` | Different calculation |
| `axialNonuniformity` | `horizontal/verticalMarkMisplacement` | DPM geometry |

VTCCP's `DmstPushParser` should detect DPM mode (via `r.general.illuminationType` or similar) and route to DPM-specific field names. This is a **parser extension needed** — not currently handled.

---

## 7. Network Architecture Confirmed

```
DM475V (10.10.10.7)
  ├── TCP 23  ←→ DataManSdkClient.cs (DMCC: trigger, IMAGE.LOAD, IMAGE.REPLAY, CONFIG.SAVE)
  ├── TCP 9004 → DmstListener.cs (push results: XML string from onResult script)
  └── TCP 23  ←→ DMST GUI (competing port — avoid simultaneous DMST Live View + VTCCP DMCC)
```

Port 23 (DMCC) is shared between VTCCP's `DataManSdkClient` and DMST's Live View mode. Running both simultaneously may cause contention. Recommend: VTCCP takes exclusive DMCC ownership during IMAGE.LOAD/REPLAY sequences.
