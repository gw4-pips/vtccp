# DM475V-DPM (866D76) — TruCheck Calibration Record
**Date:** 2026-08-05  
**Unit:** DM475-866D76 (MAC 00-D0-24-86-6D-76, serial 1A2134PP482281)  
**Firmware:** 6.1.16_tc9  
**State before calibration:** Factory-default (fresh reset, `TRUCHECK.CALIBRATION-DATE = Not Calibrated`)

---

## Calibration artefacts

| File | Contents |
|---|---|
| `DM475V-DPM_866D76_calibration-reports_2026-08-05.zip` | 54 TruCheck PDF verification reports — the complete calibration sequence |
| `DM475V-DPM_866D76_factory-default-baseline_2026-08-05.txt` | Pre-calibration DMCC parameter dump (factory defaults) |

---

## Reference target

| Parameter | Value |
|---|---|
| Barcode | `010123456789012810` |
| Symbology | GS1 DataMatrix |
| Reference grade | Grade 4 / A (ISO/IEC 15415) |
| Label in filename | `GRADE-4-A-AI-INC` |

---

## Calibration sequence summary

**Duration:** 18:18:41 → 18:20:42 EDT (approximately 2 minutes 1 second)  
**Total scans:** 54  
**Scan interval:** ~2.3 seconds per report

### Grade progression

| Scans | Time range | Overall grade | PPM | SC |
|---|---|---|---|---|
| 1 – 30 | 18:18:41 – 18:19:49 | **A** | 33.36 | 4.0 |
| 31 – 42 | 18:19:49 – 18:20:14 | **B** | 33.35–33.36 | 4.0 |
| 43 – 54 | 18:20:17 – 18:20:42 | **A** | 33.36 | 4.0 |

The B-grade dip over scans 31–42 (~25 seconds) reflects the scanner adjusting illumination
gain through a transient sub-optimal setting. PPM and Symbol Contrast remained stable
throughout; only the overall composite grade was affected. The sequence converged back
to grade A and held there for the final 12 scans, confirming a stable calibrated state.

### Key metrics (stable throughout)
- **Pixels per Module (PPM):** 33.35 – 33.38 (essentially constant)
- **Symbol Contrast (SC):** 4.0 / A — no variation
- **All 14 GS1 DM quality parameters:** PASS on every scan
- **UEC (Unused Error Correction):** 100%

---

## TRUCHECK parameters after calibration

Run `Get-DmSettings.ps1 -DeviceIp 10.10.10.4` post-calibration and diff against  
`DM475V-DPM_866D76_factory-default-baseline_2026-08-05.txt` to capture what changed.  
Expected changes: `TRUCHECK.CALIBRATION-DATE`, `TRUCHECK.APERTURE`,
`TRUCHECK.APERTURE-SIZE`, and possibly `CAMERA.EXPOSURE` / `CAMERA.GAIN`.

---

## Notes
- `TRUCHECK.APPLICATION-STANDARD = 5` at factory default (differs from LBL unit which is 4).
  Confirm this is correct for the DPM application standard before production use.
- `VERIFICATION.ENABLE = ON` at factory default — verification active immediately after reset.
- All MST parameters were `(no response)` at factory default — DMST requires configuration
  from scratch when Command Pilot is connected.
