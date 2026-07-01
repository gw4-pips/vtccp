# Lighting Investigation Plan — DM475V-DPM
**Version**: 1 — 2026-07-01
**Device**: DM475V-DPM (10.10.10.4, fw 6.1.16_sr4)
**Goals**:
1. Identify the DMCC key(s) that control illumination geometry on the DM475V-DPM
2. Empirically compare 45Q vs 30Q grading results under ISO 15415

---

## Background

The DM475V-LBL (10.10.10.7) has fixed 45Q optics and cannot produce any other illumination
geometry. The DM475V-DPM (10.10.10.4) has multi-angle LED hardware supporting all ISO 29158
illumination types: 30Q, 30T (N/S/E/W), 90D, and 45Q.

The DMCC key for lighting geometry is not yet in the confirmed key inventory. The
`PUT /vs.cfg` config sync is AES-encrypted and not readable from Wireshark payload inspection.
However, individual DMCC SET commands issued by DMST are transmitted in plaintext over TCP
and will be visible in a Wireshark capture on port 44444 or port 23.

The `<Lighting>` field in push XML codes.xml and `<FormalGrade>` confirm which illumination
the device actually used for a given verification — independent of what DMST reports.

---

## Prerequisites

- DMST open and connected to DM475V-DPM (10.10.10.4)
- TC (TruCheck) panel open
- Wireshark running on the DMST workstation (10.10.10.19)
- A printed label symbol on label stock — same physical label used for all scans in Goal 2
- VTCCP push listener active to capture full push XML per scan (or save raw push XML manually)

---

## Goal 1 — Find the DMCC lighting control key

### Wireshark filter
```
host 10.10.10.4 and (port 44444 or port 23)
```
This catches both DMCC channels. Display filter after capture:
```
tcp and (tcp.dstport == 44444 or tcp.dstport == 23 or tcp.srcport == 44444 or tcp.srcport == 23)
```
Follow TCP Stream on each connection to read plaintext DMCC commands.

### Capture session — one action at a time

| Step | Action in DMST TC UI | What to capture |
|---|---|---|
| G1-1 | Baseline — note current setting (should be ISO 15415, 45Q) | No action, just confirm starting state |
| G1-2 | Switch Grading Standard: ISO 15415/6 → ISO 29158 | Expect `SET TRUCHECK.GRADING-STANDARD 1` — confirms DMCC plaintext pattern |
| G1-3 | Switch lighting option: first available → next option, one step | Expect a DMCC SET command with the lighting key name — **this is the target** |
| G1-4 | Cycle through each remaining lighting option, one at a time | One DMCC SET per step — record key name and enum value for each option |
| G1-5 | Switch Grading Standard back: ISO 29158 → ISO 15415/6 | Confirm TRUCHECK.GRADING-STANDARD 0 roundtrip |

### What to record per lighting option

For each illumination setting change, record:
- The DMCC key name (e.g., `TRUCHECK.LIGHTING` or whatever appears)
- The DMCC value/enum for that option
- The DMST UI label (e.g., "30Q", "30T North", "90D")
- Whether `PUT /vs.cfg` also fires (expected — belt-and-suspenders config sync)

### Expected outcome

A complete map of `{UI label → DMCC key → DMCC enum value}` for all illumination
options on the DM475V-DPM. This closes the last major gap in the TC configuration key
inventory and enables VTCCP to set lighting geometry directly without DMST.

### Fallback — if no plaintext DMCC SET appears

If DMST writes lighting changes exclusively through `PUT /vs.cfg` (AES blob) with no
accompanying plaintext DMCC command, the key is not directly accessible via Wireshark
payload inspection. In that case, the path forward is:
- Check whether `GET TRUCHECK.LIGHTING` (or equivalent) is a valid DMCC read command
  by issuing it manually via raw TCP to port 44444 and observing the response
- Attempt `SET TRUCHECK.LIGHTING 0` (and variants) directly and observe `<Lighting>`
  in the resulting push XML to confirm the device honored the command

---

## Goal 2 — 45Q vs 30Q empirical comparison under ISO 15415

### Setup

- Symbol: one printed label, fixed position, do not move between scans
- Standard: ISO 15415 for all scans (switch to 29158 only to unlock 30Q, then evaluate)
- Device: DM475V-DPM only
- Capture: full push XML for every scan (codes.xml body or VTCCP push listener)
- Minimum: 3 scans per condition for basic repeatability; 5 recommended

### Scan sequence

| Scan set | Grading Standard in TC | Lighting | Notes |
|---|---|---|---|
| A (×3–5) | ISO 15415 | 45Q | Baseline — conformant ISO 15415 condition |
| B (×3–5) | ISO 29158 | 30Q | Non-conformant for label stock per ISO 15415 — empirical test condition |
| A2 (×3) | ISO 15415 | 45Q | Repeat baseline — confirm no symbol or device drift |

Do not physically move the symbol between scan sets. Change only the TC standard/lighting
setting between set A and set B.

### Fields to compare per scan

From push XML / codes.xml `<trucheck_verification_result>` block:

| Field | ISO significance |
|---|---|
| `<OverallGrade>` / numeric | Primary outcome |
| `<FormalGrade>` | Confirms lighting condition used (`45Q` or `30Q`) |
| `<SymbolContrast>` / SC numeric | Most illumination-sensitive parameter |
| `<Modulation>` / MOD numeric | Second most illumination-sensitive |
| `<UnusedErrorCorrection>` / UEC | Decode robustness — less illumination-sensitive |
| `<AxialNonUniformity>` / ANU | Geometry — should be illumination-independent |
| `<GridNonUniformity>` / GNU | Geometry — should be illumination-independent |
| `<FixedPatternDamage>` / FPD | Structure — should be illumination-independent |
| `<Lighting>` | Ground truth confirmation of actual illumination used |

### Analysis

Compute mean and range for each parameter across the scan set.
Compare set A vs set B:

- **Within ISO inter-verifier tolerance** (±0.5 numeric grade): difference is
  empirically inconsequential for this symbol and substrate
- **Beyond tolerance**: illumination difference is material — the competing
  manufacturer's "inconsequential" claim does not hold for this symbol type

Expected pattern: SC and MOD will show the largest delta (most illumination-sensitive).
ANU, GNU, FPD should show minimal delta (geometry-derived, illumination-independent).
If overall grade difference is driven by SC/MOD delta, that is the illumination effect.

### What to record

For each scan: full push XML saved as a numbered file.
Summary table: one row per scan, all fields above, condition (A or B) noted.

---

## Output documents

| Document | Content |
|---|---|
| `lighting-dmcc-key-map.md` | Goal 1 result — DMCC key name, enum values, UI labels for all DM475V-DPM illumination options |
| `lighting-45q-vs-30q-results.md` | Goal 2 result — scan-by-scan data table, parameter comparison, conclusion on materiality |

Both to be filed in `vtccp/architecture/` on completion.

---

## Relationship to VTCCP implementation

Once Goal 1 is complete:
- Add confirmed lighting key to `DmccCommand.cs`
- Add `GetLightingAsync` / `SetLightingAsync` to `DeviceSession.cs`
- Scope lighting control into the device configuration panel (pending TC window screenshots)

Goal 2 finding informs D1 report design:
- The `<Lighting>` component of the formal grade string should be treated as a compliance
  field in the D1 report, not decorative metadata
- Any lighting condition other than the standard-mandated value warrants a flagged note
  in the report (analogous to the existing `CalibrationWarning` for `FieldCalibrated=false`)
