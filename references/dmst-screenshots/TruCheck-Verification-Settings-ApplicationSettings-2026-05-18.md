# TruCheck Verification Settings — Application Settings Panel

**Screenshot**: `TruCheck-Verification-Settings-ApplicationSettings-2026-05-18.png`
**Captured**: 2026-05-18
**Context**: DMST 26.1.0 connected to DM475-63530E-PIPS-Verif-Lab [10.10.10.7]

---

## What the screenshot shows

TruCheck Verification Settings dialog → **Application Settings** panel.

Left nav items (all panels confirmed):
- **Application Settings** ← this screenshot
- Calibration Settings
- Trending Settings
- User Information
- Report Settings
- Navigation

---

## Controls — complete inventory

### Grading Standard (radio group)
| Option | State |
|---|---|
| **ISO 15415/6** | ● selected |
| **ISO 29158 (AIM-DPM)** | ○ not selected |

### Application Standard
| Control | Value / State |
|---|---|
| **Select Standard** (dropdown) | Custom |
| **Dot Peen** (checkbox) | ☐ unchecked |
| **Min X Dimension (mils)** | 8 |
| **Max X Dimension (mils)** | 30 |
| **Overall Pass Grade** (dropdown) | 1.5 |
| **Advanced Custom Application Standard** | ⊙ collapsed (expandable) |

### Data Format Check (radio group)
| Option | State |
|---|---|
| **None** | ● selected |
| **GS1** | ○ not selected |
| **HiBCC** | ○ not selected |
| **ISO 15434** | ○ not selected |

### Aperture Setting (dropdown)
| Control | Value |
|---|---|
| **Aperture Setting** | Auto 50%/80% |

---

## Mapping to push-XML fields

| UI Control | Push XML field | Live DM value | Live QR value |
|---|---|---|---|
| Grading Standard = ISO 15415/6 | `<GradingStandard>` | ISO 15415:2011 | ISO 15415:2011 |
| Select Standard = Custom | `<ApplicationStandard>` | Custom | Custom |
| Data Format Check = None | drives `<ApplicationPass>` | Fail (Quality) | Pass |
| Overall Pass Grade = 1.5 | `<MinPassGrade>` | **NA** ← mismatch, see below | **NA** |
| Aperture Setting = Auto 50%/80% | `<ApertureRef>` | 07 | 07 |

---

## Key findings

### 1. `<ApplicationPass>` encodes two sub-checks

With the current settings:
- **Quality sub-check**: overall grade must meet Overall Pass Grade threshold (1.5). The DM cal card
  failed this — hence `Fail (Quality)`. The QR loaded-image (SymbolQuality=100) passed easily.
- **Format sub-check**: only runs when Data Format Check ≠ None. Currently None → no format check.
  When set to GS1, a GS1 syntax check runs and `Fail (Format)` is possible independently of quality.

The `(Quality)` and `(Format)` suffixes in `<ApplicationPass>` tell which sub-check drove the fail.
VTCCP should parse these as separate reason flags when displaying ApplicationPass.

### 2. Data Format Check = None explains all-undefined DebugGS1

The v1.24 probe showed every GS1 AI as `undefined` in `<DebugGS1>`. This is because the device
is set to **Data Format Check = None** — no GS1 check is being run, so `r.validation.gs1` is an
object stub with no populated data. This is NOT a JavaScript API access-pattern problem.

**Action for v1.25**: To probe `r.validation.gs1` successfully, the user must:
1. Set Data Format Check to **GS1** in Application Settings
2. Scan a GS1-encoded symbol (e.g. the GS1 Format 06 cal card, which has a valid GS1 payload)
3. Run the v1.25 probe script — `r.validation.gs1` will then have real data

This single settings change unblocks the entire GS1 access-pattern investigation.

### 3. `<MinPassGrade>` returns "NA" despite UI showing 1.5

The v1.24 push output showed `<MinPassGrade>NA</MinPassGrade>` and `<MinPassRaw></MinPassRaw>`.
The UI clearly shows Overall Pass Grade = 1.5. Possible explanations:
- The JS property name for the 1.5 value is not `r.metrics.minPassGrade` — wrong accessor
- `minPassGrade` may return the **result** (NA = not applicable at this point?) rather than the threshold
- The threshold may be on a different object (`r.trucheck.minPassGrade`? `r.settings.*`?)

**Action for v1.25**: Add explicit probe of `r.trucheck.minPassGrade`, `r.settings` (or similar),
and enumerate all `r.trucheck` keys to find where 1.5 lives.

### 4. HiBCC is a Data Format Check option

HiBCC (Health Industry Business Communications Council) is a healthcare barcode standard.
When selected, the device checks the scanned symbol's data against HiBCC encoding rules.
VTCCP should log `<ApplicationStandard>HiBCC</ApplicationStandard>` when this is selected
and display appropriately — the `<ApplicationPass>` result carries the same semantics.
No immediate action, but catalog for future symbology/standard expansion.

### 5. Aperture Setting = Auto 50%/80% explains consistent ApertureRef=07

Both live scans (DM: 21.4 mil, QR: 8.7 mil) returned `<ApertureRef>07</ApertureRef>`.
The auto mode targets 50%-80% of the symbol's X dimension and selects the nearest aperture.
- DM cal card: 21.4 mil × 50% ≈ 10.7 mil → aperture 07 (nominal 6.3 mil at cal distance? needs cross-check vs aperture table)
- QR loaded image: 8.7 mil × 50% ≈ 4.4 mil → also resolves to 07
A1 (comms guide) may have the aperture size table. Not blocking any current task.

### 6. Min/Max X Dimension range check

Both live scans fell within [8, 30] mil:
- DM cal card: 21.4 mil ✓
- QR loaded image: 8.7 mil ✓ (barely — exactly at minimum+0.7 mil)

When a symbol's NominalXDim falls outside [8, 30], the application check fails with a
dimensional reason, separate from Quality and Format. VTCCP parser should be aware that
`<ApplicationPass>` can fail for dimension reasons too (not yet observed; flagged for future).

### 7. Advanced Custom Application Standard (collapsed section)

There are more controls below the fold under the "Advanced Custom Application Standard"
expandable section. These likely include: quiet zone requirements, minimum data length,
specific AI requirements, etc. A follow-up screenshot of the expanded state is pending.

### 8. Dot Peen (unchecked)

When checked, presumably engages DPM-mode parameters alongside or in place of ISO 15415/6.
Interacts with the Grading Standard radio — may force ISO 29158 (AIM-DPM) when checked.
Not relevant to current DM475V use case (not DPM scanning), but noted for C1 scope.

---

## Related files

- `TruCheck-Verification-Settings-ReportSettings-2026-05-18.md` — Report Settings panel
- `references/samples/live-scans/v1.24-2026-05-18-Probe-DataMatrix-GS1Format06.xml` — DM push capture showing `ApplicationPass=Fail (Quality)`
- `references/samples/live-scans/v1.24-2026-05-18-QR-LoadedImage-URL.xml` — QR capture showing `ApplicationPass=Pass`
