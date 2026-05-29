# DMCC Reference — HTML Pages Index

Source: Cognex DataMan Control Commands Reference, fw 6.1.16_sr4
Format: MadCap Flare WebHelp2 HTML export
Filed: 2026-05-29 — COMPLETE SET (794 files via 7z archive)

---

## Coverage summary

| Category | Overview | Individual command pages |
|---|---|---|
| Action | `Action.overview.htm` | 95 |
| Camera | `Camera.overview.htm` | 101 |
| Code Quality | `Code Quality.overview.htm` | 9 |
| Communication | `Communication.overview.htm` | 211 |
| Data Formatting | `Data Formatting.overview.htm` | 21 |
| Data Validation | `Data Validation.overview.htm` | 12 |
| Decoder | `Decoder.overview.htm` | 111 |
| I/O | `IO.overview.htm` | 51 |
| Record/Playback | `RecordPlayback.overview.htm` | 28 |
| Symbology | `Symbology.overview.htm` | 122 |
| System | `System.overview.htm` | 20 |
| **Total** | **11 overviews** | **781 detail pages** |

Also present: `dmcc-main.html` (help index), `data-formatting-tokens.htm` (token reference), `DataMan_Control_Commands_Overview.htm` (protocol overview).

**Total files: 795** (781 idp*.htm detail pages + 14 named files + this INDEX.md)

---

## Priority files for trigger investigation

| File | Command | Category | Notes |
|---|---|---|---|
| `idp10154189968.htm` | `TRIGGER.TYPE` | Camera | GET/SET — trigger mode enum. **Key file for trigger reset plan.** |
| `idp10153430112.htm` | `TRIGGER` (action) | Action | Software trigger ON/OFF |
| `idp10153440704.htm` | `TRIGGER` (state) | Action | Get current trigger state |
| `idp10153456512.htm` | `MOTION-DETECTION.ACTIVE` | Action | Check motion detection state |
| `idp10153461680.htm` | `MOTION-DETECTION.ENABLE` | Action | Enable/disable motion detection |

---

## Other files of interest

| File | Command | Category | Notes |
|---|---|---|---|
| `IO.overview.htm` | I/O Commands | I/O | 51 commands — input/output lines, beeper, etc. |
| `System.overview.htm` | System Commands | System | 20 commands — reboot, reset, firmware, etc. |
| `Symbology.overview.htm` | Symbology Commands | Symbology | 122 commands — UPC-EAN, Code 128, DataMatrix, QR settings |
| `RecordPlayback.overview.htm` | Record/Playback Commands | Record/Playback | 28 commands — IMAGE.LOAD, IMAGE.REPLAY |

### IMAGE.LOAD / IMAGE.REPLAY (D4 scope)
These live in the RecordPlayback category. Look for `idp*.htm` files linked from `RecordPlayback.overview.htm` for the exact DMCC key names and argument syntax.
