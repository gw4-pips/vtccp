# DM475V — Settings Backup, Restore & Factory Reset Reference

**Sources**: DM475V Reference Manual 25.4.1.1, DMST Setup Tool Reference Manual 25.4.1.1,
DMCC Comms & Programming Guide 25.4.1.1, DMCC digest (architecture/dmcc-6116sr4-digest.md).

---

## Hierarchy of settings layers

Understanding what each command touches is critical before any reset operation.

| Layer | What it covers | Persistence |
|---|---|---|
| **Connection layer** | `COM.DMCC-RESPONSE`, `DATA.RESULT-TYPE` | Volatile — survives session but resets on power cycle unless `DMCC.SAVE` is called |
| **Device configuration** | All application settings: symbology, aperture, grading standard, application standard, lighting, push script, network client targets, trigger mode, etc. | Volatile until `CONFIG.SAVE` (or DMST floppy icon); then flash |
| **Calibration** | Field calibration values, calibration date | Separately saved — asked at end of calibration wizard ("Save to non-volatile memory?") |

**Critical**: "Write Settings to Verifier" in DMST pushes to volatile memory only. Without
`CONFIG.SAVE` afterward, all settings revert on power cycle. The push script is part of
device configuration — it is erased by `CONFIG.DEFAULT`.

---

## Step 1 — Save a checkpoint BEFORE doing anything else

Do this before any reset or experiment. It takes 30 seconds and gives you a full restore point.

### Method A: DMST GUI (recommended — easiest)

In DMST:
1. Connect to the DM475V and confirm it is in its current (desired-to-preserve) state.
2. Menu: **File → Save Settings to File…** (may also appear as **Manage Settings → Save**)
3. Save as a `.dcf` file (DataMan Configuration File — XML-based). Name it with a timestamp:
   `DM475-63530E-PIPS-Verif-Lab_2026-05-25.dcf`
4. Keep this file in a known location.

To restore from this checkpoint:
- **File → Load Settings from File…** → select your `.dcf` → "Write to Verifier" → CONFIG.SAVE

### Method B: DMCC commands (command-line / VTCCP)

```
DEVICE.BACKUP          -- creates a full configuration snapshot on the device
BACKUP.EXPORT          -- exports that snapshot (returns the backup data)
```

The exported data can be saved externally and later restored with:
```
DEVICE.RESTORE         -- restores from a previously exported backup
```

Note: the exact byte/file mechanism for `BACKUP.EXPORT` is not fully documented in the
available digests. The GUI path (Method A) is more reliable for a one-time checkpoint.

---

## Factory Reset options — three distinct operations

### Option 1: CONFIG.DEFAULT (full device factory reset)

**Scope**: ALL device configuration — symbology settings, aperture, grading standard,
application standard, lighting, trigger mode, network client targets, push script,
FTP settings, output formatting — everything returns to Cognex factory values.

**Calibration**: Field calibration values are erased. Recalibration is required afterward.

**Push script**: Wiped. Must be reinstalled (Write Settings + CONFIG.SAVE after paste).

**Via DMCC:**
```
CONFIG.DEFAULT
CONFIG.SAVE         -- (may not be needed — default may auto-persist; send it anyway)
```

**Via DMST GUI:**
Device menu → **Manage Settings → Reset to Factory Default** (exact label may vary by version;
look for "Reset to Default" or "Factory Default" in the Device or Tools menu).

**⚠ Irreversible without a prior backup.** Always run Step 1 first.

---

### Option 2: DMST "Reset Defaults" button (Application Settings only — partial)

Found at: TruCheck Verification window → Application Settings tab → **Reset Defaults** button (bottom).

**Scope**: Application Settings panel only — aperture selection, grading standard (ISO 15415 vs
AIM-DPM), dot peen setting, application standard. Returns these to Auto defaults.

**NOT affected**: push script, network client, symbology, trigger mode, calibration.

Requires recalibration after use ("After clicking Reset Defaults, you must calibrate the verifier
before using it" — DM475V Manual §Application Settings, p.46).

This is the **least destructive** reset option and the most targeted.

---

### Option 3: DMCC.RESET (connection layer only — minimal)

```
DMCC.RESET
```

**Scope**: Connection-layer settings only — `COM.DMCC-RESPONSE`, `DATA.RESULT-TYPE`.

**NOT affected**: application settings, push script, calibration, anything visible in the DMST UI.

Use this if you suspect that a session's DMCC connection state has been corrupted.
Extremely safe — no application data is touched.

---

## CONFIG.SAVE and CONFIG.RESTORE — the everyday pair

These are not reset operations — they are the normal persistence mechanism.

| Command | What it does |
|---|---|
| `CONFIG.SAVE` | Writes current volatile device configuration to flash. Survives power cycle. Equivalent to DMST floppy/Save Settings icon. |
| `CONFIG.RESTORE` | Reverts device to the last flash-saved state. Useful if you've made changes in volatile memory that went wrong. Power cycle achieves the same effect. |

**The DMST "Write Settings to Verifier" + floppy icon sequence = volatile write + CONFIG.SAVE.**

---

## What is lost on CONFIG.DEFAULT and what must be restored

| Item | Lost? | How to restore |
|---|---|---|
| Push script (all versions) | **Yes** | Paste v1.33 into Format Data → Scripting; Write to Verifier; CONFIG.SAVE |
| Network Client target (10.10.10.19:9004) | **Yes** | Set in Communication → Network Client; Write; CONFIG.SAVE |
| Aperture / Wavelength / Lighting | **Yes** | Set in Verification settings |
| Application Standard (GS1, etc.) | **Yes** | Set in Verification → Application Standard |
| Grading Standard (ISO 15415) | **Yes** | Set in Verification → Grading Standards |
| Trigger mode (Single) | **Yes** | Set in Device Settings → Trigger |
| Field calibration | **Yes** | Full calibration procedure required (NIST card) |
| DEVICE.NAME label | **Yes** | `SET DEVICE.NAME "DM475-63530E-PIPS-Verif-Lab"` |
| COM.DMCC-RESPONSE | Resets to 0 | `SET COM.DMCC-RESPONSE 1`; `DMCC.SAVE` |
| Symbology enables (10 enabled) | **Yes** | See DMCC restore block below |
| Image mirroring (H+V both ON) | **Yes** | `SET CAMERA.MIRROR-HORIZONTAL ON` + `SET CAMERA.MIRROR-VERTICAL ON` |
| Company name | **Yes** | `SET TRUCHECK.COMPANY-NAME Product Identification and Processing Systems, Inc.` |
| Timezone | **Yes** | `SET DEVICE.TIMEZONE America/New_York` |
| NTP | **Yes** | `SET NTP.ENABLE ON` + `SET NTP.SERVER1 time.nist.gov` |

**Total restore time after factory reset with a .dcf backup:** ~5 minutes (load file, write,
save settings). Without a backup: 30–60 minutes to manually reconfigure + recalibration.

---

## PIPS-Verif-Lab complete DMCC restore sequence (post-CONFIG.DEFAULT)

Observed state 2026-05-30, fw 6.1.16_sr4, DM475-63530E-PIPS-Verif-Lab.
Run in Telnet (port 23) after CONFIG.DEFAULT + REBOOT + reconnect.
Send `SET COM.DMCC-RESPONSE 2` first to get ACKs. Finish with CONFIG.SAVE.

```
||>SET COM.DMCC-RESPONSE 2

||>SET TRAIN.AUTO-DISABLE ON

||>SET SYMBOL.DATAMATRIX ON
||>SET SYMBOL.QR ON
||>SET SYMBOL.C128 ON
||>SET SYMBOL.C93 ON
||>SET SYMBOL.C39 ON
||>SET SYMBOL.CODABAR ON
||>SET SYMBOL.I2O5 ON
||>SET SYMBOL.UPC-EAN ON
||>SET SYMBOL.PDF417 ON
||>SET SYMBOL.DATABAR ON

||>SET CAMERA.MIRROR-HORIZONTAL ON
||>SET CAMERA.MIRROR-VERTICAL ON

||>SET DEVICE.TIMEZONE America/New_York
||>SET NTP.ENABLE ON
||>SET NTP.SERVER1 time.nist.gov

||>SET TRUCHECK.COMPANY-NAME Product Identification and Processing Systems, Inc.

||>CONFIG.SAVE
```

Then restore from .dcf backup in DMST (covers push script, network client, aperture,
grading standard, application standard, trigger mode — everything else).

---

## Likely cause of the post-scan image disappearance (live preview not returning)

This is a DMST setting that controls whether the live camera feed resumes after a verification
scan completes. It was previously working and stopped after a DMST restart. Candidates:

1. **LIVEIMG.MODE setting** — `SET LIVEIMG.MODE 2` enables continuous live image polling.
   If this was reverted (e.g., by a volatile-only change that was not CONFIG.SAVEd and then
   the device power-cycled, or by a DMST session that wrote a different value), the live
   preview would stop returning. Check via `GET LIVEIMG.MODE` — expected value: 2.

2. **DMST Application preference** — DMST has a display preference for whether the image
   pane stays in live-preview mode after a scan result, or drops to static. This lives in
   DMST's own local application preferences (not in the device), so it would NOT be fixed
   by CONFIG.RESTORE. Look in DMST: View menu or TruCheck window settings for a
   "Continuous Live View" or "Stay in Live Mode" toggle.

3. **Report output format change side effect** — changing from PDF to HTML and back may
   have triggered a settings write that overrode a related setting. Worth doing a CONFIG.RESTORE
   to the last-saved flash state to see if the behavior returns to normal before considering
   a full factory reset.

**Recommended sequence before considering factory reset:**
1. Save a .dcf checkpoint (Step 1 above) from the current state.
2. Try `CONFIG.RESTORE` (reverts volatile layer to last flash save — may undo the bad change).
3. If not resolved, check `GET LIVEIMG.MODE` value and compare against a working device.
4. Check DMST's own View / display preferences for the live-view toggle.
5. Only if all else fails: factory reset with the .dcf as your restore point.
