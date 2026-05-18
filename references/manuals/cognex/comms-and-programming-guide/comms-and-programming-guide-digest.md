# DMCC Comms & Programming Guide Digest
Source: `comms-and-programming-guide-25.4.1.1.pdf`  
Digest date: 2026-05-18

---

## Key Findings Summary (Open Questions Answered)

| Question | Answer |
|---|---|
| QR Error Correction Level | NOT in `q.general`; candidate path: `r.symbology.quality` (encoded) or manual content bitstream parsing — **probe needed** for `r.trucheck.symbols[0]` keys |
| DataCodewords / ECBudget | Candidate: `r.trucheck.symbols[0].dataCodewords` and `.ecCodewords` — **probe needed to confirm on DM475V 6.1.16_sr4** |
| Per-codeword / per-module data via DMCC | `GET RESULT.XML` with "Extended Verification" enabled includes `<ModulationTable>` and `<CodewordTable>` — OR use `dmccGet()` inside script to pull raw XML block |
| Per-quadrant quiet zones | `q.topQuietZone` and `q.rightQuietZone` only in push `r` object; ULQZ/URQZ/RUQZ/RLQZ only in XML/PDF reports — confirmed structural limit |
| Image-load mechanism | `IMAGE.LOAD [size] [data]` → `IMAGE.REPLAY` confirmed; full sequence documented below |

---

## 1. DMCC Command Namespace

DMCC (DataMan Control Commands) uses an ASCII protocol over TCP port 23.

### Trigger Commands

| Command | Description |
|---|---|
| `TRIGGER ON` | Starts acquisition and decode cycle |
| `TRIGGER OFF` | Stops acquisition |
| `SW.TRIGGER` | Software trigger (equivalent to hardware pulse) |

### Result Read Commands

| Command | Description |
|---|---|
| `GET SYMBOL.RESULT` | Retrieves result string of last decode |
| `GET RESULT.XML` | Retrieves full XML result (including metrics; requires Extended Verification for ModulationTable/CodewordTable blocks) |

### Configuration Read

| Command | Description |
|---|---|
| `GET [Parameter]` | Reads a device parameter (e.g., `GET SYMBOL.DATAMATRIX`) |
| `GET DEVICE.INFO` | Returns serial number, firmware version, hardware type |
| `GET DEVICE.MAC` | Returns Ethernet MAC address |

### Configuration Write

| Command | Description |
|---|---|
| `SET [Parameter] [Val]` | Sets a parameter (e.g., `SET SYMBOL.DATAMATRIX ON`) |
| `CONFIG.SAVE` | Commits changes to non-volatile flash memory |

### Image Commands

| Command | Description |
|---|---|
| `GET IMAGE` | Retrieves last captured image (binary/Base64) |
| `SET LIVEIMG.MODE [0\|2]` | Toggles live image streaming (0=Off, 2=On) |
| `IMAGE.LOAD [size] [data]` | Transfers a raw image to the WBU (Write Buffer Unit) |
| `IMAGE.REPLAY` | Triggers a decode cycle using the buffered image instead of camera |
| `GET WBU.SIZE` | Returns current WBU buffer size |
| `SET WBU.DATA` | Low-level buffer write (used by IMAGE.LOAD) |

### File Transfer

| Command | Description |
|---|---|
| `GET FILE.LOAD "[filename]"` | Downloads a file (e.g., `.dcf` config backup) from device |
| `SET FILE.STORE "[filename]" [data]` | Uploads a file or script to device |
| `SET FTP-IMAGE.IP-ADDRESS [ip]` | Configures FTP target for automatic image/result push |

### System / Status

| Command | Description |
|---|---|
| `BEEP [count] [vol]` | Fires the device internal beeper |
| `REBOOT` | Restarts the device |

---

## 2. Push-Script JavaScript API — `r` Object Reference

| Object Path | Type | Properties / Notes |
|---|---|---|
| `r.symbology` | Object | `.name`, `.id` (AIM ID, e.g. `]d1`), `.quality` (0–100), `.moduleSize` (pixels), `.angle`, `.rows`, `.cols` |
| `r.image` | Object | **Metadata only**: `.index`, `.FoV`, `.RoI`, `.exposureTime`, `.gain`, `.illEnabled` — does NOT contain image bytes |
| `r.trucheck` | Object | All verification data — see sub-properties below |
| `r.metrics` | Object | ISO metrics: `symbolContrast`, `modulation`, `axialNonformity`, `unusedErrorCorrection` (each a `{grade, numericGrade, raw}` object) |
| `r.validation.gs1` | Object | Populated only if Data Format Check = GS1; contains parsed AI values |

**`r.trucheck` (`q`) sub-properties documented in guide:**

| Property | Type | Notes |
|---|---|---|
| `q.jpegImage` | String | Base64-encoded JPEG of verified image — primary source for `<ImageJpeg>` push element |
| `q.calibrationDate` | String | Timestamp of last calibration |
| `q.symbols[0].dataCodewords` | Number | **Candidate for DataCodewords** — needs probe to confirm presence on DM475V 6.1.16_sr4 |
| `q.symbols[0].ecCodewords` | Number | **Candidate for ECBudget** — needs probe to confirm presence on DM475V 6.1.16_sr4 |

**Symbology differences:**

| Aspect | DataMatrix | QR Code |
|---|---|---|
| `r.symbology.rows` / `.cols` | Matrix dimensions | Replaced by `version` |
| Unique metrics | `axialNonuniformity` | `contrastUniformity`, `MRD` (Minimum Reflectance Difference) |
| EC Level | Not applicable | In `r.symbology.quality` (encoded) or `q.symbols[0]` — **unconfirmed on 6.1.16_sr4** |

---

## 3. Image-Load / Image-Replay Sequence

Full confirmed sequence for loaded-image verification:

```
1. SET TRIGGER.TYPE 1        (manual trigger mode)
2. IMAGE.LOAD [size] [data]  (transfer raw image bytes to WBU buffer)
3. IMAGE.REPLAY              (decode cycle against buffered image, not camera)
4. GET SYMBOL.RESULT         (retrieve decoded data + metrics)
   -- OR --
   push script runs automatically after IMAGE.REPLAY
```

The push script fires on IMAGE.REPLAY just as it does on a hardware trigger — `r.trucheck.jpegImage` will contain the loaded image Base64-encoded in the push output.

---

## 4. File Transfer

- `GET FILE.LOAD "[filename]"` — downloads `.dcf` config or script from device
- `SET FILE.STORE "[filename]" [data]` — uploads script or config to device
- FTP: `SET FTP-IMAGE.IP-ADDRESS` / `SET FTP-IMAGE.DIRECTORY` / `SET FTP-RESULT.IP-ADDRESS` — automatic post-scan push to FTP server for image archiving

---

## 5. Configuration Read/Write

- **Reading:** `GET [GROUP].[PARAMETER]` — e.g., `GET DECODER.ROI`
- **Writing:** `SET [GROUP].[PARAMETER] [VALUE]`
- **No batch SET** in one command; use sequential `dmccSet()` calls within `onConnect` or `onResult`
- **Persistence:** `CONFIG.SAVE` (or `SET DEVICE.FLASH SAVE`) to write to flash; without this, config reverts on power cycle

---

## 6. Push vs Poll

| Data | Push (script `r` object) | Poll (`GET RESULT.XML`) |
|---|---|---|
| Decoded string | `r.content` | `<ReadString>` in XML |
| ISO metrics | `r.metrics.*` | `<Quality>` block in XML |
| Per-module reflectance | `q.modulationArray` | `<ModulationTable>` (requires Extended Verification enabled) |
| Per-codeword data | `q.codewordArray` | `<CodewordTable>` (requires Extended Verification enabled) |
| Image | `q.jpegImage` (Base64) | Separate `GET IMAGE` call (may be out-of-sync with XML result) |
| GS1 parsed AIs | `r.validation.gs1.*` | Embedded in XML if GS1 format check enabled |

**Key difference:** `q.jpegImage` is easily captured in the push script (already in `r` object). `GET IMAGE` over DMCC is a separate, potentially unsynchronized call — using push is preferred for VTCCP's image-capture flow.

---

## 7. Open Questions — Results

### a. QR Error Correction Level
- NOT in `q.general` (confirmed — 7 keys only, no ECLevel key)
- Candidate paths (unconfirmed on DM475V 6.1.16_sr4):
  1. `r.symbology.quality` — may encode EC level as part of symbology quality byte
  2. `r.trucheck.symbols[0].errorCorrectionLevel` — needs probe (add to v1.29)
  3. AIM symbology ID (`r.symbology.id`) contains version/EC info for QR

### b. DataCodewords / ECBudget
- Candidate: `r.trucheck.symbols[0].dataCodewords` and `.ecCodewords`
- These are listed in the guide as sub-properties of a `symbols` array on `r.trucheck`
- **Neither `q.symbols` nor `q.trucheck.symbols` appeared in the v1.27 47-key enumeration of `q`** — either the key is `symbols` at the top level of `q`, or this is a v25 API addition not present in firmware 6.1.16_sr4
- **Action:** Add `DebugSymbols` probe to v1.29 — probe `q.symbols`, `q.symbols[0]`, and sub-keys

### c. Per-codeword / per-module data via DMCC
- `GET RESULT.XML` with "Extended Verification" enabled returns `<ModulationTable>` and `<CodewordTable>` blocks
- Alternatively: `dmccGet("RESULT.XML")` inside the push script returns the extended XML which can be substring-searched for these blocks
- Both are available; push-array method (`q.modulationArray`, `q.codewordArray`) confirmed on DM475V is simpler and preferred

### d. Per-quadrant quiet zones
- Confirmed push-channel structural limit: only `q.topQuietZone` and `q.rightQuietZone` as flat `{grade, numericGrade}` objects
- Per-quadrant ULQZ/URQZ/RUQZ/RLQZ grades are only in `GET RESULT.XML` / PDF reports, not in default `r` object
- Could be accessed via `dmccGet("RESULT.XML")` from within the push script if needed — substring parse for `<ULQZ>` etc.

---

## 8. Network Protocol

| Parameter | Value |
|---|---|
| DMCC port | **TCP 23** (Telnet-style) |
| Push / Network Client port | **TCP 9004** (default) |
| Discovery port | UDP 1064 |
| Authentication | Username: `admin`; password via DMST (default: empty) |
| DMCC connection mode | Persistent or per-command |
| Push connection mode | Persistent — device pushes immediately on decode |
| DMCC framing | `> [cmd]\r\n`; responses start with `[` (Status ID) |
| Max message size | ~4 MB (configurable via `RESULT.MAX-SIZE`) |

---

## New Probes Identified for v1.29

| Probe | Target | Purpose |
|---|---|---|
| `DebugSymbols` | `q.symbols`, `q.symbols[0]`, `.dataCodewords`, `.ecCodewords`, `.errorCorrectionLevel` | Confirm whether symbols array exists on DM475V 6.1.16_sr4; primary path to DataCodewords, ECBudget, QR ECLevel |
| `DebugValidationGS1` | `r.validation`, `r.validation.gs1`, `.errorCorrectionLevel` | Alternate path to QR ECLevel via GS1 validation object |
