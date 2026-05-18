# Cognex DataMan Setup Tool (DMST) Reference Manual — Digest

**Source**: `setup-tool-reference-manual-25.4.1.1.pdf` (6.5 MB)
**Version**: 2025.4.1.1
**Digested by**: explorer subagent, 2026-05-18

**Provenance note**: Subagent merged DMST + Comms guide content because of
overlap; primary focus here is DMST. For commands also documented in
the Comms guide, the sibling
`references/manuals/cognex/comms-and-programming-guide/comms-and-programming-guide-25.4.1.1-digest.md`
is more authoritative.

---

## TOP-LINE FINDINGS (read this section first)

1. **DMST is a thin client.** It does NOT run the decode logic or hold
   the JavaScript runtime — those live in the device firmware. DMST sends
   DMCC commands over port 23 to configure the device. The push script is
   compiled/loaded into the **device's** volatile memory and runs on
   **device's** CPU. Worth understanding when reasoning about timing or
   "DMST is connected so triggers don't work" issues.

2. **JavaScript runtime is ES5-compliant** (Standard ECMA-262 5th
   Edition). Our empirical "no `const`/`let`/arrow functions" rule from
   v1.10 onwards is now formally confirmed.

3. **`r.trucheck.jpegImage` is the answer** to how we get the captured
   image — Base64-encoded JPEG string of the verified image, accessible
   from the script. (Independently confirmed by the Comms guide digest.)

4. **Lifecycle hooks beyond `onResult`** documented:
   - `onResult(decodeResults, readerProperties, output)` — primary
   - `onGenerateFTPFilename(decodeResults, readerProperties, output)` — for dynamic FTP file naming
   - **Global variables persist** across invocations (device power
     cycle resets, script reload resets — otherwise they survive).

5. **Output channel targeting** — instead of `output.content = ...`,
   can use `output.NetworkClient = "<xml>...</xml>"` to target a
   specific output channel. Useful if multiple Network Clients are
   configured.

6. **DMCC-from-script** — scripts can issue DMCC commands at runtime
   via `dmccGet(cmd)`, `dmccSet(cmd, args)`, `dmccCommand(cmd, arg)`,
   `dmccSend(string)`. **This opens a programmability door** we
   haven't been using — push script could potentially query device
   state mid-decode to enrich the output.

7. **Built-in `encode_base64(string)`** — for emitting binary as XML
   text content. **Already-available primitive for image-in-XML push.**

8. **Logging via `print()` or `console.log()`** outputs to DMST's
   Scripting tab console — debug aid we haven't been using during
   probe iterations. Would have shortened many probe cycles.

9. **DMST 2025.4.1.1 ↔ firmware compatibility**: "supports firmware
   6.1.x through 6.5.x" per subagent extraction. Our DM475V on
   6.1.16_sr4 is in-range. **The 2026.1 DMST you're installing right
   now is a NEWER tool version — its compatibility with 6.1.x firmware
   is NOT covered by this manual.** Watch for compatibility warnings
   during your install.

10. **Image-load via scripting is NOT supported** — "Scripting cannot
    'load' a file directly via JS (no `fs` module), but the reader can
    be triggered via DMCC `WBU` (Write Buffer) to ingest an external
    image for processing." Confirms that VTCCP's loaded-image flow goes
    through DMCC, not through the push script.

---

## 1. Overall Architecture

| Layer | Where it runs | What it does |
|---|---|---|
| **DMST** | PC (Windows) | Configuration UI, scripting editor, live view, calibration wizard. **Stateless re: decode** — purely a remote control. |
| **Device firmware** | DM475V CPU | RTOS, decoder engine, JavaScript VM, FTP/TCP push agents |
| **Push scripts** | Device CPU (in VM) | Run during decode cycle to format output |

**Connectivity**:
- DMCC over TCP port 23 (primary)
- DMCC over Serial / USB-CDC (alternative)
- Limited concurrent connections; DMST in "Live" mode polls
  aggressively and **may starve external clients** — matches our
  empirical "close DMST before pressing Trigger" rule.

---

## 2. Scripting Environment (PRIMARY FOCUS)

### Runtime
- **ECMAScript**: ES5 compliant (ECMA-262 5th Edition). No `const`,
  `let`, arrow functions, template literals, classes.
- **Execution model**: scripts thread-safe within own context.
- **Script load order**: 1. FTP Naming, 2. Communication, 3. Data Formatting.
- **Persistence**: globals defined outside hook functions persist across
  invocations until power cycle or script reload.
- **Memory/perf**: no specific limits documented in the digest —
  worth checking source PDF.

### Globals available
- `dmccGet(cmd)` — read a setting
- `dmccSet(cmd, args...)` — write a setting
- `dmccCommand(cmd, arg)` — generic
- `dmccSend(string)` — evaluate a full DMCC string (simpler form)
- `decode_sequences(string)` — convert escape sequences (e.g.
  `\\Ctrl-B;`) for HID output
- **`encode_base64(string)`** — string → base64 (handy for image-in-XML)
- `print(x)` / `console.log(x)` — debug output to DMST Scripting console

### Lifecycle hooks

| Hook | When fired | Args | Used by us? |
|---|---|---|---|
| `onResult` | On every decode result (or No Read) | `(decodeResults, readerProperties, output)` | Yes — our primary entrypoint |
| `onGenerateFTPFilename` | When FTP push is about to fire | `(decodeResults, readerProperties, output)` | No — we don't use FTP |

The function our scripts have been calling `formatResult(r)` is actually
called `onResult` per the documented API; the firmware seems to accept
either name through some alias mechanism. **Worth verifying empirically
in v1.24** — if the device strictly enforces `onResult`, our scripts
work for unknown reasons (perhaps a `formatResult` alias was added to
the firmware) and we should rename to the canonical name for safety.

### Error handling
- Uncaught exceptions are logged to DMST Scripting console.
- Reader does NOT crash on script error.
- Current trigger's output is terminated.

---

## 3. The `r` Object — Formal API (cross-reference)

Identical to the Comms guide reference — see the comms digest sibling
file for the full table. Key DMST-specific additions / clarifications:

- `r.readSetup` — integer 1-16, indicates which Read Setup was used.
  **Not in our v1.23 probe — should add to v1.24.** Lets us
  distinguish results from different verification configurations.
- `r.annotation` — added in DMST 2025.4 for multi-reader sync
  (primary/secondary clusters). **Not relevant for our single-device
  setup.**

### Empirical-vs-documented gaps (consolidated with Comms digest)

| Field | Source of evidence |
|---|---|
| `r.image.FoV / RoI / exposureTime / gain` | Documented, not probed |
| `r.ledIntensity` | Probed, not documented (likely 6.1.x firmware extension) |
| `axialNonUniformity` (capital U) | Documented spelling |
| `axialNonuniformity` (lowercase u) | Our probe spelling |
| `r.readSetup` | Documented, not probed |

---

## 4. `r.image` and Image Handling

### Image metadata
`r.image` exposes:
- `id` — unique image ID
- `index` — image index
- `exposureTime`, `gain` — capture settings
- `RoI`, `FoV` — rectangles `{x, y, w, h}`

### Actual image bytes
**`r.trucheck.jpegImage` — Base64-encoded JPEG string.** This is the
property our v1.24 push script can emit directly. Already a string, no
conversion needed inside the script.

### Loading images
- **Via DMST UI**: "Image Playback" feature loads images from disk.
- **Via scripting**: NOT possible (no `fs` module).
- **Via DMCC**: `WBU` (Write Buffer) command — push external image
  bytes to the reader, which then runs them through the decode pipeline.
  This is the path VTCCP would use for the loaded-image flow.

---

## 5. Output Formatting Modes

DMST supports several output modes per device setup:
- **Basic Formatting** — built-in templates, no script
- **Script-Based Formatting** — our path; `onResult` hook owns the
  output

### Targeting specific channels
In `onResult`:
```javascript
output.content = "Default string for all channels";
output.NetworkClient = "<xml>...</xml>";  // Channel-specific
```
Our v1.23 only sets `output.content`. This works for the Network
Client receiving the default, but if e.g. a serial output is also
enabled, it'll get the same XML. Per-channel targeting is an option if
that becomes a problem.

---

## 6. Push (Network Client) Configuration

Configured at **Communication → Network Client** in DMST UI.

Trigger conditions:
- Always
- On Pass
- On Fail

Multiple Network Clients can be configured per setup.

The IP/port pair binds the channel to a destination; in our deployment
it's `10.10.10.19:9004` (host PC + `DmstListenPort`).

Backup behavior on network failure: not detailed in the digest — worth
checking source PDF if reliability becomes a concern.

---

## 7. Configuration UI Tour

Relevant DMST panels for VTCCP setup:

| Panel | Path | Used for |
|---|---|---|
| Format Data → Basic | Application Steps sidebar | Built-in formatting (not our path) |
| Format Data → Scripting | Application Steps sidebar | Where we paste `DmstPushScript_v1.js` |
| Communication → Network Client | Application Steps sidebar | Push target IP + port |
| Verification | Application Steps sidebar | Symbology, application standard, aperture, lighting |
| Trigger | Device Settings | Trigger Type (must be Single for software TRIGGER) |
| Calibration | Setup | Card-based calibration workflow |

---

## 8. Calibration Workflow

Step-by-step as documented:
1. Place Calibration Card (e.g., Data Matrix on DM475V).
2. Use **Setup → Calibration** tab.
3. Enter the **R** (Reflectance) values printed on the card.
4. Reader adjusts exposure and internal light gain to match the
   NIST-traceable standard.
5. Calibration result stored in non-volatile memory.

Calibration is **required for "Grade" reporting to be valid** — without
it, the verifier operates in an uncalibrated mode (effectively SBG-style)
and Grade values are not 15426-2-conformant.

This is the formal basis for our OpticsCompliant computation:
```
OpticsCompliant =
  readerProperties.status3D.fieldCalibrated
  AND OpticsSource == LiveScan
  AND ApertureRef and Wavelength and Lighting all reported
```

---

## 9. Multi-Symbology / Multi-Trigger / Multi-Step

Subagent didn't go deep here. The 8072V manual is more authoritative for
verification workflows; this DMST manual primarily covers the UI side.
Worth a focused source-PDF dive if we ever need to verify >1 symbology
per trigger.

---

## 10. Firmware Compatibility

- **DMST 2025.4.1.1 supports firmware 6.1.x through 6.5.x**
- **DM475V on 6.1.16_sr4** is in-range ✓
- This DMST version is **explicitly compatible** with our deployed
  firmware

**For your 2026.1 DMST install in progress**: this manual does NOT
cover 2026.1's compatibility matrix. Likely 2026.1 supports firmware
6.5.x and a 6.6.x or 7.0.x — but won't necessarily support 6.1.x.
**If 2026.1 refuses to talk to the DM475V on 6.1.16_sr4, you may need
to either**:
- (a) keep 2025.4.1.1 installed alongside 2026.1, or
- (b) upgrade the device firmware (separate ritual, requires testing
  the whole VTCCP push pipeline against the new firmware)

Watch for compatibility warnings on first connection attempt.

---

## 11. What's New in 2025.4

(Per subagent — verify against source PDF for full release notes)

- Enhanced **TruCheckResult** object: more granular alignment + pattern
  grades.
- Improved **DMCC over Scripting**: `dmccSend` handles multi-line
  responses more gracefully.
- **Multi-Reader Sync**: `r.annotation` added for primary/secondary
  clusters.

This explains why our March 2026 v1.10/v1.11 probes saw a different
`r.trucheck` shape than the v1.23 (May) probes — `TruCheckResult` was
enhanced in this DMST release.

---

## 12. Anything Image-Load Related

- **DMST UI: Image Playback** — load disk images, run through pipeline,
  see grades + results in DMST's UI. **Not** wired through Network
  Client push (the script doesn't see playback-mode events unless
  configured to).
- **Scripting**: cannot load files (no `fs`).
- **DMCC**: `WBU` to push bytes to reader, then trigger.

---

## 13. DMCC Commands Originating from DMST

When DMST connects to a device, it issues a steady stream of DMCC
queries to populate its UI: device info, current settings, live image
poll, statistics. This stream **competes for port 23** with our
`DmccClient` — matches the README operational note that DMST must be
disconnected before the ⚡ Trigger Scan button works.

In future, if we wanted to make VTCCP coexist better with DMST, the
options are:
- Use the DataMan SDK exclusively (it does its own concurrency
  management — already what `DataManSdkClient.cs` does for the push
  session)
- Detect DMST's presence and warn the user before attempting trigger
- Negotiate `COM.DMCC-TARGET` to share the port more gracefully (if
  firmware supports it)

---

## 14. What this changes for VTCCP

### Already-known-but-now-formally-confirmed
- ES5 scripting requirement
- DMST and DmccClient port 23 contention
- Calibration required for valid Grade reporting

### New unlocks
- **`r.trucheck.jpegImage` base64 emission** — biggest single win.
- **`r.trucheck.calibrationDate` + `readerProperties.status3D.fieldCalibrated`**
  → first-class OpticsCompliant inputs (already on v1.24 wire list).
- **`r.readSetup`** → which configuration produced this scan.
  Useful for multi-mode deployments.
- **`encode_base64()`** built-in primitive — no need to roll our own
  base64 if we ever need to emit binary inside text.
- **`print()` / `console.log()`** for in-script debugging during
  v1.24 probe iterations — would have caught the
  `axialNonUniformity` vs `axialNonuniformity` mismatch immediately.

### Architecture insurance
- DMST is a remote control, not a runtime. Reasoning about timing,
  state, or capability should look at the firmware version, not the
  DMST version.
- Push scripts run in the device VM with persistent globals — opens
  the door to per-session state like "last image ID" or "sequence
  counter" that survives across scans without external storage.

### Operational note for the 2026.1 install
- 2026.1 compatibility with 6.1.16_sr4 firmware is **not covered** by
  this manual. Watch for warnings; consider keeping 2025.4.1.1
  installed alongside as a fallback. If 2026.1 refuses to connect, the
  firmware-upgrade ritual is a separate decision tree (and requires
  re-validating the entire VTCCP push pipeline against new firmware).
