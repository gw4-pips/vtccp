# DMST Connect Sequence — Splash, Home, Connecting, TruCheck Pre-Scan

**Device**: DM475-63530E-PIPS-Verif-Lab  
**Firmware**: 6.1.16_sr4  
**DMST version**: visible in browser tab as "DataMan Setup Tool 23.2 | v23.x"  
**Date**: 2026-05-25  

---

## Screenshot 1 — DMST Splash Screen

**Screenshot**: `dmst-splash-screen.png`

Appears on DMST launch before the main window is ready. Content:

- "COGNEX" logo (yellow on black)
- "Setup Tool" (large yellow text)
- "DataMan / ID Readers" (lower right)
- Version string visible (partially obscured by VTCCP window in screenshot)
- Copyright: "© Copyright 2026, Cognex Corporation..."
- Trademark notice: "Cognex and DataMan are registered trademarks of Cognex Corporation."
- Patent notice: "Protected by one or more U.S. patents listed at www.cognex.com..."

**For VTCCP**: DMST version is separate from firmware version. DMST communicates the
firmware version it discovers from the device; DMST's own version is irrelevant to
push XML format except insofar as it controls which firmware features DMST can configure.

---

## Screenshot 2 — DMST Home / Connect Screen

**Screenshot**: `dmst-home-connect-screen.png`

This is the main DMST window before connecting to any reader. Left nav:
- **Connect** (selected, blue)
- Maintenance
- Repair & Support
- Backup / Restore / Update Firmware
- Reader Groups / Image Playback / Options / About / Exit

Top bar controls: Refresh, Grouping (Interface Type), Filter, View Hidden (0)

### Device list

| Name | Type | Address | Firmware | Status | Interface | MAC | Serial |
|---|---|---|---|---|---|---|---|
| *(Serial group)* | | | | | | | |
| COM1 | RS232 | COM1 | — | Unknown | Serial | — | — |
| *(Network group)* | | | | | | | |
| DM475-63530E-PIPS-Verif-Lab | **DM470** | 10.10.10.7 | 6.1.16_sr4 | Discovered | Network | 00-D0-24-63-53-0E | 1A1903PP010754 |
| RTM Lean | RTM Controller | 127.0.0.1 | — | Discovered | Network | 02-00-00-00-00-00 | — |

Bottom buttons: **Compare Configurations** | **Process Monitor** | **Connect**

### Key findings from this screen

**DM470 vs DM475V — device type discrepancy:**
DMST's device list shows type "DM470" for what is physically a DM475V. This is the
model-family identifier used by DMST for device management, NOT the precise model string.
The DMCC command `GET DEVICE.TYPE` may return "DM470" (family) rather than "DM475V"
(exact model). **VTCCP currently stores DEVICE.TYPE in `VerificationRecord.DeviceModel`
and the Excel "Reader Model" column.** We need to verify what `GET DEVICE.TYPE` actually
returns versus what DMST shows — they may differ. If DEVICE.TYPE returns "DM470", the
Reader Model column will show "DM470", not "DM475V". The physical DM475V branding is on
the hardware label, not necessarily in the firmware type string.

**RTM Lean at 127.0.0.1:**
"RTM Lean" (RTM = Runtime Manager or Reader Test Manager) is a Cognex local service
running on the DMST host machine. It appears in the DMST device list as a loopback
network device. Not relevant to VTCCP — VTCCP connects directly to the DM475V at
10.10.10.7, not to any local Cognex service.

**COM1 at RS232 — Unknown:**
A serial port device (possibly a legacy reader or test fixture). Status "Unknown" means
it was not responding to discovery probes at the time of the screenshot. The DM475V
is on network/GigE; serial is unused for this unit.

**MAC address confirmed**: 00-D0-24-63-53-0E — Cognex OUI (00-D0-24). ✓
**Serial confirmed**: 1A1903PP010754 — matches push XML `<UnitSerial>` ✓

---

## Screenshot 3 — "Connecting..." Progress Dialog

**Screenshots**: `dmst-connecting-dialog.png` (in-progress, ~90% bar) and
`dmst-connecting-dialog-complete.png` (complete, 100% bar + "Connected" final line)

Modal dialog appears after double-clicking the device row (or clicking Connect).
Title bar: "DataMan Setup Tool – DM475-63530E-PIPS-Verif-Lab [10.10.10.7]"

Status log messages — **complete four-step sequence** (confirmed from second screenshot):
1. "Establishing connection to device..."
2. "Retrieving parameters..."
3. "Retrieving configuration..."
4. **"Connected"**

Progress bar is 100% green when "Connected" appears. The dialog then auto-dismisses
and the DMST device workspace tab opens.

Button: **Retry** (available if connection fails)

### DMCC init sequence — what DMST is doing

These three log lines map directly to the connection protocol steps. Based on the
Cognex DMCC/SDK documentation (A1 digest) and what the SDK does internally:

#### Step 1 — "Establishing connection to device"
- TCP socket opened to 10.10.10.7:44444 (GigE/Ethernet default port)
- Protocol handshake: DMST sends initial DMCC session open request
- Device acknowledges and assigns a session ID
- SDK equivalents: `DataManSystem.Connect()` → internal socket establishment

#### Step 2 — "Retrieving parameters"
DMST bulk-queries static device identity and capability parameters. Likely DMCC GETs:
```
GET DEVICE.TYPE              → "DM470" (model family)
GET DEVICE.NAME              → "DM475-63530E-PIPS-Verif-Lab"
GET DEVICE.FIRMWARE.VERSION  → "6.1.16_sr4"
GET DEVICE.SERIAL            → "1A1903PP010754"
GET DEVICE.MAC               → "00-D0-24-63-53-0E"
GET DEVICE.INFO              → (composite or additional identity fields)
```
These populate the device list columns and the DMST title bar.

#### Step 3 — "Retrieving configuration"
DMST bulk-queries the current running configuration of the device. This is the large
parameter read — every setting the DMST UI will display in its tabs. Likely includes:

**Triggering:**
```
GET TRIGGER.TYPE             → (Continuous / Single / External)
GET TRIGGER.ENABLED          → 1 or 0
```

**Image & acquisition:**
```
GET IMAGE.SIZE               → 0 (Full) / 1 (1/4) / 2 (1/16) / 3 (1/64)
GET LIVEIMG.MODE             → 2 (send image with result — expected for TruCheck)
GET EXPOSURE                 → exposure time
GET GAIN                     → sensor gain
```

**TruCheck / Verification application:**
```
GET TRUCHECK.STANDARD        → "ISO 15415:2011" (user-configured)
GET TRUCHECK.APERTURE        → 16
GET TRUCHECK.WAVELENGTH      → 660
GET TRUCHECK.LIGHTING        → "45Q"
GET TRUCHECK.OPERATOR        → "GW4"
GET TRUCHECK.COMPANY         → "Product Identification..."
... (all TC configuration keys)
```

**Push / result delivery:**
```
GET RESULT.READER.1.STATISTICS → (result delivery config)
GET DMCC.PUSH.SCRIPT         → the current JS push script (v1.33 content)
```

**UPC/EAN supplemental and other reader settings:**
```
GET UPC-EAN.SUPPLEMENT       → 0–4
... (symbology enables, GS1 parser, etc.)
```

This is why the progress bar takes a moment — it's a bulk read of all configuration
parameters, not just a handful of identity keys. The Cognex SDK's `GetParamListAsync()`
or equivalent likely issues these as a pipelined batch.

---

### What DMST does next — launching the TruCheck Verification window

After "Retrieving configuration..." completes, DMST opens the main device workspace
(the tabbed editor) AND the TruCheck Verification window simultaneously. The key
question is whether any specific DMCC command triggers the TC window to open.

**Answer: no dedicated "launch TruCheck" DMCC command exists.**

The TruCheck Verification window is a DMST UI construct, not a device mode. DMST:
1. Reads the device's TruCheck configuration settings (aperture, wavelength, etc.)
2. Opens the TC verification window in the DMST UI, populated from those settings
3. Subscribes to the device's result stream (push XML events)
4. Displays each incoming result in the TC window

The device itself is always in "TruCheck mode" when its application configuration
is set up for TruCheck — that configuration lives on the device, not in DMST. DMST
is purely a display and configuration frontend; it does not put the device into a
special state just by opening the TC window.

**Implication for VTCCP**: VTCCP does not need to issue any special "start TruCheck"
command on connect. It just:
1. Connects to the device (TCP to port 44444)
2. Issues the identity/parameter GETs (Step 2 above — minimal subset needed)
3. Registers the push XML result event handler
4. Receives results as the device scans

The device delivers push XML automatically on every decode. VTCCP's `DeviceSession.cs`
already implements this via the Cognex SDK `ResultsReceived` event subscription.

---

### LIVEIMG.MODE and the Main tab image bug

> **⚠ CORRECTION 2026-06-01** — The analysis below was based on a false theory.
> Device-confirmed known-good state (DM-KnownGood-Snapshot_2026-05-31): `LIVEIMG.MODE = 0`
> with the TC panel image fully working. `LIVEIMG.MODE = 0` is CORRECT. Setting it to 2
> is NOT the fix. The actual root cause was NVRAM corruption from a prior `COM.DMCC-SAVE`
> call; fixed by `CONFIG.DEFAULT + CONFIG.SAVE + REBOOT`.
> See `WORKING-NOTES.md §RESOLVED: Image missing` for the confirmed fix.

~~From the "Retrieving configuration" step, DMST would read `GET LIVEIMG.MODE`. The
expected value is **2** (send image data with each result). If `LIVEIMG.MODE=0` or 1,
the image stops appearing in the DMST TC Main tab.~~

**Diagnostic command**: `GET LIVEIMG.MODE` — confirmed known-good response: `0` (not 2)
**If blank image appears**: use `CONFIG.DEFAULT + CONFIG.SAVE + REBOOT` (not SET LIVEIMG.MODE)

---

## Screenshot 4 — TruCheck Window Pre-Scan (Main Tab)

**Screenshot**: `dmst-trucheck-pre-scan-main.png`

Title: "TruCheck Verification – DM475-63530E-PIPS-Verif-Lab"  
Tabs: Main | General Characteristic | Data Detail | Quality Detail | Advanced Detail | Histogram | Report

### Pre-scan Main tab state

**Image pane (left)**:
- Yellow "COGNEX" banner bar at the top of the image area
- Dark gray camera feed below with a small red **+** crosshair at center
- This is the **live view** — the camera is actively capturing and streaming to DMST
- The yellow Cognex banner is overlaid by the device's image annotation layer

**Grade/data fields (right)**:
- Overall Grade: **empty**
- Format Grade: **empty**
- Symbology: **empty**
- Data: **empty**
- Acceptance Criteria: **empty**
- **"Go Live" button**: visible (bottom right) — pressing this triggers the live
  verification loop (device triggers on each barcode present in FOV)

Status bar: "Last calibrated on: 5/20/2026 1:14:58 AM"

### Pre-scan vs post-scan image comparison — the bug context

| State | Image pane appearance |
|---|---|
| **Pre-scan (this screenshot)** | Yellow Cognex banner + dark gray live camera feed + red crosshair |
| **Post-scan NORMAL** | Captured verification image of the scanned barcode symbol |
| **Post-scan BUG (observed this session)** | Gray pane + red crosshair — same as pre-scan, no captured image |

The post-scan bug reproduces the pre-scan visual exactly — the image pane doesn't update
with the captured symbol image after a successful scan. This strongly suggests the device
is not delivering `JpegImageBase64` in the result, which in turn points to
`LIVEIMG.MODE ≠ 2`. When `LIVEIMG.MODE=0`, results are delivered without the image JPEG.

**The yellow Cognex banner**: this is the device's DMST connection indicator — it appears
in the image frame whenever DMST is connected and the device is in the "DataMan Setup Tool"
application context. It is NOT part of the verification image; it disappears in the
captured verification image shown in the Report tab (the captured symbol image in the
Report tab has no yellow banner). This banner is injected by the device's image annotation
layer specifically for the DMST live preview.

### "Go Live" button

The "Go Live" button activates continuous TruCheck verification mode — the device
continuously triggers and verifies any barcode placed in the field of view without
requiring a manual trigger per scan. In this mode the result counter (1/1, 1/2, etc.)
advances automatically. The **<<** and **>>** navigation arrows traverse the result
history; the counter shows "current/total".

---

## Summary — VTCCP DeviceSession init implications

Based on these four screenshots, the recommended minimal init sequence for
`DeviceSession.ConnectAsync()` (additions beyond current implementation):

```
1. SDK Connect() → TCP to host:44444
2. GET DEVICE.TYPE          → store in DeviceModel
3. GET DEVICE.NAME          → validate matches config name
4. GET DEVICE.FIRMWARE.VERSION → store, log
5. GET DEVICE.SERIAL        → store as UnitSerial for audit
6. GET LIVEIMG.MODE         → validate = 2; warn if not 2
7. GET TRUCHECK.STANDARD    → store grading standard edition string
   (echo this in report — do NOT hard-code "ISO 15415:2011")
8. Register ResultsReceived handler for push XML
```

Step 6 (LIVEIMG.MODE check) is the diagnostic for the image bug — add it to
`ConnectAsync` so VTCCP immediately surfaces the misconfiguration rather than
silently producing results with no images.

Step 7 (TRUCHECK.STANDARD) resolves the ISO edition question: read it from the
device rather than parsing from push XML, guaranteeing we echo the operator's
configured edition, whatever it is.

**DeviceSession wiring still pending** (from prior session plan): add `_scraper` field,
`BuildReportPath(deviceInfo.Name)` at ConnectAsync, `TryMergeAsync` in result handler,
`Stop()` at disconnect. Steps 6 and 7 above should be added at the same time.

---

## Packet capture — intercepting the full DMCC exchange

### Short answer: yes, Wireshark

The DMCC protocol is **unencrypted plaintext XML over TCP port 44444**. There is no
TLS, no obfuscation. Every byte DMST sends to the device, and every byte the device
sends back (including push XML result payloads), is fully readable in a packet capture.

### Setup on the DMST Windows host

1. **Install Npcap** (https://npcap.com) — the Windows packet capture driver.
   Wireshark installs this automatically if not already present.
2. **Install Wireshark** (https://wireshark.org) — run as Administrator for capture.
3. **Select the right interface** — the NIC on the 10.10.10.x network (the subnet
   where the DM475V lives at 10.10.10.7).
4. **Start capture** before connecting in DMST.
5. **Apply this display filter** to isolate only the DMST↔device traffic:
   ```
   ip.addr == 10.10.10.7 and tcp.port == 44444
   ```
6. **Connect to device in DMST** — the full four-step sequence flows through.
7. **Do one live scan** — the result push XML will appear as a large TCP segment.
8. **Stop capture**, then use **Analyze → Follow → TCP Stream** to read the full
   application-layer conversation as reassembled XML text.

### What you will see

The TCP stream will show interleaved client (DMST) and server (device) messages in
plain XML. Approximate structure:

```xml
<!-- DMST → device: initial greeting / protocol negotiation -->
<DMCCRequest><Command>...</Command></DMCCRequest>

<!-- device → DMST: acknowledgement -->
<DMCCResponse>...</DMCCResponse>

<!-- DMST → device: "Retrieving parameters" phase -->
<DMCCRequest><Command>GET DEVICE.TYPE</Command></DMCCRequest>
<DMCCResponse><Value>DM470</Value></DMCCResponse>

<DMCCRequest><Command>GET DEVICE.NAME</Command></DMCCRequest>
<DMCCResponse><Value>DM475-63530E-PIPS-Verif-Lab</Value></DMCCResponse>

<!-- ... more GETs ... -->

<!-- DMST → device: "Retrieving configuration" phase -->
<DMCCRequest><Command>GET LIVEIMG.MODE</Command></DMCCRequest>
<DMCCResponse><Value>2</Value></DMCCResponse>

<!-- ... all TruCheck config GETs ... -->

<!-- device → DMST: push XML result on each scan -->
<DMCCResponse>
  <PushScriptDiag>v1.33 q=r.trucheck m=found</PushScriptDiag>
  <OverallGrade>D</OverallGrade>
  ...
</DMCCResponse>
```

### Why this is valuable for VTCCP

A single Wireshark capture session gives us the **ground truth** on every question
about the DMST initialization sequence:

| Question | Answered by capture |
|---|---|
| Exact DMCC commands in "Retrieving parameters" | ✓ — every GET visible |
| Exact DMCC commands in "Retrieving configuration" | ✓ — full list |
| Whether any special TruCheck-launch command exists | ✓ — definitively |
| Order in which DMST issues commands | ✓ |
| What DMST sends when "Go Live" is pressed | ✓ |
| Full push XML payload as delivered to DMST | ✓ — same payload VTCCP receives |
| GET TRUCHECK.STANDARD response format | ✓ |
| GET LIVEIMG.MODE response | ✓ |
| GET DEVICE.TYPE response ("DM470" vs "DM475V") | ✓ |

### VTCCP-side capture

VTCCP's SDK connection also goes to port 44444. If DMST and VTCCP are connected to
the same device at the same time, both sets of traffic appear in the capture (the
device supports multiple simultaneous connections). Filter by the TCP connection
source port to distinguish which is which.

Alternatively: disconnect DMST, run VTCCP alone, and capture VTCCP's connect
sequence — this verifies that VTCCP's `DeviceSession.ConnectAsync()` issues the
right GETs in the right order.

### Alternative: Wireshark on a network tap

If the DMST host cannot run Wireshark (managed IT environment), a **network tap** or
a **managed switch with port mirroring** can mirror the 10.10.10.x port to a separate
capture machine running Wireshark. The DM475V and the DMST host are both on the same
LAN segment, so this works without any configuration on either endpoint.

### Capture checklist

- [ ] Npcap + Wireshark installed on DMST host
- [ ] Capture started **before** connecting in DMST
- [ ] Display filter: `ip.addr == 10.10.10.7 and tcp.port == 44444`
- [ ] Connect to device, observe the four-step sequence
- [ ] Press "Go Live", present a barcode, get one scan result
- [ ] Stop capture
- [ ] File → Export Specified Packets → save as `.pcapng`
- [ ] Analyze → Follow → TCP Stream → copy full exchange text
- [ ] Share exchange text for VTCCP init-sequence verification
