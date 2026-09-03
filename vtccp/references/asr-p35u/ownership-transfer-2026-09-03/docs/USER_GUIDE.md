# RFID FlexWedge Pro — User Guide

**Version:** 1.0  
**Validated:** 03 September 2026  
**Supported reader:** AsReader P35U Desktop UHF RFID Reader  
**Platform:** Windows 10 or Windows 11

> This guide describes the validated RFID FlexWedge Pro reference application.
> VeriWedge RFID is a derivative intended for a different market and may have
> different controls, defaults, and integrations in VTCCP.

## 1. What FlexWedge does

RFID FlexWedge Pro reads UHF RFID tags and sends the selected result to the
Windows application that currently has keyboard focus. It can work with
Command Pilot, Excel, Notepad, web forms, and other applications that accept
keyboard input.

For supported EPC schemes, FlexWedge also decodes useful identity fields. For
an SGTIN-96 tag these include:

- raw EPC
- GTIN-14
- serial number
- EPC URI
- RSSI signal level
- antenna number
- optional TID
- optional lock status

## 2. Before starting

You need:

- the complete FlexWedge application folder
- `AsReaderP3xU.dll` in the same folder as the application
- an AsReader P35U connected by USB
- one or more compatible UHF RFID tags

Only one program can control the reader at a time. Close the AsReader sample
application, another FlexWedge window, or any other program using the same COM
port.

## 3. Main window

### Reader controls

- **Port** — the Windows COM port assigned to the P35U.
- **Refresh** — refreshes the available COM-port list.
- **Connect / Disconnect** — opens or closes the reader connection.
- **Preset** — shows the active operating preset.
- **Power (dBm)** — selected transmit power, from 13 through 27 dBm.
- **Set** — applies the displayed power to the connected reader.

### Reading controls

- **Start Reading** — begins inventory, or arms the reader when using a Trigger
  preset.
- **Stop / Disarm** — stops inventory or disarms the hardware trigger.
- **Clear Log** — clears the current on-screen session. It does not delete an
  existing `TagLog.csv`.
- **Export CSV** — saves the current on-screen read log.
- **Close** — closes FlexWedge.

### Last Read — EPC Decode

The last-read panel can display:

- **Injected** — exact text sent to the focused application.
- **EPC** — raw EPC hexadecimal value.
- **TID** — tag identifier, when enabled and available.
- **Scheme** — decoded EPC scheme.
- **GTIN-14** — decoded trade-item identifier when applicable.
- **Serial** — decoded serial value when applicable.
- **EPC URI** — EPC identity URI.
- **RSSI** — received signal strength in dBm.
- **Antenna** — reader antenna number.
- **Lock** — lock or write-protection result when checked.

### Read Log

The session table records time, EPC, TID, scheme, GTIN-14, RSSI, and antenna.
The on-screen log retains up to 10,000 rows.

## 4. First connection

1. Connect the P35U to the laptop by USB.
2. Wait for Windows to recognize it.
3. Start FlexWedge.
4. Select the P35U COM port. If uncertain, open Windows Device Manager and look
   under **Ports (COM & LPT)**.
5. Click **Connect**.
6. Wait for the status to show **Connected — COMx**.

If connection fails, see Section 13.

## 5. Basic reading

For the first test:

1. Select the **Standard** preset.
2. Keep only one tag near the reader.
3. Click **Start Reading**.
4. Present the tag.
5. Confirm that EPC and RSSI appear.
6. Click **Stop**.

Standard mode reads continuously until stopped. Its default one-second dedupe
window prevents the same tag from being output on every inventory cycle.

## 6. Presets

Presets configure power, timing, buzzer behavior, deduplication, session, TID,
lock checking, and trigger behavior.

| Preset | Intended use | Main behavior |
|---|---|---|
| **Trigger - 1X** | Commanded single verification | Arm the reader; one physical SCAN reads one tag and stops |
| **Trigger - Continuous** | Commanded inventory interval | Physical SCAN toggles continuous reading |
| **Ghost** | Quiet, close-range QC | 13 dBm, silent, one read, TID and lock checking |
| **Stealth** | Audible close-range QC | 13 dBm, beep, one read, TID and lock checking |
| **Aware** | Moderate-range controlled reading | 20 dBm with a three-second same-tag cooldown |
| **Standard** | Everyday general use | 20 dBm, continuous, one-second same-tag cooldown |
| **Dense** | Crowded tag populations | 27 dBm, fast timing, Session 2 |
| **Hammer** | Maximum unrestricted inventory | 27 dBm, Session 0, no dedupe |

### Trigger presets

With a Trigger preset:

1. Click **Arm Trigger**.
2. Press the physical **SCAN** control on the P35U.
3. In **Trigger - 1X**, one tag is read and the reader returns to the armed
   state.
4. In **Trigger - Continuous**, SCAN starts reading and another SCAN stops it.
5. Click **Disarm** when finished.

The SCAN control is ignored while an active TID/lock quality check is finishing.

### Custom presets

The Presets menu provides two custom slots. Use **Save current settings as…**
to name and populate a slot. Empty slots remain disabled.

Other menu actions include:

- save the current preset as the startup default
- rename custom presets
- reset selected factory presets
- restore factory defaults

Custom presets capture reader and QC settings. They do not capture every Basic
Output Formatting field; save those separately with **Save Settings**.

## 7. Output formatting

### Append after tag

- **Nothing** — inject text only.
- **Tab** — inject text, then send Tab.
- **Enter** — inject text, then send Enter.

### Prefix and suffix

Enable **Add Prefix** or **Add Suffix**, then enter fixed ASCII text. The prefix
is placed before the selected tag value and the suffix after it.

### Output formats

- **HEX** — raw EPC hexadecimal string.
- **GTIN-14** — fourteen-digit GTIN when available.
- **EAN-13** — thirteen-digit EAN representation when available.
- **UPC-A** — twelve-digit UPC-A when possible; otherwise falls back to EAN-13.
- **UPC-A (as EAN-13)** — UPC-A represented with a leading zero.
- **GTIN-14 + Serial** — GTIN-14 and EPC serial joined by the configured
  delimiter.

If the chosen translation is unavailable for a tag, FlexWedge falls back to an
available representation rather than discarding the read. A UPC-A fallback is
identified in the status area.

### Display spaces

**Display Spaces** formats HEX output in byte pairs:

```text
30 34 2A 7C C8 44 C7 D0 F3 6A 06 82
```

It does not add spaces to GTIN, EAN, or UPC output.

### Timestamp

Enable **Add Timestamp** to append the current date and time. The **Delimiter**
field controls the text placed before the timestamp and is also used by the
GTIN-14 + Serial format.

### Combine Multiple Reads and deduplication

**Combine Multiple Reads** allows reads to be accumulated rather than treated
as isolated output events. **Dedupe window (ms)** suppresses repeated output of
the same EPC during the configured interval. It does not prevent the reader
from physically seeing the tag.

## 8. Sending a read to another application

FlexWedge sends output wherever Windows keyboard focus is located.

1. Connect and start or arm FlexWedge.
2. Click the destination field in Command Pilot, Excel, Notepad, or another
   application.
3. Present the RFID tag or press SCAN.
4. Do not click back into FlexWedge until the output is injected.

If the value appears in the wrong place, the wrong window had focus.

## 9. TID reads

TID is a manufacturer-assigned tag identifier stored in the TID memory bank.
It is separate from the EPC.

Ghost and Stealth enable TID reads. It can also be enabled under:

**Settings → Scan Timing → Read TID after each scan**

The P35U SDK performs TID as a separate operation:

1. EPC inventory identifies the tag.
2. Inventory stops.
3. FlexWedge requests TID memory.
4. The result returns through the normal tag-data callback.

Keep the same tag stationary in the RF field until TID finishes. The display
may briefly show **Reading…**, followed by the TID or **— (not available)**.

A missing TID does not necessarily mean the EPC read failed. The tag may have
moved, the response may have timed out, or the tag/firmware may not provide the
requested data.

## 10. Lock-status verification

Lock checking can be enabled under:

**Settings → Scan Timing → Verify tag lock status after each scan**

Possible displays include:

- **Permalocked** — memory is permanently locked.
- **Locked** — memory is write-protected.
- **Unlocked** — tag memory is writable; displayed as a warning.
- **Unknown** — no recognized status was returned.
- **Tag not in read range — rescan for lock status**

Keep the tag in the field until verification completes. Lock status is not
available in Trigger - Continuous mode; use Trigger - 1X, Ghost, or Stealth for
that verification.

## 11. Power and read range

Transmit power ranges from 13 to 27 dBm:

- lower power reduces read range and unwanted nearby reads
- higher power increases range and can read more surrounding tags

Select the value and click **Set**. Start at 13 dBm for close-range verification
and increase only when necessary.

## 12. Logs and exports

### Session export

Click **Export CSV** to save the current in-memory session. This export reflects
the values currently displayed in the table, including a TID added after the
initial EPC read.

### Automatic log

When **Log reads to file (TagLog.csv)** is enabled, FlexWedge appends reads to
`TagLog.csv` beside the application.

The automatic row may be written before the separate TID request finishes.
Therefore, the screen and an Export CSV file may show TID even when that
automatic `TagLog.csv` row has a blank TID.

## 13. Common user problems

### Connect is unavailable

- Confirm `AsReaderP3xU.dll` is beside the application.
- Confirm all files were extracted from the ZIP.
- Do not run the application directly inside the ZIP.

### “Assembly from a network location” or “Operation is not supported”

Windows has blocked the downloaded DLL:

1. Close FlexWedge.
2. Right-click `AsReaderP3xU.dll`.
3. Select **Properties**.
4. Check **Unblock**, then click **Apply**.

Technical Support can also unblock it with PowerShell.

### COM port does not appear

- Confirm the reader is connected.
- Wait for Windows device installation to finish.
- Click Refresh.
- Check Device Manager → Ports (COM & LPT).
- Try another USB cable or port.

### Connection fails

- Close any AsReader sample program or second FlexWedge instance.
- Confirm the selected COM port.
- Disconnect and reconnect USB, then reopen FlexWedge.

### Connected but no tags appear

- Place one known-good tag close to the reader.
- Use Standard at 20 dBm or temporarily increase power.
- Disable restrictive EPC filters.
- Set RSSI threshold to −99 to disable threshold filtering.
- Confirm reading is started rather than merely connected.

### TID is blank

- Use Ghost or Stealth, or enable Read TID in Settings.
- Keep one tag stationary in the field.
- Avoid multiple nearby tags.
- Retry at close range.

### Lock shows timeout or unavailable

Keep the tag in the field throughout QC and rescan. Do not use Trigger -
Continuous when lock status is required.

### Output goes to the wrong application

Select the intended destination field immediately before presenting the tag.

## 14. Ending a session

1. Click **Stop** or **Disarm**.
2. Export the session if required.
3. Click **Disconnect**.
4. Click **Close**.
5. Unplug the reader only after the application has closed.

## 15. Scope and limitations

- FlexWedge reads and decodes RFID; it does not itself perform the complete
  Command Pilot barcode/RFID verification workflow.
- The displayed EPC URI is not the same as a resolved GS1 Digital Link URL.
- Per-read RF frequency is not available from the current SDK.
- AsReader officially supports its SDK in C#/.NET; this reference application
  uses pythonnet.
- TID and lock checks are asynchronous and require the tag to remain present.
