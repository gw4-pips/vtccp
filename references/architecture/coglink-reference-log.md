# Coglink Reference Log

## URLs logged

### 1 — DataMan 380 Modular Vision Tunnel press release (2024-03-27)
**URL**: https://www.prnewswire.com/news-releases/cognex-launches-flexible-data-driven-tunnel-solution-powered-by-the-dataman-380-barcode-reader-302100146.html  
**Logged**: 2026-05-21  
**Source**: user-supplied

**Content summary**: Cognex launch announcement for the 380 Modular Vision Tunnel.
- DM380 is described as "Cognex's widest field of view barcode reader"
- Tunnel uses as few as 4 readers per side, achieves >99% read rate
- Supported by Edge Intelligence Tunnel Manager (real-time analytics, vendor compliance)
- AI-assisted decoding highlighted; deep-of-field / FOV focus
- Quote from StreamTech Engineering: "It basically includes everything you could ever want"
- Contact: Carl Gerst (EVP Vision & ID Products); Patrick Bradford (StreamTech)

**Coglink mention**: **None** — the word "Coglink" does not appear in this press release.
The user noted "COMLINK was my typo" (for Coglink). This PR covers the DM380 logistics
reader, not a connectivity-focused release. Coglink may be a feature of the DM380 but
is not described here. A DM380-specific hardware reference manual would be needed to
confirm Coglink presence and pinning on that model.

---

## Coglink — what is known from existing reference materials

Source: `references/manuals/cognex/reference-manual-DM390-25.4.1.2.pdf` (extracted to `/tmp/dm390.txt`)

### Definition
Cognex's proprietary label for the USB-C port on the DM390 / DM395 series.
The DM390 hardware diagram labels item #10 as **"Coglink/USB-C status LED"**.
No separate "Coglink" datasheet or spec sheet has been located; all information
comes from the DM390 reference manual.

### DM390 / DM395 USB-C (Coglink) — three operating modes

| Mode | Windows presentation | Notes |
|---|---|---|
| Emulated serial (USB-COM) | Generic USB-COM in Device Manager | Requires `port.DtrEnable = true` in client |
| Emulated Ethernet | Virtual network adapter | Fixed IP **192.168.111.2 / 24** |
| HID keyboard | Emulated keyboard | Language configurable via DMST |

**Power warning** (DM390 manual verbatim): *"Do not power the reader exclusively over USB.
Any load to the system might cause it to reboot."*

### DMCC commands applicable to DM390 USB-C

| Command | Platforms | Notes |
|---|---|---|
| `COM.USB-HID-POS` | DM390 (explicit) | Enables HID-POS mode; reboot required |
| `COM.USB-REQUIRE-DTR` | Older set only (not DM390) | DTR toggle for USB-COM |
| `COM.USB-MODE` | Older set only (not DM390) | 0=COM, 1=HID |

The emulated-serial and emulated-Ethernet modes on DM390/395 are configured through
DMST's Communication Settings panel, not via DMCC commands.

### Coglink connection identity strings (from Comms & Programming guide)

When using communication scripts on the device:
- USB-COM serial: `localName = "COM USB"`, `peerName = "COM USB"`
- HID: `localName = "keybrd"`, `peerName = "keybrd"`
- Ethernet: `localName = "<IP>:<PORT>"`, `peerName = "<PEER_IP>:<PORT>"`

### VTCCP ConnectionMedium wiring (implemented 2026-05-21)

`DeviceConfig.ConnectionMedium` enum: `Auto | GigE | USBEthernet | USBCOM`  
Auto-resolve logic: `192.168.111.x` or `169.254.x.x` → `"USB-Ethernet"`; all others → `"GigE"`.  
DM395V at 192.168.111.2 will resolve correctly with zero config change.

---

## Still needed

- DM380 hardware reference manual — to confirm Coglink presence and pinout on that model
- DM395V hardware reference manual — sensor spec confirmation, Coglink port details
  (DM390 manual confirms DM395 = 5MP / 2448×2048, but full DM395V hardware reference
  not yet in the reference library)
