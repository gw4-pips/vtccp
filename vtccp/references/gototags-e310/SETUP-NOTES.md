GoToTags Desktop E310 UHF RFID Reader — Setup Notes
Version 1.0 | 2026-07-13

─────────────────────────────────────────────────────────────
HARDWARE
─────────────────────────────────────────────────────────────
  SKU:       TDLP3LCFPP
  Chipset:   Impinj E310 (Indy R2000 family)
  Antenna:   Internal 2 dBi circularly polarized panel
  Frequency: 860–960 MHz (GS1 UHF band, North America 902–928 MHz)
  TX power:  1–33 dBm adjustable
  Range:     Up to 1 metre
  USB:       Select cable type at checkout (USB-A or USB-C)
  FCC cert:  Yes
  Price:     $199.93 (qty 1) from store.gototags.com

─────────────────────────────────────────────────────────────
DRIVER INSTALLATION (Windows)
─────────────────────────────────────────────────────────────
  Driver:  CDM212364_Setup.zip  (FTDI Combined Driver Model, included here)
  
  Steps:
    1. Extract CDM212364_Setup.zip.
    2. Run CDM212364_Setup.exe as Administrator.
    3. Accept the license and complete the install wizard.
    4. Plug in the E310 via USB.
    5. Open Device Manager → Ports (COM & LPT).
       Confirm a "USB Serial Port (COMx)" entry appears.
    6. Note the COMx number — that is the portName for ConnectAsync().
  
  If the device appears under "Other devices" instead of Ports,
  right-click → Update driver → Browse to the extracted CDM folder.

─────────────────────────────────────────────────────────────
VTCCP INTEGRATION
─────────────────────────────────────────────────────────────
  Factory method:  EpcReaderFactory.CreateGoToTagsE310()
  
  Usage:
    var reader = EpcReaderFactory.CreateGoToTagsE310();
    await reader.ConnectAsync("COM5");                   // adjust port
    var epcs = await reader.TriggerInventoryAsync(TimeSpan.FromSeconds(3));
    await reader.DisconnectAsync();
  
  Protocol:    GoToTags UHF RFID Reader Communication Protocol rev 5-30-23
  Baud rate:   115 200 8N1 (factory default, no change needed)
  Frame:       [ 0xFF | DataLen | CmdCode | Data | CRC_Hi | CRC_Lo ]
  CRC:         CRC-16/CCITT (init=0xFFFF, poly=0x1021), big-endian
  Ping:        Single Tag Inventory (0x21) with 200 ms timeout —
               any response (tag or no-tag status) confirms APP-layer readiness
  Inventory:   0x21 in 150 ms slices, de-duplicates EPCs across slices

─────────────────────────────────────────────────────────────
PROTOCOL REFERENCE
─────────────────────────────────────────────────────────────
  Source:  GoToTags GitLab (public repo)
    gitlab.com/gototags/public
    → UHF RFID / Readers / GoToTags / docs /
      "GoToTags UHF RFID Reader Communication Protocol - 5-30-23.pdf"

─────────────────────────────────────────────────────────────
KNOWN LIMITATIONS / HARDWARE TODO
─────────────────────────────────────────────────────────────
  - CRC polynomial confirmed from Appendix 1 of the protocol PDF.
    Verify on first hardware test: send a known command and check the
    CRC bytes match what the reader acknowledges.
  
  - No-tag status code not verified (need hardware to capture the
    exact two-byte status when 0x21 times out with no tag present).
    Current code treats any non-0x0000 status as "no tag" — correct
    in practice, but the exact error code is hardware-confirmed later.
  
  - TX power and frequency region are factory defaults (North America,
    ~23 dBm). Adjustable via APP-layer system commands (section 9 of
    the protocol doc) if read range needs tuning.
