# ASR-P35U Firmware Log

## Unit KE00048

| Date       | Component    | Version before | Version after  | Tool used              |
|------------|--------------|----------------|----------------|------------------------|
| 2026-08-10 | Main FW      | 1.2.0          | **1.8.0**      | AsReaderP3xU_Demo 1.3.0 |
| 2026-08-10 | RFID module  | RED4S_v2.2.1_C | **RED4S_v2.2.2_K_SD** | AsReaderP3xU_Demo 1.3.0 |

### Static info (as of 2026-08-10, post-update)
| Field        | Value          |
|--------------|----------------|
| S/N          | KE00048        |
| HW Version   | 1.0.2          |
| SDK Version  | 1.3.0 (2026-02-13) |
| COM port     | COM4 (Windows 11 laptop) |
| VID/PID      | 0x339C / 0x271B |

### Operating parameters (factory defaults, post-update)
| Parameter        | Value        |
|------------------|--------------|
| TX Power         | 13 dBm       |
| Read Time        | 380 ms       |
| Idle Time        | 500 ms       |
| Region           | REGION_US    |
| Session          | SESSION_S1   |
| Target           | A            |
| Anti-collision   | DynamicQ     |
| Q Start/Min/Max  | 4 / 0 / 8    |
| Buzzer           | HIGH         |
| Frequency        | Automatic (CHANNEL_25 base) |

---

## How to update firmware (future reference)

**Prerequisites**
- `AsReaderP3xU_Demo.exe` and `AsReaderP3xU_Demo.exe.config` in the same folder as `AsReaderP3xU.dll`
- Internet access (firmware files pulled from AsReader's servers)
- Wedge app closed (COM port must be free)

**Steps**
1. Run `AsReaderP3xU_Demo.exe`
2. Click **Search** → select correct COM port → **Connect**
3. Click **Get Ver** (all three) to confirm current versions
4. **Main FW:** click **Get Files** → select bin → click **Update** → confirm → wait for reboot prompt → reboot
5. **RFID FW:** click **Get Files** (RFID FW side) → select hex → click **Update** → confirm → reboot
6. After reboot, click **Get Ver** again to confirm new versions
7. Log the update in this file

**Known issue — wrong COM port**
The demo app defaults to the lowest COM number. The P35U may enumerate on a higher port (COM4 on the dev laptop). If all "Get" commands fail after connecting, disconnect, select a higher COM port, and reconnect.

**Firmware source URLs** (fetched automatically by the demo app)
- Main FW:     `https://camera.asreaderapps.com/mrxapp/MRX-D3Fw/fw/P35U.xml`
- RFID module: `https://apps.asreaderapps.com/mrxapp/MRX-D3Fw/RFIDfw/P3xURFIDModule.xml`
