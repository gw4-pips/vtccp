# Full Parameter Capture Output Example

**Source:** VCCS RFID FlexWedge Pro — "Last Read" decode card + debug log
**Preset:** Stealth (single-read, TID enabled, lock check enabled)
**Tag:** SGTIN-96, Impinj Monza R6

---

## UI Decode Card Output

The FlexWedge "Last Read — EPC Decode" panel shows:

```
Injected:  30342A7CC844C7D0F36A0676
EPC:       30342A7CC844C7D0F36A0676
TID:       E28011920008C7C
Scheme:    SGTIN-96
GTIN-14:   00696114704318
Serial:    72803288694
EPC URI:   urn:epc:id:sgtin:0696114.70431.72803288694
RSSI:      -35.0 dBm
Antenna:   1
Lock:      🔒 Permalocked
```

---

## Internal Python Dict (decoder.decode_epc output)

```python
{
    'raw':        '30342A7CC844C7D0F36A0676',
    'scheme':     'SGTIN-96',
    'gtin14':     '00696114704318',
    'gcp':        '0696114',
    'item_ref':   '704318',       # 6-digit item reference (indicator stripped)
    'indicator':  '0',            # packaging level indicator
    'serial':     '72803288694',
    'epc_uri':    'urn:epc:id:sgtin:0696114.70431.72803288694',
    'filter_val': '1',            # EPC filter digit (retail consumer trade item)
    'error':      None,
}
```

---

## Read Log Row (CSV column order)

```
Time,EPC (hex),TID,Scheme,GTIN-14,RSSI,Ant
12:46:16.760,30342A7CC844C7D0F36A0676,E28011920008C7C,SGTIN-96,00696114704318,-35.0 dBm,1
```

---

## Injected Keystrokes (HEX mode with Enter)

```
30342A7CC844C7D0F36A0676<Enter>
```

In GTIN-14 mode:
```
00696114704318<Enter>
```

In GTIN-14 + Serial mode (with Tab delimiter):
```
00696114704318	72803288694<Enter>
```

---

## Debug Log Excerpt (debug.log)

```
[12:46:15.341] SetDelegate registered
[12:46:15.342] StartInventory called (rssi=True, maxTags=1, maxSecs=0, maxCycles=0, ant=1)
[12:46:15.891] TAG epc=30342A7CC844C7D0F36A0676 tid=(none) rssi=-35 ant=1
[12:46:15.892] cb_complete: hardware auto-stop (mtnu limit reached)
[12:46:15.892] QC worker: starting (epc=30342A7CC844C7D0F36A...)
[12:46:15.894] read_tid: StopInventory before ReadMemory
[12:46:15.944] ReadMemory called (MEM_TID, offset=0, len=4, pwd=0)
[12:46:16.287] read_tid _cb_tag path: attr=tid → 'E28011920008C7C'
[12:46:16.287] read_tid 30342A7CC844… → E28011920008C7C
[12:46:16.289] CheckTagStatus called
[12:46:16.512] DLL success callback: 40
[12:46:16.513] QC worker: lock=PermaLock tid=E28011920008C7C
```

---

## TagRead Dataclass (internal Python object)

```python
@dataclass
class TagRead:
    raw_epc:   str   = '30342A7CC844C7D0F36A0676'
    pc:        int   = 0x3000   # Protocol Control word
    rssi_raw:  int   = -35      # dBm after two's complement correction
    antenna:   int   = 1
    tid:       str   = 'E28011920008C7C'  # filled by QC worker
    timestamp: float = 1754913975.341     # time.time() float

    # Derived (added by decode step):
    #   decoded['scheme']  = 'SGTIN-96'
    #   decoded['gtin14']  = '00696114704318'
    #   decoded['serial']  = '72803288694'
    #   decoded['epc_uri'] = 'urn:epc:id:sgtin:0696114.70431.72803288694'
```

---

## Config State for Stealth Preset

```python
Config(
    port               = 'COM4',
    baud               = 115200,
    power_dbm          = 13,          # low power for short-range precision read
    read_time_ms       = 200,
    idle_time_ms       = 500,
    session            = 1,           # S1
    anti_collision     = 'DynamicQ',
    stop_after_first   = True,        # single-read mode
    combine_reads      = True,
    combine_timeout_ms = 999999,      # effectively no repeat in single-read mode
    rssi_threshold     = -99,         # accept all
    buzzer_level       = 'HIGH',
    read_tid           = True,        # TID enabled
    check_lock         = True,        # lock check enabled
    output_format      = 'HEX',
    append_key         = 'Enter',
)
```
