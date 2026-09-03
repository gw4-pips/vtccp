"""
Copyright © 2026 VCCS. All rights reserved.
RFID FlexWedge™ Pro — proprietary software of VCCS.

Settings persistence — saved as JSON next to the executable.
"""
import json
import os
from dataclasses import dataclass, asdict, field


CONFIG_FILENAME = 'rfid_wedge_config.json'

# Default parameter set for a blank custom preset slot.
# Used to initialise new slots and to back-fill any keys missing from an
# older config file (forward-compatible schema migration).
_CUSTOM_PRESET_BLANK: dict = dict(
    name='',                       # empty = slot not yet configured
    power_dbm=20,
    buzzer_level='HIGH',
    combine_reads=True,
    combine_timeout_ms=1000,
    stop_after_first=False,
    read_time_ms=300,
    idle_time_ms=100,
    session=1,
    anti_collision='DynamicQ',
    anti_collision_start_q=4,
    anti_collision_min_q=0,
    anti_collision_max_q=8,
    rssi_threshold=-99,
    read_tid=True,
    check_lock=True,
)


@dataclass
class Config:
    # Connection
    port: str   = ''
    baud: int   = 115200

    # Injection / formatting
    append_key:         str  = 'Enter'   # 'Nothing', 'Tab', 'Enter'
    prefix:             str  = ''
    suffix:             str  = ''
    display_spaces:     bool = False     # space every 2 hex chars
    combine_reads:      bool = True      # deduplicate consecutive reads
    combine_timeout_ms: int  = 999999   # ms before same tag can fire again
    add_timestamp:      bool = False
    delimiter:          str  = ''        # between tag data and timestamp

    # Tag translation
    output_format:  str  = 'HEX'  # 'HEX' | 'GTIN14' | 'EAN13' | 'UPCA' | 'UPCA_EAN13' | 'GTIN_SN'
    include_filter: bool = False   # prepend EPC filter nibble

    # Reader — basic (also on main UI)
    power_dbm: int = 20

    # Reader — advanced (Advanced Settings dialog)
    read_time_ms:           int  = 300    # antenna read time per cycle (ms)
    idle_time_ms:           int  = 100    # idle time between cycles (ms)
    session:                int  = 1      # 0=S0, 1=S1, 2=S2, 3=S3
    anti_collision:         str  = 'DynamicQ'   # 'DynamicQ' or 'FixedQ'
    anti_collision_start_q: int  = 4
    anti_collision_min_q:   int  = 0
    anti_collision_max_q:   int  = 8
    rssi_threshold:         int  = -99   # dBm; -99 = off (accept everything)
    buzzer_level:           str  = 'HIGH'  # 'OFF', 'LOW', or 'HIGH'
    stop_after_first:       bool = False  # stop inventory after first tag read

    # Tag filter — EPC packaging-level checkboxes
    # select_filter_values: list of int (0–7); empty = filter disabled.
    # Reserved values (3, 5) are never written here; UI prevents it.
    select_mask_enabled:   bool = False
    select_filter_values:  list = field(default_factory=list)

    # Legacy fields — kept so old JSON loads don't throw; retired from active use.
    # select_mask_prefix is migrated away on load (see Config.load).
    select_mask_bank:    str  = 'EPC'    # unused — always EPC
    select_mask_prefix:  str  = ''       # retired; migration clears this

    # QC — tag interrogation after each scan
    read_tid:   bool = False  # read TID memory bank via ReadMemory; shown in card
    check_lock: bool = True   # call CheckTagStatus after each scan; alert on UnLock

    # Logging
    log_enabled: bool = True

    # Startup default preset — the preset applied automatically on launch.
    # Stores the key name (e.g. 'Standard', 'custom_1').  Changed via
    # Presets → "Save current selection as startup default".
    default_preset: str = 'Standard'

    # User-defined custom presets (2 slots; name='' means slot is unused)
    custom_presets: list = field(default_factory=lambda: [
        dict(_CUSTOM_PRESET_BLANK),
        dict(_CUSTOM_PRESET_BLANK),
    ])

    def save(self, path: str = ''):
        path = path or self._default_path()
        d = asdict(self)
        with open(path, 'w') as f:
            json.dump(d, f, indent=2)

    @classmethod
    def load(cls, path: str = '') -> 'Config':
        path = path or cls._default_path()
        cfg = cls()
        if os.path.exists(path):
            try:
                with open(path) as f:
                    d = json.load(f)
                for k, v in d.items():
                    if hasattr(cfg, k):
                        setattr(cfg, k, v)
                # Back-compat: old JSON stored buzzer as a bool
                if 'buzzer' in d and 'buzzer_level' not in d:
                    cfg.buzzer_level = 'HIGH' if d['buzzer'] else 'OFF'
                # Back-compat: migrate upc_ean bool → output_format enum
                if 'upc_ean' in d and 'output_format' not in d:
                    cfg.output_format = 'GTIN14' if d['upc_ean'] else 'HEX'
                # Back-compat: migrate old 'GTIN' enum value → 'GTIN14'
                if cfg.output_format == 'GTIN':
                    cfg.output_format = 'GTIN14'
                # Back-compat: retire select_mask_prefix raw-hex field.
                # If it was non-empty the user had a manual hex filter; we
                # cannot reliably parse it back to filter values, so we clear
                # it and leave the filter disabled (safe default).
                if d.get('select_mask_prefix', ''):
                    cfg.select_mask_prefix = ''
                    cfg.select_mask_enabled = False
                    cfg.select_filter_values = []
                # Ensure select_filter_values is a clean list of ints 0–7
                if isinstance(cfg.select_filter_values, list):
                    cfg.select_filter_values = [
                        int(v) for v in cfg.select_filter_values
                        if isinstance(v, (int, float)) and 0 <= int(v) <= 7
                    ]
                else:
                    cfg.select_filter_values = []
                # Ensure custom_presets is always a 2-slot list with all keys
                slots = cfg.custom_presets if isinstance(cfg.custom_presets, list) else []
                validated = []
                for i in range(2):
                    base = dict(_CUSTOM_PRESET_BLANK)
                    if i < len(slots) and isinstance(slots[i], dict):
                        base.update(slots[i])
                    validated.append(base)
                cfg.custom_presets = validated
            except Exception:
                pass  # fall back to defaults silently
        return cfg

    @staticmethod
    def _default_path() -> str:
        base = getattr(__import__('sys'), 'frozen', False) and \
               getattr(__import__('sys'), '_MEIPASS', None)
        folder = os.path.dirname(base) if base else os.path.dirname(
            os.path.abspath(__file__))
        return os.path.join(folder, CONFIG_FILENAME)
