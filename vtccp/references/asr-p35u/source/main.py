"""
Copyright © 2026 VCCS. All rights reserved.
RFID FlexWedge™ Pro — proprietary software of VCCS.

Keyboard wedge replacement for the MTI RFID ME / RFID Wedge tool.
Runs as a standalone Windows .exe (packaged with PyInstaller).
"""
import array
import csv
import datetime
import io
import logging
import math
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import wave
import winsound

try:
    import pystray
    from PIL import Image as _PILImage, ImageDraw as _PILDraw
    _TRAY_AVAILABLE = True
except ImportError:
    _TRAY_AVAILABLE = False

import injector as inj
from config import Config
from decoder import decode_epc, format_epc_for_inject, gtin14_check_ok
from reader import E310Reader, TagRead, dll_present, pythonnet_present

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s %(name)s: %(message)s')

# ── Suppress native stdout (FD 1) ────────────────────────────────────
# This is a GUI app; all diagnostic output goes to debug.log.
# The AsReader DLL writes "write error" directly to the OS file descriptor
# (not Python's sys.stdout), so only an fd-level redirect silences it.
try:
    _nul_fd = os.open(os.devnull, os.O_WRONLY)
    os.dup2(_nul_fd, 1)
    os.close(_nul_fd)
    sys.stdout = open(os.devnull, 'w')
except OSError:
    pass

APP_TITLE   = 'RFID FlexWedge\u2122 Pro'   # ™ = U+2122
APP_VERSION = '1.0.0'


# ── Alert chime (Unlocked) ────────────────────────────────────────────
def _make_unlock_chime(sample_rate: int = 44100) -> bytes:
    """Synthesise a two-note descending bell chime as in-memory WAV bytes.

    Design: A5 (880 Hz) → D5 (587 Hz), staggered by 220 ms, each with an
    exponential bell-decay envelope.  Total duration ~0.7 s.  The descending
    figure has a gentle "heads-up" quality — noticeable without being alarming.
    Played via winsound.SND_MEMORY so no temp file is written to disk.
    """
    total   = int(sample_rate * 0.7)
    onset2  = int(sample_rate * 0.22)        # D5 starts 220 ms after A5
    samples = array.array('h', [0] * total)

    for i in range(total):
        t  = i / sample_rate
        # A5 — 880 Hz, fast bell decay (τ ≈ 110 ms)
        e1 = math.exp(-t * 9.0)
        s1 = e1 * math.sin(2 * math.pi * 880 * t)
        # D5 — 587 Hz, slightly slower decay (τ ≈ 150 ms), delayed onset
        t2 = (i - onset2) / sample_rate if i >= onset2 else 0.0
        e2 = math.exp(-t2 * 6.5) if i >= onset2 else 0.0
        s2 = e2 * math.sin(2 * math.pi * 587 * t2)
        # Mix at 45 % amplitude each to stay well below clipping
        samples[i] = int((s1 + s2) * 0.45 * 32767)

    buf = io.BytesIO()
    with wave.open(buf, 'wb') as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(sample_rate)
        w.writeframes(samples.tobytes())
    return buf.getvalue()


_UNLOCK_CHIME_WAV: bytes = _make_unlock_chime()

# ── Colour palette (matches EPC decoder web tool) ────────────────────
BG          = '#eef2f7'
SURFACE     = '#ffffff'
ACCENT      = '#2b6cb0'
ACCENT_DK   = '#2c5282'
GREEN       = '#276749'   # connected-idle dot
GREEN_LT    = '#c6f6d5'
GREEN_ACTIVE = '#4ade80'  # bright chartreuse-green — active reading dot
YELLOW      = '#d97706'   # amber — paused dot
RED         = '#c53030'
RED_LT      = '#fed7d7'
MUTED       = '#4a5568'
BORDER      = '#d1dae6'
MONO_FONT   = ('Consolas', 10)
UI_FONT     = ('Segoe UI', 9)
UI_FONT_B   = ('Segoe UI', 9, 'bold')
TITLE_FONT  = ('Segoe UI', 12, 'bold')

# ── Scan presets ─────────────────────────────────────────────────────
# Factory preset definitions — immutable reference; used for per-preset reset.
# Tunable params (power, buzzer, timing, RSSI, dedup, QC) may be adjusted
# on factory presets without losing the preset label.
# Locked params (stop_after_first, session, anti-collision) define the
# fundamental character of each preset and cannot be changed while it is active.
FACTORY_PRESETS = {
    'Trigger - 1X': dict(
        power_dbm=13, buzzer_level='HIGH',
        combine_reads=True,  combine_timeout_ms=999999,
        stop_after_first=True,
        read_time_ms=200, idle_time_ms=100, session=1,
        anti_collision='DynamicQ', anti_collision_start_q=4,
        anti_collision_min_q=0,   anti_collision_max_q=8,
        read_tid=False, trigger_mode='oneshot',
        label='🔘 Trigger - 1X — SCAN button reads one tag then stops',
    ),
    'Trigger - Continuous': dict(
        power_dbm=13, buzzer_level='HIGH',
        combine_reads=True,  combine_timeout_ms=999999,
        stop_after_first=False,
        read_time_ms=300, idle_time_ms=100, session=1,
        anti_collision='DynamicQ', anti_collision_start_q=4,
        anti_collision_min_q=0,   anti_collision_max_q=8,
        read_tid=False, trigger_mode='continuous',
        label='🔄 Trigger - Continuous — SCAN button starts/stops continuous reading',
    ),
    'Ghost': dict(
        power_dbm=13, buzzer_level='OFF',
        combine_reads=True,  combine_timeout_ms=999999,
        stop_after_first=True,
        read_time_ms=100, idle_time_ms=500, session=1,
        anti_collision='DynamicQ', anti_collision_start_q=4,
        anti_collision_min_q=0,   anti_collision_max_q=8,
        read_tid=True,
        label='🔇 Ghost — lowest power, silent, one read then stop',
    ),
    'Stealth': dict(
        power_dbm=13, buzzer_level='HIGH',
        combine_reads=True,  combine_timeout_ms=999999,
        stop_after_first=True,
        read_time_ms=100, idle_time_ms=500, session=1,
        anti_collision='DynamicQ', anti_collision_start_q=4,
        anti_collision_min_q=0,   anti_collision_max_q=8,
        read_tid=True,
        label='🔔 Stealth — lowest power, beep on read, one read then stop',
    ),
    'Aware': dict(
        power_dbm=20, buzzer_level='HIGH',
        combine_reads=True,  combine_timeout_ms=3000,
        stop_after_first=False,
        read_time_ms=200, idle_time_ms=200, session=1,
        anti_collision='DynamicQ', anti_collision_start_q=4,
        anti_collision_min_q=0,   anti_collision_max_q=8,
        read_tid=False,
        label='👀 Aware — moderate power, 3 s cooldown between same tag',
    ),
    'Standard': dict(
        power_dbm=20, buzzer_level='HIGH',
        combine_reads=True,  combine_timeout_ms=1000,
        stop_after_first=False,
        read_time_ms=300, idle_time_ms=100, session=1,
        anti_collision='DynamicQ', anti_collision_start_q=4,
        anti_collision_min_q=0,   anti_collision_max_q=8,
        read_tid=False,
        label='⚖️ Standard — balanced for everyday use',
    ),
    'Dense': dict(
        power_dbm=27, buzzer_level='HIGH',
        combine_reads=True,  combine_timeout_ms=500,
        stop_after_first=False,
        read_time_ms=500, idle_time_ms=50, session=2,
        anti_collision='DynamicQ', anti_collision_start_q=4,
        anti_collision_min_q=0,   anti_collision_max_q=8,
        read_tid=False,
        label='📦 Dense — high power, fast, for crowded tag environments',
    ),
    'Hammer': dict(
        power_dbm=27, buzzer_level='HIGH',
        combine_reads=False, combine_timeout_ms=0,
        stop_after_first=False,
        read_time_ms=500, idle_time_ms=0, session=0,
        anti_collision='DynamicQ', anti_collision_start_q=4,
        anti_collision_min_q=0,   anti_collision_max_q=8,
        read_tid=False,
        label='🔨 Hammer — max power, no filter, read everything non-stop',
    ),
}

# Parameters that are fixed when a factory preset is active.
# They define the preset's fundamental character (one-shot vs continuous,
# session behaviour, RF collision arbitration).
FACTORY_LOCKED_PARAMS = frozenset({
    'stop_after_first', 'session', 'trigger_mode',
    'anti_collision', 'anti_collision_start_q',
    'anti_collision_min_q', 'anti_collision_max_q',
})

# Stable internal keys for the two user-defined custom preset slots.
CUSTOM_SLOT_KEYS = ['Custom 1', 'Custom 2']


# ─────────────────────────────────────────────────────────────────────
class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title(f'{APP_TITLE}  v{APP_VERSION}')
    def __init__(self):
        super().__init__()
        self.title(f'{APP_TITLE}  v{APP_VERSION}')
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(820, 580)

        self.cfg = Config.load()
        self._reader: E310Reader | None = None
        self._log_rows: list[dict] = []
        self._last_epcs: dict[str, float] = {}
        self._inject_lock = threading.Lock()
        self._raw_buf: list[str] = []
        self._dll_was_missing: bool = False
        self._device_info: dict = {}   # populated after connect

        self._preset_var = tk.StringVar(value='Standard')
        self._buzzer_var = tk.StringVar(value=self.cfg.buzzer_level)

        self._tray_icon: 'pystray.Icon | None' = None
        self._tray_state: str = 'idle'     # 'idle' | 'reading' | 'alert'
        self._window_hidden: bool = False  # True while withdrawn to tray
        self._tray_hide_id: str | None = None  # pending after() ID for _tray_hide
        # Trigger-mode state — set by _apply_preset when a Trigger preset is active.
        # '' = normal (Start/Stop buttons); 'oneshot'/'continuous' = SCAN-button driven.
        self._trigger_mode:  str  = ''
        self._trigger_armed: bool = False  # True while armed for SCAN button presses
        self._qc_active:     bool = False  # True while a QC worker (TID/lock) is running

        self._setup_styles()
        self._build_menu()
        self._build_ui()
        self._refresh_ports()
        self._apply_cfg_to_ui()
        self._check_dll()

        self.protocol('WM_DELETE_WINDOW', self._on_close)
        self.bind('<Unmap>', self._on_unmap)   # minimize → hide to tray
        self.after(2000, self._periodic_port_refresh)
        self.after(600,  self._setup_tray)

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use('clam')
        s.configure('.',           background=BG,  font=UI_FONT)
        s.configure('TFrame',      background=BG)
        s.configure('TLabelframe', background=BG,  font=UI_FONT_B)
        s.configure('TLabelframe.Label', background=BG, foreground=ACCENT, font=UI_FONT_B)
        s.configure('TLabel',      background=BG,  font=UI_FONT)
        s.configure('TCheckbutton',background=BG,  font=UI_FONT)
        s.configure('TRadiobutton',background=BG,  font=UI_FONT)
        s.configure('TEntry',      font=MONO_FONT, fieldbackground=SURFACE)
        s.configure('TCombobox',   font=UI_FONT,   fieldbackground=SURFACE)
        s.configure('Treeview',    font=MONO_FONT, rowheight=22,
                    background=SURFACE, fieldbackground=SURFACE)
        s.configure('Treeview.Heading', font=UI_FONT_B, background=BG)
        s.map('Treeview', background=[('selected', ACCENT)])
        s.configure('Accent.TButton', background=ACCENT,
                    foreground='white', font=UI_FONT_B, padding=(12, 5))
        s.map('Accent.TButton',
              background=[('active', ACCENT_DK), ('pressed', ACCENT_DK)])
        s.configure('Stop.TButton', background=RED,
                    foreground='white', font=UI_FONT_B, padding=(12, 5))
        s.map('Stop.TButton',
              background=[('disabled', '#e8a0a0'), ('active', '#9b2c2c'), ('pressed', '#9b2c2c')],
              foreground=[('disabled', '#5a1a1a')])
        s.configure('Flat.TButton', background=BG, font=UI_FONT, padding=(8, 4))

    def _build_menu(self):
        menubar = tk.Menu(self)

        # ── Presets ──
        self._preset_menu = tk.Menu(menubar, tearoff=0)
        self._populate_preset_menu()
        menubar.add_cascade(label='Presets', menu=self._preset_menu)

        # ── Settings ──
        menubar.add_command(label='Settings', command=self._open_advanced)

        # ── Help ──
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label='View debug.log',
                              command=self._open_debug_log)
        help_menu.add_command(label='Open README',
                              command=self._open_readme)
        help_menu.add_separator()
        help_menu.add_command(label=f'About {APP_TITLE}…',
                              command=self._open_about)
        menubar.add_cascade(label='Help', menu=help_menu)

        self.config(menu=menubar)

    def _build_ui(self):
        # ── Top bar (connection) ──────────────────────────────────────
        top = ttk.Frame(self, padding=(12, 8))
        top.pack(side='top', fill='x')

        ttk.Label(top, text=APP_TITLE, font=TITLE_FONT,
                  foreground=ACCENT).pack(side='left', padx=(0, 20))

        ttk.Label(top, text='Port:').pack(side='left')
        self._port_var = tk.StringVar()
        self._port_cb = ttk.Combobox(top, textvariable=self._port_var,
                                     width=10, state='readonly')
        self._port_cb.pack(side='left', padx=(4, 2))

        ttk.Button(top, text='↻', width=3, style='Flat.TButton',
                   command=self._refresh_ports).pack(side='left', padx=(0, 10))

        self._connect_btn = ttk.Button(top, text='Connect',
                                       style='Accent.TButton',
                                       command=self._toggle_connect)
        self._connect_btn.pack(side='left', padx=(0, 12))

        self._status_canvas = tk.Canvas(top, width=12, height=12,
                                        bg=BG, highlightthickness=0)
        self._status_canvas.pack(side='left', padx=(0, 4))
        self._status_dot = self._status_canvas.create_oval(
            1, 1, 11, 11, fill=RED, outline='')

        self._status_var = tk.StringVar(value='Not connected')
        ttk.Label(top, textvariable=self._status_var,
                  foreground=MUTED).pack(side='left')

        # Preset indicator — always shows the active preset name
        self._preset_lbl_var = tk.StringVar(value='')
        ttk.Label(top, textvariable=self._preset_lbl_var,
                  foreground=ACCENT, font=('', 9, 'bold')).pack(side='left', padx=(10, 0))

        # Power — kept on main UI per user preference
        ttk.Label(top, text='Power (dBm):').pack(side='right', padx=(12, 4))
        self._power_var = tk.StringVar(value=str(self.cfg.power_dbm))
        pw_sb = ttk.Spinbox(top, from_=13, to=27, width=4,
                            textvariable=self._power_var, font=UI_FONT)
        pw_sb.pack(side='right', padx=(0, 2))
        ttk.Button(top, text='Set', style='Flat.TButton',
                   command=self._set_power).pack(side='right', padx=(0, 4))

        # ── Separator ─────────────────────────────────────────────────
        ttk.Separator(self, orient='horizontal').pack(fill='x')

        # ── Body: settings left, log right ───────────────────────────
        body = ttk.Frame(self, padding=(0, 0))
        body.pack(fill='both', expand=True)

        settings_frame = ttk.Frame(body, padding=(8, 8), width=260)
        settings_frame.pack(side='left', fill='y')
        settings_frame.pack_propagate(False)

        ttk.Separator(body, orient='vertical').pack(side='left', fill='y')

        right_frame = ttk.Frame(body, padding=(8, 8))
        right_frame.pack(side='left', fill='both', expand=True)

        self._build_settings(settings_frame)
        self._build_log(right_frame)

        # ── Debug raw-bytes panel (collapsed by default) ──────────────
        self._debug_visible = tk.BooleanVar(value=False)
        debug_bar = ttk.Frame(self, padding=(12, 2))
        debug_bar.pack(side='bottom', fill='x')
        ttk.Checkbutton(debug_bar, text='Show raw RX bytes (debug)',
                        variable=self._debug_visible,
                        command=self._toggle_debug).pack(side='left')
        self._debug_frame = ttk.Frame(self, padding=(12, 4))
        self._debug_text = tk.Text(self._debug_frame, height=5,
                                   font=('Consolas', 8), bg='#1a1a2e',
                                   fg='#a0f0a0', state='disabled',
                                   wrap='word')
        self._debug_text.pack(side='left', fill='both', expand=True)
        dbsb = ttk.Scrollbar(self._debug_frame, orient='vertical',
                             command=self._debug_text.yview)
        self._debug_text.configure(yscrollcommand=dbsb.set)
        dbsb.pack(side='right', fill='y')
        ttk.Button(debug_bar, text='Clear', style='Flat.TButton',
                   command=self._clear_debug).pack(side='left', padx=(8, 0))

        # ── Bottom bar ────────────────────────────────────────────────
        ttk.Separator(self, orient='horizontal').pack(fill='x')
        bottom = ttk.Frame(self, padding=(12, 6))
        bottom.pack(side='bottom', fill='x')

        self._start_btn = ttk.Button(bottom, text='▶  Start Reading',
                                     style='Accent.TButton',
                                     command=self._on_start_btn)
        self._start_btn.pack(side='left', padx=(0, 6))

        self._stop_btn = ttk.Button(bottom, text='■  Stop',
                                    style='Stop.TButton',
                                    command=self._on_stop_btn,
                                    state='disabled')
        self._stop_btn.pack(side='left', padx=(0, 16))

        ttk.Button(bottom, text='Clear Log', style='Flat.TButton',
                   command=self._clear_log).pack(side='left')

        ttk.Button(bottom, text='Export CSV', style='Flat.TButton',
                   command=self._export_csv).pack(side='left', padx=(4, 0))

        self._count_var = tk.StringVar(value='0 reads')
        ttk.Label(bottom, textvariable=self._count_var,
                  foreground=MUTED).pack(side='right', padx=(8, 0))
        ttk.Button(bottom, text='Close', style='Flat.TButton',
                   command=self._on_close).pack(side='right', padx=(4, 0))

    def _build_settings(self, parent):
        # ── Inject / Output ──────────────────────────────────────────
        grp = ttk.LabelFrame(parent, text='Basic Output Formatting', padding=(8, 6))
        grp.pack(fill='x', pady=(0, 8))

        ttk.Label(grp, text='Append after tag:').grid(row=0, column=0,
                                                       columnspan=3, sticky='w')
        self._append_var = tk.StringVar(value=self.cfg.append_key)
        for col, choice in enumerate(['Nothing', 'Tab', 'Enter']):
            ttk.Radiobutton(grp, text=choice, value=choice,
                            variable=self._append_var).grid(
                row=1, column=col, sticky='w', padx=(0, 4))

        ttk.Label(grp, text='Prefix:').grid(row=2, column=0, sticky='w',
                                             pady=(6, 0), columnspan=3)
        self._prefix_var = tk.StringVar(value=self.cfg.prefix)
        self._prefix_chk_var = tk.BooleanVar(value=bool(self.cfg.prefix))
        ttk.Checkbutton(grp, text='Add Prefix  (ASCII characters)',
                        variable=self._prefix_chk_var).grid(
            row=3, column=0, columnspan=3, sticky='w')
        self._prefix_entry = ttk.Entry(grp, textvariable=self._prefix_var,
                                       width=22, font=MONO_FONT)
        self._prefix_entry.grid(row=4, column=0, columnspan=3, sticky='ew',
                                 pady=(2, 4))

        ttk.Label(grp, text='Suffix:').grid(row=5, column=0, sticky='w',
                                             columnspan=3)
        self._suffix_var = tk.StringVar(value=self.cfg.suffix)
        self._suffix_chk_var = tk.BooleanVar(value=bool(self.cfg.suffix))
        ttk.Checkbutton(grp, text='Add Suffix  (ASCII characters)',
                        variable=self._suffix_chk_var).grid(
            row=6, column=0, columnspan=3, sticky='w')
        self._suffix_entry = ttk.Entry(grp, textvariable=self._suffix_var,
                                       width=22, font=MONO_FONT)
        self._suffix_entry.grid(row=7, column=0, columnspan=3, sticky='ew',
                                 pady=(2, 4))
        grp.columnconfigure(0, weight=1)

        # ── Behaviour ────────────────────────────────────────────────
        grp2 = ttk.LabelFrame(parent, text='Behaviour', padding=(8, 6))
        grp2.pack(fill='x', pady=(0, 8))

        self._combine_var = tk.BooleanVar(value=self.cfg.combine_reads)
        ttk.Checkbutton(grp2, text='Combine Multiple Reads',
                        variable=self._combine_var).pack(anchor='w')

        timeout_row = ttk.Frame(grp2)
        timeout_row.pack(anchor='w', pady=(2, 4))
        ttk.Label(timeout_row, text='  Dedupe window (ms):').pack(side='left')
        self._timeout_var = tk.StringVar(value=str(self.cfg.combine_timeout_ms))
        ttk.Spinbox(timeout_row, from_=100, to=10000, width=6,
                    textvariable=self._timeout_var,
                    font=UI_FONT).pack(side='left', padx=(4, 0))

        self._spaces_var = tk.BooleanVar(value=self.cfg.display_spaces)
        ttk.Checkbutton(grp2, text='Display Spaces (AA BB CC…)',
                        variable=self._spaces_var).pack(anchor='w')

        self._ts_var = tk.BooleanVar(value=self.cfg.add_timestamp)
        ttk.Checkbutton(grp2, text='Add Timestamp',
                        variable=self._ts_var).pack(anchor='w')

        delim_row = ttk.Frame(grp2)
        delim_row.pack(anchor='w', pady=(2, 0))
        ttk.Label(delim_row, text='  Delimiter: ').pack(side='left')
        self._delim_var = tk.StringVar(value=self.cfg.delimiter)
        ttk.Entry(delim_row, textvariable=self._delim_var,
                  width=6, font=MONO_FONT).pack(side='left')

        # ── Output Format ─────────────────────────────────────────────
        grp3 = ttk.LabelFrame(parent, text='Output Format', padding=(8, 6))
        grp3.pack(fill='x', pady=(0, 8))

        self._output_format_var = tk.StringVar(value=self.cfg.output_format)
        for value, label in [
            ('HEX',        'HEX  — raw EPC hex string'),
            ('GTIN14',     'GTIN-14  — 14-digit GTIN'),
            ('EAN13',      'EAN-13  — 13-digit'),
            ('UPCA',       'UPC-A  — 12-digit (falls back to EAN-13)'),
            ('UPCA_EAN13', 'UPC-A (as EAN-13)  — 13-digit'),
            ('GTIN_SN',    'GTIN-14 + Serial  — GTIN-14 + delimiter + serial'),
        ]:
            ttk.Radiobutton(grp3, text=label, value=value,
                            variable=self._output_format_var).pack(anchor='w')

        self._filter_var = tk.BooleanVar(value=self.cfg.include_filter)
        ttk.Checkbutton(grp3, text='Include Filter digit  (HEX mode only)',
                        variable=self._filter_var).pack(anchor='w', pady=(4, 0))

        # ── Logging ──────────────────────────────────────────────────
        grp4 = ttk.LabelFrame(parent, text='Logging', padding=(8, 6))
        grp4.pack(fill='x', pady=(0, 8))

        self._log_var = tk.BooleanVar(value=self.cfg.log_enabled)
        ttk.Checkbutton(grp4, text='Log reads to file (TagLog.csv)',
                        variable=self._log_var).pack(anchor='w')

        # ── Active buzzer level indicator ─────────────────────────────
        buzzer_row = ttk.Frame(parent)
        buzzer_row.pack(fill='x', pady=(0, 4))
        ttk.Label(buzzer_row, text='Buzzer:', foreground=MUTED).pack(side='left')
        self._buzzer_lbl = ttk.Label(buzzer_row, textvariable=self._buzzer_var,
                                     font=UI_FONT_B, foreground=MUTED)
        self._buzzer_lbl.pack(side='left', padx=(4, 0))
        self._update_buzzer_indicator()

        # ── Save settings button ──────────────────────────────────────
        ttk.Button(parent, text='Save Settings', style='Flat.TButton',
                   command=self._save_settings).pack(fill='x', pady=(4, 0))

    def _build_log(self, parent):
        # ── Last-read decode card ─────────────────────────────────────
        decode_frame = ttk.LabelFrame(parent, text='Last Read — EPC Decode',
                                       padding=(8, 6))
        decode_frame.pack(fill='x', pady=(0, 8))

        self._decode_vars = {}
        self._lock_status_lbl = None   # kept for foreground updates
        labels = [
            ('Injected', 'injected'),
            ('EPC',      'epc'),
            ('TID',      'tid'),
            ('Scheme',   'scheme'),
            ('GTIN-14',  'gtin14'),
            ('Serial',   'serial'),
            ('EPC URI',  'epc_uri'),
            ('RSSI',     'rssi'),
            ('Antenna',  'antenna'),
            ('Lock',     'lock_status'),
        ]
        for row, (label, key) in enumerate(labels):
            ttk.Label(decode_frame, text=f'{label}:', width=10,
                      anchor='e', foreground=MUTED).grid(
                row=row, column=0, sticky='e', pady=1)
            var = tk.StringVar(value='—')
            self._decode_vars[key] = var
            lbl = ttk.Label(decode_frame, textvariable=var,
                            font=MONO_FONT, foreground=ACCENT)
            lbl.grid(row=row, column=1, sticky='w', padx=(6, 0), pady=1)
            if key == 'lock_status':
                self._lock_status_lbl = lbl
            if key == 'gtin14':
                self._gtin_warn_lbl = ttk.Label(decode_frame, text='',
                                                 foreground=RED,
                                                 font=MONO_FONT)
                self._gtin_warn_lbl.grid(row=row, column=2, sticky='w', padx=(6, 0))
        decode_frame.columnconfigure(1, weight=1)

        # ── Live log table ────────────────────────────────────────────
        log_frame = ttk.LabelFrame(parent, text='Read Log', padding=(4, 4))
        log_frame.pack(fill='both', expand=True)

        cols = ('time', 'epc', 'tid', 'scheme', 'gtin14', 'rssi', 'ant')
        self._tree = ttk.Treeview(log_frame, columns=cols,
                                   show='headings', height=12)

        col_defs = [
            ('time',   'Time',      120, 'center'),
            ('epc',    'EPC (hex)', 200, 'w'),
            ('tid',    'TID',       160, 'w'),
            ('scheme', 'Scheme',     80, 'center'),
            ('gtin14', 'GTIN-14',   110, 'center'),
            ('rssi',   'RSSI',       70, 'center'),
            ('ant',    'Ant',        40, 'center'),
        ]
        for col, heading, width, anchor in col_defs:
            self._tree.heading(col, text=heading)
            self._tree.column(col, width=width, anchor=anchor, stretch=False)
        self._tree.column('epc', stretch=True)

        vsb = ttk.Scrollbar(log_frame, orient='vertical',
                             command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

    def _update_buzzer_indicator(self):
        """Refresh the Buzzer indicator label to match cfg.buzzer_level."""
        level = self.cfg.buzzer_level
        self._buzzer_var.set(level)
        color = {
            'OFF':  MUTED,
            'LOW':  YELLOW,
            'HIGH': GREEN,
        }.get(level, MUTED)
        self._buzzer_lbl.configure(foreground=color)

    def _apply_cfg_to_ui(self):
        self._port_var.set(self.cfg.port)
        self._power_var.set(str(self.cfg.power_dbm))
        # Restore the user's chosen startup default preset (applies its settings
        # to cfg + UI; hardware calls are no-ops until reader connects).
        default = self.cfg.default_preset or 'Standard'
        self._preset_var.set(default)
        self._apply_preset(default)

    def _read_cfg_from_ui(self):
        self.cfg.port               = self._port_var.get()
        self.cfg.append_key         = self._append_var.get()
        self.cfg.prefix             = self._prefix_var.get() if self._prefix_chk_var.get() else ''
        self.cfg.suffix             = self._suffix_var.get() if self._suffix_chk_var.get() else ''
        self.cfg.display_spaces     = self._spaces_var.get()
        self.cfg.combine_reads      = self._combine_var.get()
        self.cfg.add_timestamp      = self._ts_var.get()
        self.cfg.delimiter          = self._delim_var.get()
        self.cfg.output_format      = self._output_format_var.get()
        self.cfg.include_filter     = self._filter_var.get()
        self.cfg.log_enabled        = self._log_var.get()
        self.cfg.power_dbm          = int(self._power_var.get() or 20)
        try:
            self.cfg.combine_timeout_ms = int(self._timeout_var.get())
        except ValueError:
            self.cfg.combine_timeout_ms = 1000

    def _save_settings(self):
        self._read_cfg_from_ui()
        self.cfg.save()
        self._set_status('Settings saved', GREEN)

    def _check_dll(self):
        if not pythonnet_present():
            self._connect_btn.config(state='disabled')
            self._set_status('pythonnet missing — see message', RED)
            messagebox.showerror(
                'pythonnet not installed',
                'pythonnet is not installed — run\n\n'
                '    pip install pythonnet\n\n'
                'then restart the app.\n\n'
                'The Connect button will remain disabled until pythonnet is present.'
            )
            return
        if not dll_present():
            self._dll_was_missing = True
            self._connect_btn.config(state='disabled')
            self._set_status('AsReaderP3xU.dll missing — see message', RED)
            messagebox.showerror(
                'DLL not found',
                'AsReaderP3xU.dll not found — copy it from the SDK zip into '
                'this folder.\n\nThe Connect button will remain disabled until '
                'the file is present.\n\n'
                'If the DLL is present but still fails, right-click it → '
                'Properties → Unblock.'
            )

    def _refresh_ports(self):
        ports = E310Reader.list_ports()
        self._port_cb['values'] = ports
        if ports and not self._port_var.get():
            self._port_var.set(ports[0])

    def _periodic_port_refresh(self):
        if not self._reader or not self._reader.is_connected():
            self._refresh_ports()
        if self._dll_was_missing and dll_present():
            self._dll_was_missing = False
            self._connect_btn.config(state='normal')
            self._set_status('Ready')
        self.after(3000, self._periodic_port_refresh)

    def _toggle_connect(self):
        if self._reader and self._reader.is_connected():
            self._do_disconnect()
        else:
            self._do_connect()

    def _toggle_debug(self):
        if self._debug_visible.get():
            self._debug_frame.pack(side='bottom', fill='x')
        else:
            self._debug_frame.pack_forget()

    def _append_raw(self, data: bytes):
        hex_str = ' '.join(f'{b:02X}' for b in data)
        self.after(0, self._insert_debug, hex_str)

    def _insert_debug(self, hex_str: str):
        self._debug_text.configure(state='normal')
        self._debug_text.insert('end', hex_str + '\n')
        content = self._debug_text.get('1.0', 'end')
        if len(content) > 4000:
            self._debug_text.delete('1.0', f'1.0 + {len(content)-2000}c')
        self._debug_text.see('end')
        self._debug_text.configure(state='disabled')

    def _clear_debug(self):
        self._debug_text.configure(state='normal')
        self._debug_text.delete('1.0', 'end')
        self._debug_text.configure(state='disabled')

    def _do_connect(self):
        port = self._port_var.get()
        if not port:
            messagebox.showwarning('No port', 'Select a COM port first.')
            return
        self._read_cfg_from_ui()
        # Disable button immediately so the user can't double-click
        self._connect_btn.config(state='disabled', text='Connecting…')
        self._set_status('Connecting…')

        reader = E310Reader(
            on_tag=self._on_tag,
            on_status=self._on_status_cb,
            on_raw=self._append_raw,
            on_disconnect=self._on_hardware_disconnect,
            on_stopped=self._on_hw_stopped,
            on_trigger=self._on_hw_trigger,
        )

        def _worker():
            ok = reader.connect(port)
            self.after(0, lambda: self._on_connect_done(reader, ok))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_connect_done(self, reader: E310Reader, ok: bool):
        if ok:
            self._reader = reader
            self._trigger_armed = False   # always start each connection disarmed
            self._connect_btn.config(state='normal', text='Disconnect')
            self._status_canvas.itemconfig(self._status_dot, fill=GREEN)
            if self._trigger_mode:
                self._start_btn.config(state='normal', text='▶  Arm Trigger')
                self._stop_btn.config(state='disabled', text='■  Disarm')
            else:
                self._start_btn.config(state='normal')
            # Force an immediate repaint so the green dot and button state are
            # visible before the blocking DLL calls below run.
            self.update_idletasks()
            # Apply advanced settings to hardware right after connect
            self._apply_advanced_to_reader()
            # Fetch device info for About / Device tab
            self._device_info = self._reader.get_device_info()
        else:
            self._reader = None
            self._connect_btn.config(state='normal', text='Connect')

    def _do_disconnect(self):
        if self._reader:
            self._stop_reading()
            self._reader.disconnect()
            self._reader = None
        self._connect_btn.config(text='Connect')
        self._status_canvas.itemconfig(self._status_dot, fill=RED)
        self._start_btn.config(state='disabled')
        self._stop_btn.config(state='disabled')
        self._device_info = {}

    def _on_hardware_disconnect(self):
        self.after(0, self._apply_hardware_disconnect_ui)

    def _on_hw_stopped(self):
        """Called from a DLL callback thread when the hardware hits its mtnu limit.

        Marshals the UI update onto the main tkinter thread.
        In trigger mode, stay armed after auto-stop rather than going to Paused.
        """
        if self._trigger_mode:
            self.after(0, self._re_arm_after_read)
        else:
            self.after(0, self._stop_reading)

    def _on_tid_read(self, epc: str, tid: str):
        """Receive ReadMemory TID result on the main thread.

        Updates the decode card and retroactively fills the TID cell in the
        most recent Read Log row for this EPC (the row was written before the
        QC worker returned, so its TID column starts empty).
        """
        self._decode_vars['tid'].set(tid if tid else '— (not available)')
        if not tid:
            return
        # Find the most recent log row for this EPC that still has no TID.
        for row in self._log_rows:
            if row['epc'] == epc and not row['tid']:
                row['tid'] = tid
                try:
                    self._tree.set(row['iid'], 'tid', tid)
                except Exception:
                    pass  # row may have been pruned; decode panel already updated
                break

    def _on_lock_status(self, epc: str, status: str):
        """Receive CheckTagStatus result on the main thread and update the UI."""
        colour_map = {
            'PermaLock': GREEN,
            'Lock':      YELLOW,
            'UnLock':    RED,
            'Unknown':   MUTED,
            'Error':     MUTED,
        }
        label_map = {
            'PermaLock': '🔒 Permalocked',
            'Lock':      '🔓 Locked',
            'UnLock':    '⚠  Unlocked',
            'Unknown':   'Unknown',
            'Error':     'Tag not in read range — rescan for lock status',
        }
        if self._trigger_mode == 'continuous':
            display = 'Lock status not available in Continuous mode — switch to 1X Trigger for this feature'
            colour   = MUTED
        else:
            display = label_map.get(status, status)
            colour  = colour_map.get(status, MUTED)
        self._decode_vars['lock_status'].set(display)
        if self._lock_status_lbl:
            self._lock_status_lbl.configure(foreground=colour)
        if status == 'UnLock':
            # Restore from tray so the operator sees the alert
            if self._window_hidden:
                self._tray_show()
            self._tray_update('alert', f'{APP_TITLE} — ⚠ Unlocked Tag')
            try:
                winsound.PlaySound(_UNLOCK_CHIME_WAV,
                                   winsound.SND_MEMORY | winsound.SND_ASYNC)
            except Exception:
                self.bell()
            self._set_status('⚠  Tag is Unlocked', RED)

    def _apply_hardware_disconnect_ui(self):
        self._reader = None
        self._device_info = {}
        self._qc_active = False   # clear any in-flight QC so SCAN works after reconnect
        self._connect_btn.config(text='Connect')
        self._status_canvas.itemconfig(self._status_dot, fill=RED)
        self._start_btn.config(state='disabled')
        self._stop_btn.config(state='disabled')

    def _start_reading(self):
        if not self._reader or not self._reader.is_connected():
            return
        self._read_cfg_from_ui()
        self._last_epcs.clear()
        self._clear_card()   # erase stale data from the previous read session
        # max_tags=1: hardware stops itself after the first tag, eliminating the
        # ~50 ms race window that causes a second inventory cycle (and beep).
        max_tags = 1 if self.cfg.stop_after_first else 0
        fv = self.cfg.select_filter_values if self.cfg.select_mask_enabled else []
        self._reader.start_inventory(max_tags=max_tags, use_tid=self.cfg.read_tid,
                                     filter_values=fv)
        self._start_btn.config(state='disabled')
        self._stop_btn.config(state='normal', text='■  Stop')
        self._status_canvas.itemconfig(self._status_dot, fill=GREEN_ACTIVE)
        self._status_var.set('Reading…')
        self._tray_update('reading', f'{APP_TITLE} — Reading')

    def _stop_reading(self):
        if self._reader:
            self._reader.stop_inventory()
        self._start_btn.config(state='normal')
        self._stop_btn.config(state='disabled', text='■  Stopped')
        # Show Paused if still connected; disconnect path overwrites this with RED
        if self._reader and self._reader.is_connected():
            self._status_canvas.itemconfig(self._status_dot, fill=YELLOW)
            self._status_var.set('Paused…')
        self._tray_update('idle', f'{APP_TITLE} — Idle')

    def _on_start_btn(self):
        """Start button — arms trigger in trigger-mode presets, starts reading otherwise."""
        if self._trigger_mode:
            self._arm_trigger()
        else:
            self._start_reading()

    def _on_stop_btn(self):
        """Stop button — disarms trigger in trigger-mode presets, stops reading otherwise."""
        if self._trigger_mode:
            self._disarm_trigger()
        else:
            self._stop_reading()

    def _arm_trigger(self):
        """Enter armed state: reader sits idle, SCAN button will fire inventory."""
        self._trigger_armed = True
        self._start_btn.config(state='disabled')
        self._stop_btn.config(text='■  Disarm', state='normal')
        self._status_canvas.itemconfig(self._status_dot, fill=YELLOW)
        self._status_var.set('Armed — press SCAN to read')
        self._tray_update('idle', f'{APP_TITLE} — Armed')

    def _disarm_trigger(self):
        """Leave armed state: stop any active inventory, ignore future SCAN presses."""
        self._trigger_armed = False
        if self._reader:
            self._reader.stop_inventory()
        self._start_btn.config(text='▶  Arm Trigger', state='normal')
        self._stop_btn.config(text='■  Disarm', state='disabled')
        self._status_canvas.itemconfig(self._status_dot, fill=YELLOW)
        self._status_var.set('Disarmed')
        self._tray_update('idle', f'{APP_TITLE} — Idle')

    def _re_arm_after_read(self):
        """After a hardware auto-stop in trigger mode, stay armed for the next press.

        Called instead of _stop_reading so the Stop/Disarm button stays enabled
        and the status reflects the ready-to-scan state.
        """
        if self._reader:
            self._reader.stop_inventory()   # belt-and-suspenders
        self._status_canvas.itemconfig(self._status_dot, fill=YELLOW)
        self._status_var.set('Armed — press SCAN to read')
        self._tray_update('idle', f'{APP_TITLE} — Armed')
        # Start btn stays disabled (still armed); Stop btn stays as '■  Disarm'.

    def _on_qc_done(self):
        """Called on the tkinter thread when the QC worker finishes. Clears the
        busy flag so the SCAN button can trigger the next read."""
        self._qc_active = False

    def _on_hw_trigger(self, trigger: int) -> None:
        """Called from reader._cb_trigger when the hardware SCAN button fires.

        trigger == 1: button pressed.  trigger == 0: button released.
        Only acts on press; dispatches based on active trigger_mode.
        Called from a DLL callback thread — marshals to tkinter via after().
        """
        if trigger != 1 or not self._trigger_armed:
            return
        if not self._reader or not self._reader.is_connected():
            return
        if self._qc_active:
            return   # previous QC (ReadMemory / CheckTagStatus) still in flight
        if self._trigger_mode == 'oneshot':
            if not self._reader.reading:
                self.after(0, self._start_reading)
        elif self._trigger_mode == 'continuous':
            self.after(0, self._toggle_trigger_reading)

    def _toggle_trigger_reading(self):
        """Toggle inventory on/off for Trigger-Continuous mode."""
        if self._reader and self._reader.reading:
            self._reader.stop_inventory()
            self._status_canvas.itemconfig(self._status_dot, fill=YELLOW)
            self._status_var.set('Armed — press SCAN to read')
            self._tray_update('idle', f'{APP_TITLE} — Armed')
        else:
            self._start_reading()

    def _clear_card(self):
        """Reset every field in the Last Read card to '—' (called on each Start)."""
        for key, var in self._decode_vars.items():
            var.set('—')
        if self._lock_status_lbl:
            self._lock_status_lbl.configure(foreground=ACCENT)
        if hasattr(self, '_gtin_warn_lbl'):
            self._gtin_warn_lbl.config(text='')

    def _set_power(self):
        if not self._reader or not self._reader.is_connected():
            return
        try:
            dbm = int(self._power_var.get())
        except ValueError:
            return
        self._reader.set_power(antenna=1, read_dbm=dbm)
        self.cfg.power_dbm = dbm

    def _on_tag(self, tag: TagRead):
        self._read_cfg_from_ui()

        # Deduplication
        if self.cfg.combine_reads:
            now = time.time()
            last = self._last_epcs.get(tag.raw_epc, 0)
            if now - last < self.cfg.combine_timeout_ms / 1000.0:
                return
            self._last_epcs[tag.raw_epc] = now

        decoded = decode_epc(tag.raw_epc)
        inject_str, fallback_note = format_epc_for_inject(decoded, self.cfg)

        if self.cfg.add_timestamp:
            ts_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            inject_str = inject_str + self.cfg.delimiter + ts_str

        threading.Thread(
            target=inj.inject,
            args=(inject_str, self.cfg.append_key),
            daemon=True
        ).start()

        self.after(0, self._update_ui_for_tag, tag, decoded, inject_str, fallback_note)

        # QC worker — TID read then lock check, sequentially in one thread so
        # only one DLL command is in flight at a time.  Tag must still be in
        # the RF field; safe for static/near-static QC use.
        if (self.cfg.read_tid or self.cfg.check_lock) and self._reader and not self._qc_active:
            epc         = tag.raw_epc
            reader      = self._reader
            # If SetInventoryType(PC_EPC_TID) worked, td.tid is already populated
            # in the TagRead and no ReadMemory fallback is needed.
            tid_from_inv = tag.tid   # non-empty when PC_EPC_TID mode is active
            do_tid      = self.cfg.read_tid and not tid_from_inv
            do_lock     = self.cfg.check_lock
            self._qc_active = True   # block SCAN button until QC finishes
            def _qc_worker(e=epc, r=reader):
                # ── TID source invariant ──────────────────────────────────────
                # There are exactly two paths by which a TID ends up in the log:
                #
                # Path A — PC_EPC_TID inventory mode (preferred, no extra RF round-trip):
                #   SetInventoryType(PC_EPC_TID) makes the firmware embed TID words
                #   directly in the inventory response.  tag.tid is non-empty, so
                #   the log row is seeded at insert time (see _update_ui_for_tag),
                #   tid_from_inv is truthy, do_tid is False, and ReadMemory is
                #   skipped entirely.
                #
                # Path B — ReadMemory fallback (EPC-only mode, or silent firmware
                #   fall-back to EPC-only when PC_EPC_TID is nominally active):
                #   tag.tid is empty, do_tid is True (requires cfg.read_tid), so
                #   r.read_tid() is called below.  _on_tid_read() then fills the
                #   TID column in both the decode panel and the log row.
                #
                # If tag.tid is empty AND cfg.read_tid is False the row stays blank;
                # that is the expected outcome when TID reading is disabled.
                #
                # ── ReadMemory timing ─────────────────────────────────────────
                # ReadMemory requires inventory to be idle.  Wait up to 600 ms
                # for the hardware auto-stop (_cb_complete → _stop_reading) to
                # propagate before issuing any DLL commands.
                try:
                    _deadline = time.time() + 0.6
                    while r.reading and time.time() < _deadline:
                        time.sleep(0.05)
                    if do_tid:
                        tid = r.read_tid(e)
                        self.after(0, self._on_tid_read, e, tid)
                    if do_lock:
                        status = r.check_tag_status(e)
                        self.after(0, self._on_lock_status, e, status)
                finally:
                    self.after(0, self._on_qc_done)   # always clears _qc_active
            threading.Thread(target=_qc_worker, daemon=True).start()

        # Stop-after-first safety fallback: if the hardware auto-stop callback
        # (_on_hw_stopped → _stop_reading) fires first this is a no-op;
        # if for any reason the DLL doesn't fire _cb_complete this cleans up.
        if self.cfg.stop_after_first:
            self.after(50, self._re_arm_after_read if self._trigger_mode else self._stop_reading)

    def _update_ui_for_tag(self, tag: TagRead, decoded: dict,
                           inject_str: str = '', fallback_note: str | None = None):
        ts   = datetime.datetime.fromtimestamp(tag.timestamp).strftime('%H:%M:%S.%f')[:-3]
        rssi = f'{tag.rssi_dbm:.1f} dBm'

        self._decode_vars['injected'].set(inject_str or '—')
        self._decode_vars['epc'].set(tag.raw_epc)
        # TID: show immediately if PC_EPC_TID inventory mode populated it;
        # otherwise show 'Reading…' while the ReadMemory fallback runs.
        if tag.tid:
            self._decode_vars['tid'].set(tag.tid)
        elif self.cfg.read_tid:
            self._decode_vars['tid'].set('Reading…')
        else:
            self._decode_vars['tid'].set('—')
        self._decode_vars['scheme'].set(decoded.get('scheme', '—'))
        gtin14 = decoded.get('gtin14')
        self._decode_vars['gtin14'].set(gtin14 or '—')
        # Check-digit / decode-failure warning
        self._gtin_warn_lbl.config(text='' if gtin14_check_ok(gtin14) else '⚠')
        self._decode_vars['serial'].set(decoded.get('serial') or '—')
        self._decode_vars['epc_uri'].set(decoded.get('epc_uri') or '—')
        self._decode_vars['rssi'].set(rssi)
        self._decode_vars['antenna'].set(str(tag.antenna))
        # Reset lock status to pending while CheckTagStatus runs in background
        self._decode_vars['lock_status'].set('Checking…')
        if self._lock_status_lbl:
            self._lock_status_lbl.configure(foreground=MUTED)

        # UPC-A fallback note (shown amber in the status bar)
        if fallback_note:
            self._set_status(f'⚠  {fallback_note}', YELLOW)

        iid = self._tree.insert('', 0, values=(
            ts,
            tag.raw_epc,
            tag.tid or '',
            decoded.get('scheme', ''),
            decoded.get('gtin14') or '',
            rssi,
            str(tag.antenna),
        ))
        row = {
            'time':   ts,
            'epc':    tag.raw_epc,
            'tid':    tag.tid,
            'scheme': decoded.get('scheme', ''),
            'gtin14': decoded.get('gtin14') or '',
            'rssi':   rssi,
            'ant':    str(tag.antenna),
            'raw_ts': tag.timestamp,
            'iid':    iid,          # Treeview row handle for retroactive TID update
        }
        self._log_rows.insert(0, row)

        if len(self._log_rows) > 10000:
            self._log_rows = self._log_rows[:10000]
            children = self._tree.get_children()
            if len(children) > 10000:
                self._tree.delete(children[-1])

        count = len(self._log_rows)
        self._count_var.set(f'{count} read{"s" if count != 1 else ""}')

        if self.cfg.log_enabled:
            self._append_log_file(row)

    def _log_path(self) -> str:
        base = os.path.dirname(os.path.abspath(
            sys.executable if getattr(sys, 'frozen', False) else __file__))
        return os.path.join(base, 'TagLog.csv')

    def _append_log_file(self, row: dict):
        path = self._log_path()
        write_header = not os.path.exists(path)
        try:
            with open(path, 'a', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                if write_header:
                    w.writerow(['Timestamp', 'EPC', 'TID', 'Scheme', 'GTIN-14', 'RSSI', 'Antenna'])
                w.writerow([
                    datetime.datetime.fromtimestamp(
                        row['raw_ts']).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                    row['epc'], row['tid'], row['scheme'], row['gtin14'],
                    row['rssi'], row['ant'],
                ])
        except Exception as exc:
            logging.error('Log write error: %s', exc)

    def _clear_log(self):
        self._log_rows.clear()
        self._last_epcs.clear()
        for child in self._tree.get_children():
            self._tree.delete(child)
        self._count_var.set('0 reads')

    def _export_csv(self):
        if not self._log_rows:
            messagebox.showinfo('Empty', 'No reads to export.')
            return
        path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')],
            title='Export Read Log',
        )
        if not path:
            return
        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(['Timestamp', 'EPC', 'TID', 'Scheme', 'GTIN-14', 'RSSI', 'Antenna'])
                for row in reversed(self._log_rows):
                    w.writerow([
                        datetime.datetime.fromtimestamp(
                            row['raw_ts']).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                        row['epc'], row['tid'], row['scheme'], row['gtin14'],
                        row['rssi'], row['ant'],
                    ])
            messagebox.showinfo('Exported', f'Saved to:\n{path}')
        except Exception as exc:
            messagebox.showerror('Export failed', str(exc))

    def _on_status_cb(self, msg: str):
        self.after(0, self._set_status, msg)

    def _set_status(self, msg: str, color: str = MUTED):
        self._status_var.set(msg)

    def _on_close(self):
        """X button: always quit.  Minimize-button → tray is the hide gesture."""
        self._quit_app()

    def _quit_app(self):
        """Fully exit: save config, disconnect reader, destroy window."""
        # Re-entrancy guard — Close button, X button, and tray Exit can all
        # fire close-to-simultaneously; only the first call should proceed.
        if getattr(self, '_quitting', False):
            return
        self._quitting = True
        try:
            self._read_cfg_from_ui()
            self.cfg.save()
            # Disconnect in a daemon thread with a hard timeout.
            # StopInventory()/DisConnect() can block inside the CLR indefinitely;
            # if they hang, the finally clause still delivers os._exit(0).
            if self._reader:
                t = threading.Thread(target=self._reader.disconnect, daemon=True)
                t.start()
                t.join(timeout=1.5)
            # Stop pystray: set visible=False synchronously (NIM_DELETE) so the
            # tray icon disappears immediately, then spin off stop() as a daemon
            # thread so its Win32 message-loop join cannot block os._exit(0).
            if self._tray_icon is not None:
                try:
                    self._tray_icon.visible = False
                    threading.Thread(target=self._tray_icon.stop,
                                     daemon=True, name='tray-stop').start()
                except Exception:
                    pass
                self._tray_icon = None
            self.destroy()
        finally:
            # Hard-exit the process.  pythonnet's CLR keeps non-daemon .NET
            # threads alive; os._exit(0) kills them immediately.  The finally
            # clause guarantees this runs even if destroy() or any step above
            # raises an exception.
            import os as _os
            _os._exit(0)

    def _on_unmap(self, event):
        """Minimize button → hide to tray (fires only on the root window).

        Rapid minimize/restore/minimize sequences can queue multiple stale
        after() callbacks.  We prevent this by:
          1. Ignoring the event if the window is already hidden (withdraw()
             itself can re-fire Unmap on some Windows builds).
          2. Cancelling any already-pending hide before scheduling a new one
             so at most one _tray_hide call is ever in flight at a time.
        The stored ID lets _tray_show() cancel the callback if the user
        restores the window before the 50 ms delay expires.
        """
        if event.widget is self and self._tray_icon is not None:
            if self._window_hidden:
                return
            # Cancel any previously queued hide before scheduling a fresh one.
            if self._tray_hide_id is not None:
                self.after_cancel(self._tray_hide_id)
            self._tray_hide_id = self.after(50, self._tray_hide)

    @staticmethod
    def _make_tray_image(state: str) -> 'Image | None':
        """Draw a 64×64 RFID icon in the state colour (idle/reading/alert)."""
        if not _TRAY_AVAILABLE:
            return None
        color_map = {
            'idle':    '#6b7280',
            'reading': '#1a56db',
            'alert':   '#e53e3e',
        }
        fill = color_map.get(state, '#6b7280')
        sz = 64
        img = _PILImage.new('RGBA', (sz, sz), (0, 0, 0, 0))
        d = _PILDraw.Draw(img)
        d.ellipse([2, 2, sz - 2, sz - 2], fill=fill)
        cx, cy = sz // 2, sz // 2
        # Three RFID signal arcs (right-facing)
        for r in (9, 15, 21):
            d.arc([cx - r, cy - r, cx + r, cy + r],
                  start=-55, end=55, fill='white', width=3)
        # Tag dot
        d.ellipse([cx - 4, cy - 4, cx + 4, cy + 4], fill='white')
        return img

    def _setup_tray(self):
        """Create the system tray icon in a daemon thread (no-op if unavailable)."""
        if not _TRAY_AVAILABLE:
            return

        def _is_reading():
            return bool(self._reader and self._reader.reading)

        def _is_connected():
            return bool(self._reader and self._reader.is_connected())

        def _on_show(icon, item):
            self.after(0, self._tray_show)

        def _on_toggle_read(icon, item):
            if _is_reading():
                self.after(0, self._on_stop_btn)
            else:
                self.after(0, self._on_start_btn)

        def _on_quit(icon, item):
            self.after(0, self._quit_app)

        menu = pystray.Menu(
            pystray.MenuItem('Show Window', _on_show, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                lambda item: (
                    '■  Stop Reading' if _is_reading() else
                    '■  Disarm'       if self._trigger_armed else
                    '▶  Arm Trigger'  if self._trigger_mode else
                    '▶  Start Reading'
                ),
                _on_toggle_read,
                enabled=lambda item: _is_connected(),
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('Exit', _on_quit),
        )

        img = self._make_tray_image('idle')
        self._tray_icon = pystray.Icon(
            name='rfid_wedge_pro',
            icon=img,
            title=f'{APP_TITLE} — Idle',
            menu=menu,
        )
        threading.Thread(target=self._tray_icon.run,
                         daemon=True, name='tray').start()

    def _tray_show(self):
        """Restore the main window from tray. Must be called from tkinter thread."""
        # Cancel any pending hide so a minimize → restore within 50 ms does not
        # withdraw the window again after it has been made visible.
        if self._tray_hide_id is not None:
            self.after_cancel(self._tray_hide_id)
            self._tray_hide_id = None
        self._window_hidden = False
        self.deiconify()
        self.lift()
        self.focus_force()

    def _tray_hide(self):
        """Withdraw the main window to tray. Must be called from tkinter thread."""
        self._tray_hide_id = None
        self._window_hidden = True
        self.withdraw()

    def _tray_update(self, state: str, tooltip: str | None = None):
        """Update tray icon colour and tooltip. Must be called from tkinter thread."""
        if self._tray_icon is None:
            return
        self._tray_state = state
        img = self._make_tray_image(state)
        if img is not None:
            self._tray_icon.icon = img
        if tooltip is not None:
            self._tray_icon.title = tooltip

    def _populate_preset_menu(self):
        """(Re-)build the entire Presets menu contents from scratch."""
        m = self._preset_menu
        m.delete(0, 'end')

        # ── Factory presets ──
        for name, data in FACTORY_PRESETS.items():
            m.add_radiobutton(
                label=data['label'],
                variable=self._preset_var,
                value=name,
                command=lambda n=name: self._apply_preset(n),
            )

        m.add_separator()

        # ── Custom presets ──
        for i, key in enumerate(CUSTOM_SLOT_KEYS):
            slot  = self.cfg.custom_presets[i]
            uname = slot.get('name', '').strip()
            label = f'★  {uname}' if uname else f'Custom {i + 1} — (empty)'
            state = 'normal' if uname else 'disabled'
            m.add_radiobutton(
                label=label,
                variable=self._preset_var,
                value=key,
                state=state,
                command=lambda k=key: self._apply_preset(k),
            )

        m.add_separator()
        m.add_command(label='⭐  Save current selection as startup default',
                      command=self._set_default_preset)
        m.add_separator()
        m.add_command(label='Save current settings as…',
                      command=self._open_save_as_dialog)
        m.add_command(label='Rename custom presets…',
                      command=self._open_rename_dialog)
        m.add_separator()
        m.add_command(label='Reset factory presets…',
                      command=self._open_reset_factory_dialog)
        m.add_separator()
        m.add_command(label='Restore factory defaults…',
                      command=self._open_restore_factory_defaults_dialog)

    def _apply_preset(self, key: str):
        """Apply a named preset (factory key or 'Custom 1' / 'Custom 2')."""
        if key in FACTORY_PRESETS:
            p = FACTORY_PRESETS[key]
        elif key in CUSTOM_SLOT_KEYS:
            idx = CUSTOM_SLOT_KEYS.index(key)
            p   = self.cfg.custom_presets[idx]
            if not p.get('name', '').strip():
                return   # slot empty
        else:
            return

        for attr in ('power_dbm', 'buzzer_level', 'combine_reads',
                     'combine_timeout_ms', 'stop_after_first',
                     'read_time_ms', 'idle_time_ms', 'session',
                     'anti_collision', 'anti_collision_start_q',
                     'anti_collision_min_q', 'anti_collision_max_q',
                     'rssi_threshold', 'read_tid', 'check_lock'):
            if attr in p:
                setattr(self.cfg, attr, p[attr])

        # trigger_mode is not a Config field — lives on self.
        # On mode change, stop any active reading/arming and relabel buttons.
        new_trigger_mode = p.get('trigger_mode', '')
        if new_trigger_mode != self._trigger_mode:
            if self._trigger_armed:
                self._disarm_trigger()
            elif self._reader and self._reader.reading:
                self._stop_reading()
            self._trigger_mode  = new_trigger_mode
            self._trigger_armed = False
            self._qc_active     = False   # discard any in-flight QC from the old mode
            self._last_epcs.clear()       # clear dedup table so first scan in new mode always fires
            connected = bool(self._reader and self._reader.is_connected())
            if connected:
                if new_trigger_mode:
                    self._start_btn.config(text='▶  Arm Trigger', state='normal')
                    self._stop_btn.config(text='■  Disarm', state='disabled')
                else:
                    self._start_btn.config(text='▶  Start Reading', state='normal')
                    self._stop_btn.config(text='■  Stopped', state='disabled')
        else:
            self._trigger_mode = new_trigger_mode

        self._power_var.set(str(self.cfg.power_dbm))
        self._preset_var.set(key)
        self._apply_advanced_to_reader()
        self._update_buzzer_indicator()
        self.cfg.save()
        display = (p.get('name') or key) if key in CUSTOM_SLOT_KEYS else key
        self._preset_lbl_var.set(f'[{display}]')
        self._set_status(f'Preset: {display}')

    def _open_save_as_dialog(self):
        """Dialog: capture current settings into a custom preset slot."""
        dlg = tk.Toplevel(self)
        dlg.title('Save Preset')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text='Save current settings as:',
                  font=UI_FONT_B).pack(padx=20, pady=(16, 8))

        slot_var  = tk.IntVar(value=0)
        name_vars = []
        for i, key in enumerate(CUSTOM_SLOT_KEYS):
            slot  = self.cfg.custom_presets[i]
            uname = slot.get('name', '').strip() or f'Custom {i + 1}'
            row   = ttk.Frame(dlg)
            row.pack(fill='x', padx=20, pady=2)
            ttk.Radiobutton(row, text=f'Slot {i + 1}:',
                            variable=slot_var, value=i).pack(side='left')
            nv = tk.StringVar(value=uname)
            name_vars.append(nv)
            ttk.Entry(row, textvariable=nv, width=24).pack(side='left', padx=(6, 0))

        def _do_save():
            idx  = slot_var.get()
            name = name_vars[idx].get().strip()
            if not name:
                messagebox.showwarning('Name required',
                                       'Please enter a name for this preset.',
                                       parent=dlg)
                return
            self._read_cfg_from_ui()
            self.cfg.custom_presets[idx] = dict(
                name=name,
                power_dbm=self.cfg.power_dbm,
                buzzer_level=self.cfg.buzzer_level,
                combine_reads=self.cfg.combine_reads,
                combine_timeout_ms=self.cfg.combine_timeout_ms,
                stop_after_first=self.cfg.stop_after_first,
                read_time_ms=self.cfg.read_time_ms,
                idle_time_ms=self.cfg.idle_time_ms,
                session=self.cfg.session,
                anti_collision=self.cfg.anti_collision,
                anti_collision_start_q=self.cfg.anti_collision_start_q,
                anti_collision_min_q=self.cfg.anti_collision_min_q,
                anti_collision_max_q=self.cfg.anti_collision_max_q,
                rssi_threshold=self.cfg.rssi_threshold,
                read_tid=self.cfg.read_tid,
                check_lock=self.cfg.check_lock,
            )
            self.cfg.save()
            self._preset_var.set(CUSTOM_SLOT_KEYS[idx])
            self._populate_preset_menu()
            dlg.destroy()
            self._set_status(f'Saved preset: {name}')

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill='x', padx=20, pady=(12, 16))
        ttk.Button(btn_row, text='Save', style='Accent.TButton',
                   command=_do_save).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=dlg.destroy).pack(side='right')

        dlg.update_idletasks()
        pw = self.winfo_x() + self.winfo_width() // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    def _open_rename_dialog(self):
        """Dialog: rename the two custom preset slots."""
        dlg = tk.Toplevel(self)
        dlg.title('Rename Custom Presets')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text='Custom preset names:',
                  font=UI_FONT_B).pack(padx=20, pady=(16, 8))

        name_vars = []
        for i in range(2):
            uname = self.cfg.custom_presets[i].get('name', '').strip()
            row   = ttk.Frame(dlg)
            row.pack(fill='x', padx=20, pady=4)
            ttk.Label(row, text=f'Slot {i + 1}:', width=8).pack(side='left')
            nv = tk.StringVar(value=uname)
            name_vars.append(nv)
            ttk.Entry(row, textvariable=nv, width=24).pack(side='left', padx=(4, 0))

        ttk.Label(dlg, text='Leave blank to mark the slot as empty.',
                  foreground=MUTED, font=('Segoe UI', 8)).pack(padx=20, pady=(4, 0))

        def _do_rename():
            for i, nv in enumerate(name_vars):
                self.cfg.custom_presets[i]['name'] = nv.get().strip()
            self.cfg.save()
            self._populate_preset_menu()
            dlg.destroy()

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill='x', padx=20, pady=(12, 16))
        ttk.Button(btn_row, text='OK', style='Accent.TButton',
                   command=_do_rename).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=dlg.destroy).pack(side='right')

        dlg.update_idletasks()
        pw = self.winfo_x() + self.winfo_width() // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    def _set_default_preset(self):
        """Save the currently selected preset as the startup default."""
        key = self._preset_var.get()
        # Resolve a display name for the confirmation message
        if key in FACTORY_PRESETS:
            display = FACTORY_PRESETS[key]['label']
        else:
            try:
                idx = CUSTOM_SLOT_KEYS.index(key)
                display = self.cfg.custom_presets[idx].get('name', '').strip() or key
            except ValueError:
                display = key
        self.cfg.default_preset = key
        self.cfg.save()
        self._set_status(f'Default preset set to "{display}"', GREEN)

    def _open_reset_factory_dialog(self):
        """Dialog: restore selected factory presets to their built-in defaults."""
        dlg = tk.Toplevel(self)
        dlg.title('Reset Factory Presets')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg,
                  text='Select factory presets to restore to built-in defaults.\n'
                       'Tunable adjustments (power, buzzer, timing, QC) will be reset.\n'
                       'User-defined presets will not be affected.',
                  font=UI_FONT, justify='left').pack(padx=20, pady=(16, 8))

        check_vars: dict[str, tk.BooleanVar] = {}
        for name in FACTORY_PRESETS:
            var = tk.BooleanVar(value=False)
            check_vars[name] = var
            ttk.Checkbutton(dlg, text=name, variable=var).pack(
                anchor='w', padx=28, pady=1)

        def _select_all():
            for v in check_vars.values():
                v.set(True)

        def _do_reset():
            restored = []
            for name, var in check_vars.items():
                if not var.get():
                    continue
                p = FACTORY_PRESETS[name]
                if self._preset_var.get() == name:
                    # Active preset — push changes to live cfg and hardware
                    for attr in ('power_dbm', 'buzzer_level', 'combine_reads',
                                 'combine_timeout_ms', 'stop_after_first',
                                 'read_time_ms', 'idle_time_ms', 'session',
                                 'anti_collision', 'anti_collision_start_q',
                                 'anti_collision_min_q', 'anti_collision_max_q'):
                        if attr in p:
                            setattr(self.cfg, attr, p[attr])
                    self._power_var.set(str(self.cfg.power_dbm))
                    self._apply_advanced_to_reader()
                    self._update_buzzer_indicator()
                restored.append(name)
            if restored:
                self.cfg.save()
                self._set_status(f'Reset: {", ".join(restored)}')
            dlg.destroy()

        link_row = ttk.Frame(dlg)
        link_row.pack(fill='x', padx=20, pady=(8, 0))
        ttk.Button(link_row, text='Select all', style='Flat.TButton',
                   command=_select_all).pack(side='left')

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill='x', padx=20, pady=(8, 16))
        ttk.Button(btn_row, text='Reset selected', style='Accent.TButton',
                   command=_do_reset).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=dlg.destroy).pack(side='right')

        dlg.update_idletasks()
        pw = self.winfo_x() + self.winfo_width() // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    def _open_restore_factory_defaults_dialog(self):
        """Confirmation dialog: wipe all settings and restore factory defaults."""
        dlg = tk.Toplevel(self)
        dlg.title('Restore Factory Defaults')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg,
                  text='Restore ALL settings to factory defaults?',
                  font=UI_FONT_B).pack(padx=20, pady=(16, 6))

        ttk.Label(dlg,
                  text=(
                      'The following will be reset to their out-of-box values:\n\n'
                      '  • Output format, delimiter, prefix, suffix\n'
                      '  • Tag filter and COM port selection\n'
                      '  • Both custom preset slots (cleared)\n'
                      '  • Startup default preset → Standard\n'
                      '  • All advanced settings (power, timing, buzzer, QC)\n\n'
                      'The saved config file will be overwritten.\n'
                      'This cannot be undone.'
                  ),
                  font=UI_FONT, justify='left',
                  foreground=MUTED).pack(padx=20, pady=(0, 12))

        def _do_restore():
            self.cfg = Config()
            self.cfg.save()
            # Push every setting back to the UI widgets
            self._append_var.set(self.cfg.append_key)
            self._prefix_var.set(self.cfg.prefix)
            self._prefix_chk_var.set(bool(self.cfg.prefix))
            self._suffix_var.set(self.cfg.suffix)
            self._suffix_chk_var.set(bool(self.cfg.suffix))
            self._combine_var.set(self.cfg.combine_reads)
            self._timeout_var.set(str(self.cfg.combine_timeout_ms))
            self._spaces_var.set(self.cfg.display_spaces)
            self._ts_var.set(self.cfg.add_timestamp)
            self._delim_var.set(self.cfg.delimiter)
            self._output_format_var.set(self.cfg.output_format)
            self._filter_var.set(self.cfg.include_filter)
            self._log_var.set(self.cfg.log_enabled)
            self._port_var.set(self.cfg.port)
            self._power_var.set(str(self.cfg.power_dbm))
            # Reset preset selector and rebuild menu
            self._preset_var.set('Standard')
            self._populate_preset_menu()
            self._apply_preset('Standard')
            self._update_buzzer_indicator()
            dlg.destroy()
            self._set_status('Factory defaults restored', GREEN)

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill='x', padx=20, pady=(0, 16))
        ttk.Button(btn_row, text='Restore', style='Accent.TButton',
                   command=_do_restore).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=dlg.destroy).pack(side='right')

        dlg.update_idletasks()
        pw = self.winfo_x() + self.winfo_width() // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    def _apply_advanced_to_reader(self):
        """Push all advanced config values to the reader hardware (if connected).

        The DLL rejects configuration calls while inventory is running and prints
        "write error" to stdout.  Stop inventory first, apply, then restart.
        """
        r = self._reader
        if not r or not r.is_connected():
            return
        was_reading = r.reading
        if was_reading:
            r.stop_inventory()
            time.sleep(0.15)   # give the DLL a moment to settle before reconfiguring
        r.set_power(antenna=1, read_dbm=self.cfg.power_dbm)
        r.set_read_time(self.cfg.read_time_ms)
        r.set_idle_time(self.cfg.idle_time_ms)
        r.set_session(self.cfg.session)
        r.set_anti_collision(
            self.cfg.anti_collision,
            self.cfg.anti_collision_start_q,
            self.cfg.anti_collision_min_q,
            self.cfg.anti_collision_max_q,
        )
        r.set_rssi_threshold(self.cfg.rssi_threshold)
        level = self.cfg.buzzer_level
        r.set_buzzer(level != 'OFF', level)
        # TARGET_A for stop-after-first; TOGGLE for continuous presets so tags
        # naturally fall out of repeated singulation without Python-side timers.
        r.set_basic_target('TARGET_A' if self.cfg.stop_after_first else 'TOGGLE')
        if was_reading:
            max_tags = 1 if self.cfg.stop_after_first else 0
            fv = self.cfg.select_filter_values if self.cfg.select_mask_enabled else []
            r.start_inventory(max_tags=max_tags, filter_values=fv)

    def _open_advanced(self):
        AdvancedSettingsDialog(self,
                               is_factory=self._preset_var.get() in FACTORY_PRESETS)

    def _open_debug_log(self):
        log_path = os.path.join(
            os.path.dirname(os.path.abspath(
                sys.executable if getattr(sys, 'frozen', False) else __file__)),
            'debug.log',
        )
        if not os.path.exists(log_path):
            messagebox.showinfo('debug.log', 'No debug.log found yet — start the reader to generate one.')
            return
        try:
            os.startfile(log_path)
        except Exception:
            subprocess.Popen(['notepad.exe', log_path])

    def _open_readme(self):
        readme = os.path.join(
            os.path.dirname(os.path.abspath(
                sys.executable if getattr(sys, 'frozen', False) else __file__)),
            'README.md',
        )
        if not os.path.exists(readme):
            messagebox.showinfo('README', 'README.md not found next to the application.')
            return
        try:
            os.startfile(readme)
        except Exception:
            subprocess.Popen(['notepad.exe', readme])

    def _open_about(self):
        AboutDialog(self, self._device_info)
        }
        fill = color_map.get(state, '#6b7280')
        sz = 64
        img = _PILImage.new('RGBA', (sz, sz), (0, 0, 0, 0))
        d = _PILDraw.Draw(img)
        d.ellipse([2, 2, sz - 2, sz - 2], fill=fill)
        cx, cy = sz // 2, sz // 2
        # Three RFID signal arcs (right-facing)
        for r in (9, 15, 21):
            d.arc([cx - r, cy - r, cx + r, cy + r],
                  start=-55, end=55, fill='white', width=3)
        # Tag dot
        d.ellipse([cx - 4, cy - 4, cx + 4, cy + 4], fill='white')
        return img

    def _setup_tray(self):
        """Create the system tray icon in a daemon thread (no-op if unavailable)."""
        if not _TRAY_AVAILABLE:
            return

        def _is_reading():
            return bool(self._reader and self._reader.reading)

        def _is_connected():
            return bool(self._reader and self._reader.is_connected())

        def _on_show(icon, item):
            self.after(0, self._tray_show)

        def _on_toggle_read(icon, item):
            if _is_reading():
                self.after(0, self._on_stop_btn)
            else:
                self.after(0, self._on_start_btn)

        def _on_quit(icon, item):
            self.after(0, self._quit_app)

        menu = pystray.Menu(
            pystray.MenuItem('Show Window', _on_show, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                lambda item: (
                    '■  Stop Reading' if _is_reading() else
                    '■  Disarm'       if self._trigger_armed else
                    '▶  Arm Trigger'  if self._trigger_mode else
                    '▶  Start Reading'
                ),
                _on_toggle_read,
                enabled=lambda item: _is_connected(),
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('Exit', _on_quit),
        )

        img = self._make_tray_image('idle')
        self._tray_icon = pystray.Icon(
            name='rfid_wedge_pro',
            icon=img,
            title=f'{APP_TITLE} — Idle',
            menu=menu,
        )
        threading.Thread(target=self._tray_icon.run,
                         daemon=True, name='tray').start()

    def _tray_show(self):
        """Restore the main window from tray. Must be called from tkinter thread."""
        # Cancel any pending hide so a minimize → restore within 50 ms does not
        # withdraw the window again after it has been made visible.
        if self._tray_hide_id is not None:
            self.after_cancel(self._tray_hide_id)
            self._tray_hide_id = None
        self._window_hidden = False
        self.deiconify()
        self.lift()
        self.focus_force()

    def _tray_hide(self):
        """Withdraw the main window to tray. Must be called from tkinter thread."""
        self._tray_hide_id = None
        self._window_hidden = True
        self.withdraw()

    def _tray_update(self, state: str, tooltip: str | None = None):
        """Update tray icon colour and tooltip. Must be called from tkinter thread."""
        if self._tray_icon is None:
            return
        self._tray_state = state
        img = self._make_tray_image(state)
        if img is not None:
            self._tray_icon.icon = img
        if tooltip is not None:
            self._tray_icon.title = tooltip

    # ── Presets ──────────────────────────────────────────────────────

    def _populate_preset_menu(self):
        """(Re-)build the entire Presets menu contents from scratch."""
        m = self._preset_menu
        m.delete(0, 'end')

        # ── Factory presets ──
        for name, data in FACTORY_PRESETS.items():
            m.add_radiobutton(
                label=data['label'],
                variable=self._preset_var,
                value=name,
                command=lambda n=name: self._apply_preset(n),
            )

        m.add_separator()

        # ── Custom presets ──
        for i, key in enumerate(CUSTOM_SLOT_KEYS):
            slot  = self.cfg.custom_presets[i]
            uname = slot.get('name', '').strip()
            label = f'★  {uname}' if uname else f'Custom {i + 1} — (empty)'
            state = 'normal' if uname else 'disabled'
            m.add_radiobutton(
                label=label,
                variable=self._preset_var,
                value=key,
                state=state,
                command=lambda k=key: self._apply_preset(k),
            )

        m.add_separator()
        m.add_command(label='⭐  Save current selection as startup default',
                      command=self._set_default_preset)
        m.add_separator()
        m.add_command(label='Save current settings as…',
                      command=self._open_save_as_dialog)
        m.add_command(label='Rename custom presets…',
                      command=self._open_rename_dialog)
        m.add_separator()
        m.add_command(label='Reset factory presets…',
                      command=self._open_reset_factory_dialog)
        m.add_separator()
        m.add_command(label='Restore factory defaults…',
                      command=self._open_restore_factory_defaults_dialog)

    def _apply_preset(self, key: str):
        """Apply a named preset (factory key or 'Custom 1' / 'Custom 2')."""
        if key in FACTORY_PRESETS:
            p = FACTORY_PRESETS[key]
        elif key in CUSTOM_SLOT_KEYS:
            idx = CUSTOM_SLOT_KEYS.index(key)
            p   = self.cfg.custom_presets[idx]
            if not p.get('name', '').strip():
                return   # slot empty
        else:
            return

        for attr in ('power_dbm', 'buzzer_level', 'combine_reads',
                     'combine_timeout_ms', 'stop_after_first',
                     'read_time_ms', 'idle_time_ms', 'session',
                     'anti_collision', 'anti_collision_start_q',
                     'anti_collision_min_q', 'anti_collision_max_q',
                     'rssi_threshold', 'read_tid', 'check_lock'):
            if attr in p:
                setattr(self.cfg, attr, p[attr])

        # trigger_mode is not a Config field — lives on self.
        # On mode change, stop any active reading/arming and relabel buttons.
        new_trigger_mode = p.get('trigger_mode', '')
        if new_trigger_mode != self._trigger_mode:
            if self._trigger_armed:
                self._disarm_trigger()
            elif self._reader and self._reader.reading:
                self._stop_reading()
            self._trigger_mode  = new_trigger_mode
            self._trigger_armed = False
            self._qc_active     = False   # discard any in-flight QC from the old mode
            self._last_epcs.clear()       # clear dedup table so first scan in new mode always fires
            connected = bool(self._reader and self._reader.is_connected())
            if connected:
                if new_trigger_mode:
                    self._start_btn.config(text='▶  Arm Trigger', state='normal')
                    self._stop_btn.config(text='■  Disarm', state='disabled')
                else:
                    self._start_btn.config(text='▶  Start Reading', state='normal')
                    self._stop_btn.config(text='■  Stopped', state='disabled')
        else:
            self._trigger_mode = new_trigger_mode

        self._power_var.set(str(self.cfg.power_dbm))
        self._preset_var.set(key)
        self._apply_advanced_to_reader()
        self._update_buzzer_indicator()
        self.cfg.save()
        display = (p.get('name') or key) if key in CUSTOM_SLOT_KEYS else key
        self._preset_lbl_var.set(f'[{display}]')
        self._set_status(f'Preset: {display}')

    def _open_save_as_dialog(self):
        """Dialog: capture current settings into a custom preset slot."""
        dlg = tk.Toplevel(self)
        dlg.title('Save Preset')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text='Save current settings as:',
                  font=UI_FONT_B).pack(padx=20, pady=(16, 8))

        slot_var  = tk.IntVar(value=0)
        name_vars = []
        for i, key in enumerate(CUSTOM_SLOT_KEYS):
            slot  = self.cfg.custom_presets[i]
            uname = slot.get('name', '').strip() or f'Custom {i + 1}'
            row   = ttk.Frame(dlg)
            row.pack(fill='x', padx=20, pady=2)
            ttk.Radiobutton(row, text=f'Slot {i + 1}:',
                            variable=slot_var, value=i).pack(side='left')
            nv = tk.StringVar(value=uname)
            name_vars.append(nv)
            ttk.Entry(row, textvariable=nv, width=24).pack(side='left', padx=(6, 0))

        def _do_save():
            idx  = slot_var.get()
            name = name_vars[idx].get().strip()
            if not name:
                messagebox.showwarning('Name required',
                                       'Please enter a name for this preset.',
                                       parent=dlg)
                return
            self._read_cfg_from_ui()
            self.cfg.custom_presets[idx] = dict(
                name=name,
                power_dbm=self.cfg.power_dbm,
                buzzer_level=self.cfg.buzzer_level,
                combine_reads=self.cfg.combine_reads,
                combine_timeout_ms=self.cfg.combine_timeout_ms,
                stop_after_first=self.cfg.stop_after_first,
                read_time_ms=self.cfg.read_time_ms,
                idle_time_ms=self.cfg.idle_time_ms,
                session=self.cfg.session,
                anti_collision=self.cfg.anti_collision,
                anti_collision_start_q=self.cfg.anti_collision_start_q,
                anti_collision_min_q=self.cfg.anti_collision_min_q,
                anti_collision_max_q=self.cfg.anti_collision_max_q,
                rssi_threshold=self.cfg.rssi_threshold,
                read_tid=self.cfg.read_tid,
                check_lock=self.cfg.check_lock,
            )
            self.cfg.save()
            self._preset_var.set(CUSTOM_SLOT_KEYS[idx])
            self._populate_preset_menu()
            dlg.destroy()
            self._set_status(f'Saved preset: {name}')

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill='x', padx=20, pady=(12, 16))
        ttk.Button(btn_row, text='Save', style='Accent.TButton',
                   command=_do_save).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=dlg.destroy).pack(side='right')

        dlg.update_idletasks()
        pw = self.winfo_x() + self.winfo_width() // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    def _open_rename_dialog(self):
        """Dialog: rename the two custom preset slots."""
        dlg = tk.Toplevel(self)
        dlg.title('Rename Custom Presets')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text='Custom preset names:',
                  font=UI_FONT_B).pack(padx=20, pady=(16, 8))

        name_vars = []
        for i in range(2):
            uname = self.cfg.custom_presets[i].get('name', '').strip()
            row   = ttk.Frame(dlg)
            row.pack(fill='x', padx=20, pady=4)
            ttk.Label(row, text=f'Slot {i + 1}:', width=8).pack(side='left')
            nv = tk.StringVar(value=uname)
            name_vars.append(nv)
            ttk.Entry(row, textvariable=nv, width=24).pack(side='left', padx=(4, 0))

        ttk.Label(dlg, text='Leave blank to mark the slot as empty.',
                  foreground=MUTED, font=('Segoe UI', 8)).pack(padx=20, pady=(4, 0))

        def _do_rename():
            for i, nv in enumerate(name_vars):
                self.cfg.custom_presets[i]['name'] = nv.get().strip()
            self.cfg.save()
            self._populate_preset_menu()
            dlg.destroy()

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill='x', padx=20, pady=(12, 16))
        ttk.Button(btn_row, text='OK', style='Accent.TButton',
                   command=_do_rename).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=dlg.destroy).pack(side='right')

        dlg.update_idletasks()
        pw = self.winfo_x() + self.winfo_width() // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    def _set_default_preset(self):
        """Save the currently selected preset as the startup default."""
        key = self._preset_var.get()
        # Resolve a display name for the confirmation message
        if key in FACTORY_PRESETS:
            display = FACTORY_PRESETS[key]['label']
        else:
            try:
                idx = CUSTOM_SLOT_KEYS.index(key)
                display = self.cfg.custom_presets[idx].get('name', '').strip() or key
            except ValueError:
                display = key
        self.cfg.default_preset = key
        self.cfg.save()
        self._set_status(f'Default preset set to "{display}"', GREEN)

    def _open_reset_factory_dialog(self):
        """Dialog: restore selected factory presets to their built-in defaults."""
        dlg = tk.Toplevel(self)
        dlg.title('Reset Factory Presets')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg,
                  text='Select factory presets to restore to built-in defaults.\n'
                       'Tunable adjustments (power, buzzer, timing, QC) will be reset.\n'
                       'User-defined presets will not be affected.',
                  font=UI_FONT, justify='left').pack(padx=20, pady=(16, 8))

        check_vars: dict[str, tk.BooleanVar] = {}
        for name in FACTORY_PRESETS:
            var = tk.BooleanVar(value=False)
            check_vars[name] = var
            ttk.Checkbutton(dlg, text=name, variable=var).pack(
                anchor='w', padx=28, pady=1)

        def _select_all():
            for v in check_vars.values():
                v.set(True)

        def _do_reset():
            restored = []
            for name, var in check_vars.items():
                if not var.get():
                    continue
                p = FACTORY_PRESETS[name]
                if self._preset_var.get() == name:
                    # Active preset — push changes to live cfg and hardware
                    for attr in ('power_dbm', 'buzzer_level', 'combine_reads',
                                 'combine_timeout_ms', 'stop_after_first',
                                 'read_time_ms', 'idle_time_ms', 'session',
                                 'anti_collision', 'anti_collision_start_q',
                                 'anti_collision_min_q', 'anti_collision_max_q'):
                        if attr in p:
                            setattr(self.cfg, attr, p[attr])
                    self._power_var.set(str(self.cfg.power_dbm))
                    self._apply_advanced_to_reader()
                    self._update_buzzer_indicator()
                restored.append(name)
            if restored:
                self.cfg.save()
                self._set_status(f'Reset: {", ".join(restored)}')
            dlg.destroy()

        link_row = ttk.Frame(dlg)
        link_row.pack(fill='x', padx=20, pady=(8, 0))
        ttk.Button(link_row, text='Select all', style='Flat.TButton',
                   command=_select_all).pack(side='left')

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill='x', padx=20, pady=(8, 16))
        ttk.Button(btn_row, text='Reset selected', style='Accent.TButton',
                   command=_do_reset).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=dlg.destroy).pack(side='right')

        dlg.update_idletasks()
        pw = self.winfo_x() + self.winfo_width() // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    def _open_restore_factory_defaults_dialog(self):
        """Confirmation dialog: wipe all settings and restore factory defaults."""
        dlg = tk.Toplevel(self)
        dlg.title('Restore Factory Defaults')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg,
                  text='Restore ALL settings to factory defaults?',
                  font=UI_FONT_B).pack(padx=20, pady=(16, 6))

        ttk.Label(dlg,
                  text=(
                      'The following will be reset to their out-of-box values:\n\n'
                      '  • Output format, delimiter, prefix, suffix\n'
                      '  • Tag filter and COM port selection\n'
                      '  • Both custom preset slots (cleared)\n'
                      '  • Startup default preset → Standard\n'
                      '  • All advanced settings (power, timing, buzzer, QC)\n\n'
                      'The saved config file will be overwritten.\n'
                      'This cannot be undone.'
                  ),
                  font=UI_FONT, justify='left',
                  foreground=MUTED).pack(padx=20, pady=(0, 12))

        def _do_restore():
            self.cfg = Config()
            self.cfg.save()
            # Push every setting back to the UI widgets
            self._append_var.set(self.cfg.append_key)
            self._prefix_var.set(self.cfg.prefix)
            self._prefix_chk_var.set(bool(self.cfg.prefix))
            self._suffix_var.set(self.cfg.suffix)
            self._suffix_chk_var.set(bool(self.cfg.suffix))
            self._combine_var.set(self.cfg.combine_reads)
            self._timeout_var.set(str(self.cfg.combine_timeout_ms))
            self._spaces_var.set(self.cfg.display_spaces)
            self._ts_var.set(self.cfg.add_timestamp)
            self._delim_var.set(self.cfg.delimiter)
            self._output_format_var.set(self.cfg.output_format)
            self._filter_var.set(self.cfg.include_filter)
            self._log_var.set(self.cfg.log_enabled)
            self._port_var.set(self.cfg.port)
            self._power_var.set(str(self.cfg.power_dbm))
            # Reset preset selector and rebuild menu
            self._preset_var.set('Standard')
            self._populate_preset_menu()
            self._apply_preset('Standard')
            self._update_buzzer_indicator()
            dlg.destroy()
            self._set_status('Factory defaults restored', GREEN)

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill='x', padx=20, pady=(0, 16))
        ttk.Button(btn_row, text='Restore', style='Accent.TButton',
                   command=_do_restore).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=dlg.destroy).pack(side='right')

        dlg.update_idletasks()
        pw = self.winfo_x() + self.winfo_width() // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    # ── Advanced settings ─────────────────────────────────────────────

    def _apply_advanced_to_reader(self):
        """Push all advanced config values to the reader hardware (if connected).

        The DLL rejects configuration calls while inventory is running and prints
        "write error" to stdout.  Stop inventory first, apply, then restart.
        """
        r = self._reader
        if not r or not r.is_connected():
            return
        was_reading = r.reading
        if was_reading:
            r.stop_inventory()
            time.sleep(0.15)   # give the DLL a moment to settle before reconfiguring
        r.set_power(antenna=1, read_dbm=self.cfg.power_dbm)
        r.set_read_time(self.cfg.read_time_ms)
        r.set_idle_time(self.cfg.idle_time_ms)
        r.set_session(self.cfg.session)
        r.set_anti_collision(
            self.cfg.anti_collision,
            self.cfg.anti_collision_start_q,
            self.cfg.anti_collision_min_q,
            self.cfg.anti_collision_max_q,
        )
        r.set_rssi_threshold(self.cfg.rssi_threshold)
        level = self.cfg.buzzer_level
        r.set_buzzer(level != 'OFF', level)
        # TARGET_A for stop-after-first; TOGGLE for continuous presets so tags
        # naturally fall out of repeated singulation without Python-side timers.
        r.set_basic_target('TARGET_A' if self.cfg.stop_after_first else 'TOGGLE')
        if was_reading:
            max_tags = 1 if self.cfg.stop_after_first else 0
            fv = self.cfg.select_filter_values if self.cfg.select_mask_enabled else []
            r.start_inventory(max_tags=max_tags, filter_values=fv)

    def _open_advanced(self):
        AdvancedSettingsDialog(self,
                               is_factory=self._preset_var.get() in FACTORY_PRESETS)

    # ── Help menu actions ─────────────────────────────────────────────

    def _open_debug_log(self):
        log_path = os.path.join(
            os.path.dirname(os.path.abspath(
                sys.executable if getattr(sys, 'frozen', False) else __file__)),
            'debug.log',
        )
        if not os.path.exists(log_path):
            messagebox.showinfo('debug.log', 'No debug.log found yet — start the reader to generate one.')
            return
        try:
            os.startfile(log_path)
        except Exception:
            subprocess.Popen(['notepad.exe', log_path])

    def _open_readme(self):
        readme = os.path.join(
            os.path.dirname(os.path.abspath(
                sys.executable if getattr(sys, 'frozen', False) else __file__)),
            'README.md',
        )
        if not os.path.exists(readme):
            messagebox.showinfo('README', 'README.md not found next to the application.')
            return
        try:
            os.startfile(readme)
        except Exception:
            subprocess.Popen(['notepad.exe', readme])

    def _open_about(self):
        AboutDialog(self, self._device_info)


# ─────────────────────────────────────────────────────────────────────
class AdvancedSettingsDialog(tk.Toplevel):
    """Tabbed Advanced Settings dialog.

    Tabs: RF / Power | Scan Timing | Tag Filter | Device
    """

    SESSION_LABELS = ['S0 — re-read same tag every cycle (Hammer)',
                      'S1 — standard  (recommended)',
                      'S2 — persistent flip-flop  (Dense / conveyor)',
                      'S3 — persistent flip-flop  (alternate)']

    def __init__(self, parent: App, is_factory: bool = False):
        super().__init__(parent)
        self.parent     = parent
        self.cfg        = parent.cfg
        self.is_factory = is_factory
        self._locked_widgets: list = []   # populated during tab builds; disabled if is_factory
        self.title('Settings')
        self.configure(bg=BG)
        self.resizable(False, False)
        self.grab_set()   # modal

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=12, pady=(12, 4))

        self._tab_rf     = ttk.Frame(nb, padding=12)
        self._tab_timing = ttk.Frame(nb, padding=12)
        self._tab_filter = ttk.Frame(nb, padding=12)
        self._tab_device = ttk.Frame(nb, padding=12)

        nb.add(self._tab_rf,     text='RF / Power')
        nb.add(self._tab_timing, text='Scan Timing')
        nb.add(self._tab_filter, text='Tag Filter')
        nb.add(self._tab_device, text='Device')

        self._build_rf_tab()
        self._build_timing_tab()
        self._build_filter_tab()
        self._build_device_tab()

        # Disable factory-locked controls now that all tabs are built
        if self.is_factory:
            for w in self._locked_widgets:
                try:
                    w.configure(state='disabled')
                except Exception:
                    pass

        # ── Button row ──
        btn_row = ttk.Frame(self, padding=(12, 4))
        btn_row.pack(fill='x')
        ttk.Button(btn_row, text='Apply', style='Accent.TButton',
                   command=self._apply).pack(side='right', padx=(4, 0))
        ttk.Button(btn_row, text='Cancel', style='Flat.TButton',
                   command=self.destroy).pack(side='right')

        self._center()

    def _center(self):
        self.update_idletasks()
        pw = self.parent.winfo_x() + self.parent.winfo_width() // 2
        ph = self.parent.winfo_y() + self.parent.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f'+{pw - w // 2}+{ph - h // 2}')

    # ── RF / Power tab ────────────────────────────────────────────────

    def _build_rf_tab(self):
        f = self._tab_rf
        row = 0

        if self.is_factory:
            ttk.Label(f,
                      text='⚠  Factory preset — Session and Anti-Collision are fixed.',
                      foreground='#b7791f',
                      font=('Segoe UI', 8)).grid(
                row=row, column=0, columnspan=2, sticky='w', pady=(0, 8))
            row += 1

        # Power — query range from hardware if connected
        cur, lo, hi = self.parent._reader.get_tx_power_range() \
            if (self.parent._reader and self.parent._reader.is_connected()) \
            else (self.cfg.power_dbm, 13, 27)

        ttk.Label(f, text=f'TX Power (dBm)  [{lo}–{hi}]:').grid(
            row=row, column=0, sticky='w', pady=(0, 4))
        self._power_var = tk.StringVar(value=str(self.cfg.power_dbm))
        ttk.Spinbox(f, from_=lo, to=hi, width=6,
                    textvariable=self._power_var).grid(
            row=row, column=1, sticky='w', pady=(0, 4))
        row += 1

        # RSSI threshold
        ttk.Label(f, text='RSSI Threshold (dBm, −99 = off):').grid(
            row=row, column=0, sticky='w', pady=(0, 4))
        self._rssi_var = tk.StringVar(value=str(self.cfg.rssi_threshold))
        ttk.Spinbox(f, from_=-99, to=0, width=6,
                    textvariable=self._rssi_var).grid(
            row=row, column=1, sticky='w', pady=(0, 4))
        row += 1

        # Session
        ttk.Label(f, text='Session:').grid(row=row, column=0, sticky='w', pady=(0, 4))
        self._session_var = tk.IntVar(value=self.cfg.session)
        session_cb = ttk.Combobox(f, textvariable=self._session_var,
                                   values=list(range(4)), width=4, state='readonly')
        session_cb.grid(row=row, column=1, sticky='w', pady=(0, 4))
        self._locked_widgets.append(session_cb)
        row += 1
        ttk.Label(f, text=self.SESSION_LABELS[self.cfg.session],
                  foreground=MUTED, font=('Segoe UI', 8)).grid(
            row=row, column=0, columnspan=2, sticky='w', pady=(0, 8))
        row += 1

        # Anti-collision
        ttk.Separator(f, orient='horizontal').grid(row=row, column=0, columnspan=2,
                                                    sticky='ew', pady=(0, 8))
        row += 1
        # Anti-Collision: label in col 0, both radio buttons in a col-1 frame
        # — mirrors the label/control pattern of TX Power, RSSI, Session rows.
        self._ac_var = tk.StringVar(value=self.cfg.anti_collision)
        ttk.Label(f, text='Anti-Collision:').grid(row=row, column=0, sticky='w')
        ac_col1 = ttk.Frame(f)
        ac_col1.grid(row=row, column=1, sticky='w')
        rb_dq = ttk.Radiobutton(ac_col1, text='DynamicQ', value='DynamicQ',
                                 variable=self._ac_var)
        rb_dq.pack(side='left')
        self._locked_widgets.append(rb_dq)
        rb_fq = ttk.Radiobutton(ac_col1, text='FixedQ', value='FixedQ',
                                 variable=self._ac_var)
        rb_fq.pack(side='left', padx=(12, 0))
        self._locked_widgets.append(rb_fq)
        row += 1

        q_frame = ttk.Frame(f)
        q_frame.grid(row=row, column=0, columnspan=2, sticky='w', pady=(4, 0))
        for label, attr, default in [
            ('Start Q:', 'anti_collision_start_q', self.cfg.anti_collision_start_q),
            ('Min Q:',   'anti_collision_min_q',   self.cfg.anti_collision_min_q),
            ('Max Q:',   'anti_collision_max_q',   self.cfg.anti_collision_max_q),
        ]:
            ttk.Label(q_frame, text=label).pack(side='left', padx=(0, 2))
            var = tk.StringVar(value=str(default))
            setattr(self, f'_{attr}_var', var)
            sb = ttk.Spinbox(q_frame, from_=0, to=8, width=4, textvariable=var)
            sb.pack(side='left', padx=(0, 10))
            self._locked_widgets.append(sb)

    # ── Scan Timing tab ───────────────────────────────────────────────

    def _build_timing_tab(self):
        f = self._tab_timing
        row = 0

        if self.is_factory:
            ttk.Label(f,
                      text='⚠  Factory preset — Stop After First is fixed.',
                      foreground='#b7791f',
                      font=('Segoe UI', 8)).grid(
                row=row, column=0, columnspan=2, sticky='w', pady=(0, 8))
            row += 1

        ttk.Label(f, text='Read Time (ms per cycle):').grid(
            row=row, column=0, sticky='w', pady=(0, 6))
        self._read_time_var = tk.StringVar(value=str(self.cfg.read_time_ms))
        ttk.Spinbox(f, from_=10, to=5000, increment=50, width=7,
                    textvariable=self._read_time_var).grid(
            row=row, column=1, sticky='w', pady=(0, 6))
        row += 1

        ttk.Label(f, text='Idle Time (ms between cycles):').grid(
            row=row, column=0, sticky='w', pady=(0, 6))
        self._idle_time_var = tk.StringVar(value=str(self.cfg.idle_time_ms))
        ttk.Spinbox(f, from_=0, to=5000, increment=50, width=7,
                    textvariable=self._idle_time_var).grid(
            row=row, column=1, sticky='w', pady=(0, 6))
        row += 1

        ttk.Separator(f, orient='horizontal').grid(row=row, column=0, columnspan=2,
                                                    sticky='ew', pady=(4, 8))
        row += 1

        self._stop_first_var = tk.BooleanVar(value=self.cfg.stop_after_first)
        cb_stop = ttk.Checkbutton(f, text='Stop after first tag read (One-Shot mode)',
                                   variable=self._stop_first_var)
        cb_stop.grid(row=row, column=0, columnspan=2, sticky='w')
        self._locked_widgets.append(cb_stop)
        row += 1

        ttk.Separator(f, orient='horizontal').grid(row=row, column=0, columnspan=2,
                                                    sticky='ew', pady=(8, 8))
        row += 1

        ttk.Label(f, text='Buzzer volume:').grid(
            row=row, column=0, sticky='w', pady=(0, 4))
        row += 1
        self._buzzer_level_var = tk.StringVar(value=self.cfg.buzzer_level)
        buzzer_row = ttk.Frame(f)
        buzzer_row.grid(row=row, column=0, columnspan=2, sticky='w')
        for choice in ('OFF', 'LOW', 'HIGH'):
            ttk.Radiobutton(buzzer_row, text=choice, value=choice,
                            variable=self._buzzer_level_var).pack(
                side='left', padx=(0, 12))
        row += 1

        ttk.Separator(f, orient='horizontal').grid(row=row, column=0, columnspan=2,
                                                    sticky='ew', pady=(8, 8))
        row += 1

        self._read_tid_var = tk.BooleanVar(value=self.cfg.read_tid)
        ttk.Checkbutton(f,
                        text='Read TID after each scan (ReadMemory — tag must stay in field)',
                        variable=self._read_tid_var).grid(
            row=row, column=0, columnspan=2, sticky='w')
        row += 1

        self._check_lock_var = tk.BooleanVar(value=self.cfg.check_lock)
        ttk.Checkbutton(f,
                        text='Verify tag lock status after each scan (QC mode — alerts on Unlocked)',
                        variable=self._check_lock_var).grid(
            row=row, column=0, columnspan=2, sticky='w')

    # ── Tag Filter tab ────────────────────────────────────────────────

    # EPC packaging-level filter definitions (GS1 Gen2 filter values 0–7).
    # Reserved values (3, 5) are permanently disabled in the UI.
    _EPC_FILTER_VALUES = [
        (0, '000', 'All Others / Unclassified',               False, False),
        (1, '001', 'POS Trade Item  — individual retail unit', False, True),
        (2, '010', 'Full Case for Transport  — carton',        False, True),
        (3, '011', 'Reserved',                                 True,  False),
        (4, '100', 'Inner Pack Trade Item  — shipper',         False, True),
        (5, '101', 'Reserved',                                 True,  False),
        (6, '110', 'Unit Load  — pallet',                      False, True),
        (7, '111', 'Unit / Component Inside Product',          False, False),
    ]  # (value, binary_str, label, reserved, common)

    def _build_filter_tab(self):
        f = self._tab_filter

        # ── Enable toggle ───────────────────────────────────────────────
        self._mask_enabled_var = tk.BooleanVar(value=self.cfg.select_mask_enabled)
        ttk.Checkbutton(f, text='Enable EPC filter  (only inventory matching tags)',
                        variable=self._mask_enabled_var).pack(anchor='w', pady=(0, 10))

        # ── Section header + All / None buttons ─────────────────────────
        hdr_row = ttk.Frame(f)
        hdr_row.pack(fill='x', pady=(0, 4))
        ttk.Label(hdr_row, text='Packaging level — select one or more:',
                  font=UI_FONT_B).pack(side='left')
        ttk.Button(hdr_row, text='All', style='Flat.TButton', width=4,
                   command=self._filter_select_all).pack(side='right', padx=(4, 0))
        ttk.Button(hdr_row, text='None', style='Flat.TButton', width=4,
                   command=self._filter_select_none).pack(side='right')

        # ── Checkbox rows ───────────────────────────────────────────────
        box_frame = ttk.Frame(f, relief='sunken', borderwidth=1)
        box_frame.pack(fill='x', pady=(0, 10))

        self._filter_chk_vars: list[tk.BooleanVar] = []
        sel = set(self.cfg.select_filter_values)

        for (val, binstr, label, reserved, common) in self._EPC_FILTER_VALUES:
            var = tk.BooleanVar(value=(val in sel and not reserved))
            self._filter_chk_vars.append(var)
            var.trace_add('write', self._update_filter_mask)

            row = ttk.Frame(box_frame)
            row.pack(fill='x')

            chk = ttk.Checkbutton(row, variable=var,
                                  state='disabled' if reserved else 'normal')
            chk.pack(side='left', padx=(6, 2))

            # value · binary badge
            ttk.Label(row, text=f'{val} · {binstr}',
                      font=MONO_FONT, foreground=MUTED,
                      width=8).pack(side='left', padx=(0, 6))

            # description
            fg = MUTED if reserved else '#111111'
            ttk.Label(row, text=label, foreground=fg,
                      font=('Segoe UI', 9, 'italic') if reserved else UI_FONT
                      ).pack(side='left', padx=(0, 6))

            # COMMON badge
            if common:
                ttk.Label(row, text='COMMON', foreground=ACCENT,
                          font=('Segoe UI', 8, 'bold')).pack(side='right', padx=(0, 8))

            ttk.Separator(box_frame, orient='horizontal').pack(fill='x')

        # ── Auto-computed read-only mask field ──────────────────────────
        ttk.Label(f, text='Gen2 Select mask  (EPC bank · offset 21 · length 3) — auto-computed:',
                  foreground=MUTED, font=('Segoe UI', 8)).pack(anchor='w')
        self._filter_mask_var = tk.StringVar()
        ttk.Entry(f, textvariable=self._filter_mask_var,
                  state='readonly', font=MONO_FONT, width=38).pack(
            anchor='w', pady=(2, 0))

        # ── Multi-select info note ───────────────────────────────────────
        self._filter_info_var = tk.StringVar()
        self._filter_info_lbl = ttk.Label(f, textvariable=self._filter_info_var,
                                          foreground='#92400e',
                                          font=('Segoe UI', 8),
                                          wraplength=380)
        self._filter_info_lbl.pack(anchor='w', pady=(4, 0))

        # Initialise mask display
        self._update_filter_mask()

    def _filter_select_all(self):
        """Check all non-reserved filter value boxes."""
        for (val, _b, _l, reserved, _c), var in zip(
                self._EPC_FILTER_VALUES, self._filter_chk_vars):
            if not reserved:
                var.set(True)

    def _filter_select_none(self):
        """Uncheck all filter value boxes."""
        for var in self._filter_chk_vars:
            var.set(False)

    def _update_filter_mask(self, *_args):
        """Recompute the read-only Gen2 Select mask string from checked boxes."""
        selected = [
            val for (val, _b, _l, reserved, _c), var
            in zip(self._EPC_FILTER_VALUES, self._filter_chk_vars)
            if var.get() and not reserved
        ]
        if not selected:
            self._filter_mask_var.set('— no filter active — all tags respond —')
            self._filter_info_var.set('')
            return
        # EPC Gen2 Select: bank=EPC, pointer=21 bits, length=3 bits.
        # Mask byte = filter_value << 5  (3 bits left-aligned into a byte).
        parts = [f'0x{(v << 5):02X}' for v in selected]
        self._filter_mask_var.set('  |  '.join(parts))
        if len(selected) > 1:
            self._filter_info_var.set(
                'ℹ  Multiple values selected — one Gen2 Select command will be '
                'issued per value in sequence before inventory starts.')
        else:
            self._filter_info_var.set('')

    # ── Device tab ────────────────────────────────────────────────────

    def _build_device_tab(self):
        f = self._tab_device
        info = self.parent._device_info

        fields = [
            ('Firmware version',    info.get('fw',      '—')),
            ('Hardware version',    info.get('hw',      '—')),
            ('RFID module FW',      info.get('rfid_fw', '—')),
            ('Serial number',       info.get('sn',      '—')),
            ('SDK version',         info.get('sdk',     '—')),
            ('COM port',            self.cfg.port or '—'),
            ('Frequency band',      '902–928 MHz  (US, 50-ch FHSS)'),
        ]
        for r, (label, value) in enumerate(fields):
            ttk.Label(f, text=f'{label}:', foreground=MUTED, width=22,
                      anchor='e').grid(row=r, column=0, sticky='e', pady=2)
            ttk.Label(f, text=value, font=MONO_FONT).grid(
                row=r, column=1, sticky='w', padx=(8, 0), pady=2)

        if not self.parent._reader or not self.parent._reader.is_connected():
            ttk.Label(f, text='Connect to the reader to populate device info.',
                      foreground=MUTED, font=('Segoe UI', 8)).grid(
                row=len(fields)+1, column=0, columnspan=2, pady=(12, 0))

        ttk.Separator(f, orient='horizontal').grid(
            row=len(fields)+2, column=0, columnspan=2, sticky='ew', pady=(16, 8))
        ttk.Button(f, text='Factory Reset…', style='Flat.TButton',
                   command=self._factory_reset).grid(
            row=len(fields)+3, column=0, columnspan=2)

    def _factory_reset(self):
        if not self.parent._reader or not self.parent._reader.is_connected():
            messagebox.showwarning('Not connected', 'Connect the reader first.')
            return
        if not messagebox.askyesno(
            'Factory Reset',
            'Reset all reader settings to factory defaults?\n\n'
            'This will overwrite power, session, anti-collision and other\n'
            'hardware settings on the device itself.',
            icon='warning',
        ):
            return
        ok = self.parent._reader.factory_reset()
        if ok:
            messagebox.showinfo('Factory Reset', 'Reader reset to factory defaults.')
        else:
            messagebox.showerror('Factory Reset', 'Reset command failed — check debug.log.')

    # ── Apply ─────────────────────────────────────────────────────────

    def _apply(self):
        cfg = self.cfg

        # RF / Power
        try:
            cfg.power_dbm     = int(self._power_var.get())
        except ValueError:
            pass
        try:
            cfg.rssi_threshold = int(self._rssi_var.get())
        except ValueError:
            pass
        if not self.is_factory:
            try:
                cfg.session    = int(self._session_var.get())
            except ValueError:
                pass
            cfg.anti_collision = self._ac_var.get()
            try:
                cfg.anti_collision_start_q = int(self._anti_collision_start_q_var.get())
                cfg.anti_collision_min_q   = int(self._anti_collision_min_q_var.get())
                cfg.anti_collision_max_q   = int(self._anti_collision_max_q_var.get())
            except ValueError:
                pass

        # Scan Timing
        try:
            cfg.read_time_ms  = int(self._read_time_var.get())
            cfg.idle_time_ms  = int(self._idle_time_var.get())
        except ValueError:
            pass
        if not self.is_factory:
            cfg.stop_after_first = self._stop_first_var.get()
        cfg.buzzer_level     = self._buzzer_level_var.get()
        cfg.read_tid         = self._read_tid_var.get()
        cfg.check_lock       = self._check_lock_var.get()

        # Tag Filter
        cfg.select_mask_enabled = self._mask_enabled_var.get()
        cfg.select_filter_values = [
            val for (val, _b, _l, reserved, _c), var
            in zip(AdvancedSettingsDialog._EPC_FILTER_VALUES, self._filter_chk_vars)
            if var.get() and not reserved
        ]

        # Sync power spinbox on main UI
        self.parent._power_var.set(str(cfg.power_dbm))

        # Push to hardware and refresh buzzer indicator
        self.parent._apply_advanced_to_reader()
        self.parent._update_buzzer_indicator()
        cfg.save()
        self.destroy()


# ─────────────────────────────────────────────────────────────────────
class AboutDialog(tk.Toplevel):

    def __init__(self, parent: App, device_info: dict):
        super().__init__(parent)
        self.title(f'About {APP_TITLE}')
        self.configure(bg=BG)
        self.resizable(False, False)
        self.grab_set()

        pad = ttk.Frame(self, padding=24)
        pad.pack()

        ttk.Label(pad, text=APP_TITLE, font=TITLE_FONT,
                  foreground=ACCENT).pack()
        ttk.Label(pad, text=f'Version {APP_VERSION}',
                  foreground=MUTED).pack(pady=(2, 16))

        info_frame = ttk.Frame(pad)
        info_frame.pack(pady=(0, 16))

        rows = [
            ('COM port',        parent.cfg.port or 'Not connected'),
            ('Firmware',        device_info.get('fw',  '—')),
            ('Hardware',        device_info.get('hw',  '—')),
            ('RFID module FW',  device_info.get('rfid_fw', '—')),
            ('Serial number',   device_info.get('sn',  '—')),
            ('Frequency',       device_info.get('frequency', '902–928 MHz  (US, 50-ch FHSS)')),
        ]
        for r, (label, value) in enumerate(rows):
            ttk.Label(info_frame, text=f'{label}:', foreground=MUTED,
                      width=18, anchor='e').grid(row=r, column=0, sticky='e', pady=1)
            ttk.Label(info_frame, text=value, font=MONO_FONT).grid(
                row=r, column=1, sticky='w', padx=(8, 0), pady=1)

        ttk.Separator(pad, orient='horizontal').pack(fill='x', pady=(0, 12))
        ttk.Label(pad, text=f'\u00a9 2026 VCCS. All rights reserved.',
                  foreground=MUTED, font=('Segoe UI', 8)).pack()

        ttk.Button(pad, text='OK', style='Accent.TButton',
                   command=self.destroy).pack(pady=(16, 0))

        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width() // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f'+{pw - w // 2}+{ph - h // 2}')


# ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    app = App()
    app.mainloop()
