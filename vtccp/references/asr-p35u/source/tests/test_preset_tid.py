"""
Test: TID reading turns off for throughput presets (Dense, Hammer, …)
      and turns on for quiet presets (Ghost, Stealth).

Exercises the *real* App._apply_preset method — not a local copy.
Runs in CI without hardware — all Windows / GUI / hardware imports are
stubbed before main.py is loaded.
"""
import sys
import types
import unittest
from unittest.mock import MagicMock, patch

# ── Stub out every platform-specific / hardware dependency ─────────────
# tkinter: make a module whose attributes are all no-op classes so that
#   class App(tk.Tk):
#   class AdvancedSettingsDialog(tk.Toplevel):
# all work without a running X / Windows display.

class _W:
    """Universal no-op widget / variable stub."""
    def __init__(self, *a, **k): pass
    def __call__(self, *a, **k): return self
    def pack(self, *a, **k): pass
    def grid(self, *a, **k): pass
    def configure(self, *a, **k): pass
    def config(self, *a, **k): pass
    def protocol(self, *a, **k): pass
    def bind(self, *a, **k): pass
    def after(self, *a, **k): pass
    def resizable(self, *a, **k): pass
    def minsize(self, *a, **k): pass
    def title(self, *a, **k): pass
    def withdraw(self, *a, **k): pass
    def deiconify(self, *a, **k): pass
    def destroy(self, *a, **k): pass
    def grab_set(self, *a, **k): pass
    def set(self, *a, **k): pass
    def get(self, *a, **k): return ''
    def insert(self, *a, **k): pass
    def delete(self, *a, **k): pass
    def focus_set(self, *a, **k): pass
    def add_cascade(self, *a, **k): pass
    def add_command(self, *a, **k): pass
    def add_separator(self, *a, **k): pass
    def add(self, *a, **k): pass
    def pack_propagate(self, *a, **k): pass
    def columnconfigure(self, *a, **k): pass
    def rowconfigure(self, *a, **k): pass
    def create_oval(self, *a, **k): return 0
    def itemconfig(self, *a, **k): pass
    def heading(self, *a, **k): pass
    def column(self, *a, **k): pass
    def tag_configure(self, *a, **k): pass
    def map(self, *a, **k): pass
    def theme_use(self, *a, **k): pass
    def yview(self, *a, **k): pass
    def see(self, *a, **k): pass


class _TkModule(types.ModuleType):
    """Tkinter stub: every attribute access returns the _W no-op class."""
    def __getattr__(self, name: str):
        return _W


_tk_mod = _TkModule('tkinter')
sys.modules['tkinter'] = _tk_mod
for _sub in ('tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox'):
    sys.modules[_sub] = _TkModule(_sub)
_tk_mod.filedialog  = sys.modules['tkinter.filedialog']
_tk_mod.messagebox  = sys.modules['tkinter.messagebox']
_tk_mod.ttk         = sys.modules['tkinter.ttk']

# winsound (Windows-only audio)
sys.modules['winsound'] = types.ModuleType('winsound')

# pystray (system-tray)
_pystray = types.ModuleType('pystray')
_pystray.Icon = _W
sys.modules['pystray'] = _pystray

# Pillow (PIL) — used by pystray for the tray icon
for _sub in ('PIL', 'PIL.Image', 'PIL.ImageDraw'):
    sys.modules[_sub] = types.ModuleType(_sub)

# injector — keyboard injection (uses WinAPI / ctypes)
_inj_mod = types.ModuleType('injector')
_inj_mod.inject = None
sys.modules['injector'] = _inj_mod

# reader — wraps the AsReader DLL (hardware + pythonnet)
_reader_mod = types.ModuleType('reader')
_reader_mod.E310Reader        = _W
_reader_mod.TagRead           = _W
_reader_mod.dll_present       = lambda: False
_reader_mod.pythonnet_present = lambda: False
sys.modules['reader'] = _reader_mod

# decoder — depends on reader types at import time
_dec_mod = types.ModuleType('decoder')
_dec_mod.decode_epc            = None
_dec_mod.format_epc_for_inject = None
_dec_mod.gtin14_check_ok       = None
sys.modules['decoder'] = _dec_mod

# ── Now it is safe to import from the tools/rfid-wedge directory ───────
import os as _os
sys.path.insert(0, _os.path.join(_os.path.dirname(__file__), '..'))

from config import Config          # noqa: E402  — pure Python, no stubs needed
import main as _main               # noqa: E402
FACTORY_PRESETS = _main.FACTORY_PRESETS
App             = _main.App


# ── Helper: build a minimal App-like instance without invoking __init__ ─
# App.__init__ builds the full Tk UI and connects to hardware.  We skip it
# entirely by using object.__new__ and setting only the attributes that
# App._apply_preset actually touches:
#
#   self.cfg                       — Config dataclass
#   self._power_var.set(str)       — updates power spinbox
#   self._preset_var.set(str)      — updates preset label
#   self._apply_advanced_to_reader()  — pushes cfg to the reader
#   self._update_buzzer_indicator()   — refreshes UI indicator
#   self._set_status(str)          — updates status bar text
#   self.cfg.save()                — persists config to disk

def _make_app(cfg: Config | None = None) -> App:
    """Return a bare App with cfg wired up and all UI side-effects stubbed."""
    app = object.__new__(App)
    app.cfg = cfg if cfg is not None else Config()

    # StringVar stubs: record the last set() value for inspection
    class _Var:
        def __init__(self):  self.value = None
        def set(self, v):    self.value = v
        def get(self):       return self.value

    app._power_var  = _Var()
    app._preset_var = _Var()

    # UI / reader side-effects — not under test here
    app._apply_advanced_to_reader = MagicMock()
    app._update_buzzer_indicator  = MagicMock()
    app._set_status               = MagicMock()
    app.cfg.save                  = MagicMock()   # don't write to disk

    return app


# ══════════════════════════════════════════════════════════════════════
class TestPresetTidDeclarations(unittest.TestCase):
    """FACTORY_PRESETS must declare read_tid with the correct boolean value."""

    TID_ON  = {'Ghost', 'Stealth'}
    TID_OFF = {'Aware', 'Standard', 'Dense', 'Hammer'}

    def test_all_factory_presets_declare_read_tid(self):
        for name in FACTORY_PRESETS:
            with self.subTest(preset=name):
                self.assertIn('read_tid', FACTORY_PRESETS[name],
                              f"'{name}' preset is missing the 'read_tid' key")

    def test_ghost_and_stealth_declare_tid_true(self):
        for name in self.TID_ON:
            with self.subTest(preset=name):
                self.assertTrue(
                    FACTORY_PRESETS[name]['read_tid'],
                    f"'{name}': FACTORY_PRESETS should have read_tid=True",
                )

    def test_throughput_presets_declare_tid_false(self):
        for name in self.TID_OFF:
            with self.subTest(preset=name):
                self.assertFalse(
                    FACTORY_PRESETS[name]['read_tid'],
                    f"'{name}': FACTORY_PRESETS should have read_tid=False",
                )


# ══════════════════════════════════════════════════════════════════════
class TestApplyPresetFlipsTid(unittest.TestCase):
    """App._apply_preset (the real method) must set cfg.read_tid correctly."""

    # ── presets that must enable TID ──────────────────────────────────

    def test_ghost_enables_tid(self):
        app = _make_app()
        app.cfg.read_tid = False          # start opposite to prove the flip
        app._apply_preset('Ghost')
        self.assertTrue(app.cfg.read_tid, "Ghost preset should set read_tid=True")

    def test_stealth_enables_tid(self):
        app = _make_app()
        app.cfg.read_tid = False
        app._apply_preset('Stealth')
        self.assertTrue(app.cfg.read_tid, "Stealth preset should set read_tid=True")

    # ── presets that must disable TID ─────────────────────────────────

    def test_aware_disables_tid(self):
        app = _make_app()
        app.cfg.read_tid = True
        app._apply_preset('Aware')
        self.assertFalse(app.cfg.read_tid, "Aware preset should set read_tid=False")

    def test_standard_disables_tid(self):
        app = _make_app()
        app.cfg.read_tid = True
        app._apply_preset('Standard')
        self.assertFalse(app.cfg.read_tid, "Standard preset should set read_tid=False")

    def test_dense_disables_tid(self):
        app = _make_app()
        app.cfg.read_tid = True
        app._apply_preset('Dense')
        self.assertFalse(app.cfg.read_tid, "Dense preset should set read_tid=False")

    def test_hammer_disables_tid(self):
        app = _make_app()
        app.cfg.read_tid = True
        app._apply_preset('Hammer')
        self.assertFalse(app.cfg.read_tid, "Hammer preset should set read_tid=False")

    # ── preset-switching scenarios ────────────────────────────────────

    def test_switching_ghost_to_dense_disables_tid(self):
        """Switching from a TID-on preset to Dense must clear read_tid."""
        app = _make_app()
        app._apply_preset('Ghost')
        self.assertTrue(app.cfg.read_tid, "sanity: Ghost should enable TID")
        app._apply_preset('Dense')
        self.assertFalse(app.cfg.read_tid,
                         "Switching Ghost→Dense must set read_tid=False")

    def test_switching_ghost_to_hammer_disables_tid(self):
        """Switching from a TID-on preset to Hammer must clear read_tid."""
        app = _make_app()
        app._apply_preset('Ghost')
        self.assertTrue(app.cfg.read_tid, "sanity: Ghost should enable TID")
        app._apply_preset('Hammer')
        self.assertFalse(app.cfg.read_tid,
                         "Switching Ghost→Hammer must set read_tid=False")

    def test_switching_hammer_to_stealth_enables_tid(self):
        """Switching from Hammer back to Stealth must restore read_tid."""
        app = _make_app()
        app._apply_preset('Hammer')
        self.assertFalse(app.cfg.read_tid, "sanity: Hammer should disable TID")
        app._apply_preset('Stealth')
        self.assertTrue(app.cfg.read_tid,
                        "Switching Hammer→Stealth must set read_tid=True")

    def test_switching_aware_to_ghost_enables_tid(self):
        app = _make_app()
        app._apply_preset('Aware')
        self.assertFalse(app.cfg.read_tid, "sanity: Aware should disable TID")
        app._apply_preset('Ghost')
        self.assertTrue(app.cfg.read_tid,
                        "Switching Aware→Ghost must set read_tid=True")

    def test_repeated_throughput_cycle_stays_off(self):
        """Cycling between TID-off presets must never re-enable read_tid."""
        app = _make_app()
        for preset in ('Dense', 'Hammer', 'Dense', 'Standard', 'Hammer', 'Aware'):
            app._apply_preset(preset)
            self.assertFalse(
                app.cfg.read_tid,
                f"read_tid must remain False after applying '{preset}'",
            )

    # ── side-effect verification ──────────────────────────────────────

    def test_apply_calls_reader_update(self):
        """_apply_preset must push the changed config to the reader."""
        app = _make_app()
        app._apply_preset('Dense')
        app._apply_advanced_to_reader.assert_called()

    def test_apply_saves_config(self):
        """_apply_preset must persist the new config to disk."""
        app = _make_app()
        app._apply_preset('Ghost')
        app.cfg.save.assert_called()


# ══════════════════════════════════════════════════════════════════════
class TestCustomPresetTid(unittest.TestCase):
    """Custom preset slots must carry read_tid through the save/load cycle."""

    # ── helpers ───────────────────────────────────────────────────────

    def _make_app_with_custom(self, slot_index: int, read_tid: bool) -> App:
        """Return an App whose custom preset slot *slot_index* has read_tid set."""
        cfg = Config()
        # Slots default name is '' (empty = unused).  Give it a name so
        # _apply_preset does not bail out early with "slot empty".
        cfg.custom_presets[slot_index]['name']     = f'My Custom {slot_index + 1}'
        cfg.custom_presets[slot_index]['read_tid'] = read_tid
        return _make_app(cfg)

    # ── Custom 1 ─────────────────────────────────────────────────────

    def test_custom1_read_tid_false_is_applied(self):
        """Applying Custom 1 with read_tid=False must set cfg.read_tid=False."""
        app = self._make_app_with_custom(0, read_tid=False)
        app.cfg.read_tid = True       # start opposite to prove the flip
        app._apply_preset('Custom 1')
        self.assertFalse(
            app.cfg.read_tid,
            "Custom 1 preset saved with read_tid=False must set cfg.read_tid=False",
        )

    def test_custom1_read_tid_true_is_applied(self):
        """Applying Custom 1 with read_tid=True must set cfg.read_tid=True."""
        app = self._make_app_with_custom(0, read_tid=True)
        app.cfg.read_tid = False      # start opposite to prove the flip
        app._apply_preset('Custom 1')
        self.assertTrue(
            app.cfg.read_tid,
            "Custom 1 preset saved with read_tid=True must set cfg.read_tid=True",
        )

    # ── Custom 2 ─────────────────────────────────────────────────────

    def test_custom2_read_tid_false_is_applied(self):
        """Applying Custom 2 with read_tid=False must set cfg.read_tid=False."""
        app = self._make_app_with_custom(1, read_tid=False)
        app.cfg.read_tid = True
        app._apply_preset('Custom 2')
        self.assertFalse(
            app.cfg.read_tid,
            "Custom 2 preset saved with read_tid=False must set cfg.read_tid=False",
        )

    def test_custom2_read_tid_true_is_applied(self):
        """Applying Custom 2 with read_tid=True must set cfg.read_tid=True."""
        app = self._make_app_with_custom(1, read_tid=True)
        app.cfg.read_tid = False
        app._apply_preset('Custom 2')
        self.assertTrue(
            app.cfg.read_tid,
            "Custom 2 preset saved with read_tid=True must set cfg.read_tid=True",
        )

    # ── empty slot guard ──────────────────────────────────────────────

    def test_empty_custom_slot_is_ignored(self):
        """An unconfigured slot (name='') must not change cfg at all."""
        cfg = Config()
        cfg.custom_presets[0]['name']     = ''    # slot is empty
        cfg.custom_presets[0]['read_tid'] = False
        app = _make_app(cfg)
        app.cfg.read_tid = True
        app._apply_preset('Custom 1')
        self.assertTrue(
            app.cfg.read_tid,
            "An empty custom slot must not alter cfg.read_tid",
        )

    # ── side-effect verification ──────────────────────────────────────

    def test_custom_preset_calls_reader_update(self):
        """Applying a custom preset must push the changed config to the reader."""
        app = self._make_app_with_custom(0, read_tid=False)
        app._apply_preset('Custom 1')
        app._apply_advanced_to_reader.assert_called()

    def test_custom_preset_saves_config(self):
        """Applying a custom preset must persist the new config to disk."""
        app = self._make_app_with_custom(1, read_tid=True)
        app._apply_preset('Custom 2')
        app.cfg.save.assert_called()

    # ── switching scenarios ───────────────────────────────────────────

    def test_switching_factory_tid_on_to_custom_tid_off(self):
        """Ghost (TID on) → Custom 1 (TID off) must clear read_tid."""
        app = self._make_app_with_custom(0, read_tid=False)
        app._apply_preset('Ghost')
        self.assertTrue(app.cfg.read_tid, "sanity: Ghost should enable TID")
        app._apply_preset('Custom 1')
        self.assertFalse(
            app.cfg.read_tid,
            "Switching Ghost→Custom 1 (read_tid=False) must disable TID",
        )

    def test_switching_factory_tid_off_to_custom_tid_on(self):
        """Dense (TID off) → Custom 2 (TID on) must enable read_tid."""
        app = self._make_app_with_custom(1, read_tid=True)
        app._apply_preset('Dense')
        self.assertFalse(app.cfg.read_tid, "sanity: Dense should disable TID")
        app._apply_preset('Custom 2')
        self.assertTrue(
            app.cfg.read_tid,
            "Switching Dense→Custom 2 (read_tid=True) must enable TID",
        )


# ══════════════════════════════════════════════════════════════════════
class TestCustomPresetPersistence(unittest.TestCase):
    """Custom preset read_tid must survive the full save → load → apply cycle.

    These tests mirror exactly what _open_save_as_dialog._do_save() does:
      1. Copy current cfg fields into cfg.custom_presets[idx] (the same dict
         construction used by _do_save at main.py ~line 1269-1286).
      2. Persist via cfg.save() to a real temporary file.
      3. Reload into a fresh Config via Config.load().
      4. Apply the slot with _apply_preset and assert cfg.read_tid.

    A regression in Config.save/load, or a missing 'read_tid' key in the
    slot-capture dict, will be caught here.
    """

    # ── internal helper: replicate _do_save's slot-capture dict ──────

    @staticmethod
    def _capture_slot(cfg: Config, slot_index: int, name: str) -> None:
        """Write current cfg fields into custom_presets[slot_index].

        This is the same dict construction as _open_save_as_dialog._do_save()
        (main.py ~line 1269-1286), extracted so tests can call it without a
        running Tk window.
        """
        cfg.custom_presets[slot_index] = dict(
            name=name,
            power_dbm=cfg.power_dbm,
            buzzer_level=cfg.buzzer_level,
            combine_reads=cfg.combine_reads,
            combine_timeout_ms=cfg.combine_timeout_ms,
            stop_after_first=cfg.stop_after_first,
            read_time_ms=cfg.read_time_ms,
            idle_time_ms=cfg.idle_time_ms,
            session=cfg.session,
            anti_collision=cfg.anti_collision,
            anti_collision_start_q=cfg.anti_collision_start_q,
            anti_collision_min_q=cfg.anti_collision_min_q,
            anti_collision_max_q=cfg.anti_collision_max_q,
            rssi_threshold=cfg.rssi_threshold,
            read_tid=cfg.read_tid,
            check_lock=cfg.check_lock,
        )

    # ── tests ─────────────────────────────────────────────────────────

    def test_custom1_tid_false_survives_save_load(self):
        """Custom 1 with throughput settings (read_tid=False) persists correctly."""
        import tempfile, os

        # 1. Build a config as if the user applied a throughput preset then saved
        #    their own custom slot.
        cfg = Config()
        cfg.read_tid = False          # throughput setting chosen by the user
        self._capture_slot(cfg, 0, 'My Throughput')

        with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as f:
            tmp = f.name
        try:
            cfg.save(tmp)

            # 2. Reload from disk — simulates app restart.
            cfg2 = Config.load(tmp)

            # 3. Apply the slot and assert.
            app = _make_app(cfg2)
            app.cfg.read_tid = True   # start opposite to prove the load worked
            app._apply_preset('Custom 1')
            self.assertFalse(
                app.cfg.read_tid,
                "Custom 1 (read_tid=False) must remain False after save → load → apply",
            )
        finally:
            os.unlink(tmp)

    def test_custom1_tid_true_survives_save_load(self):
        """Custom 1 with TID-on settings persists correctly."""
        import tempfile, os

        cfg = Config()
        cfg.read_tid = True
        self._capture_slot(cfg, 0, 'My TID Slot')

        with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as f:
            tmp = f.name
        try:
            cfg.save(tmp)
            cfg2 = Config.load(tmp)

            app = _make_app(cfg2)
            app.cfg.read_tid = False
            app._apply_preset('Custom 1')
            self.assertTrue(
                app.cfg.read_tid,
                "Custom 1 (read_tid=True) must remain True after save → load → apply",
            )
        finally:
            os.unlink(tmp)

    def test_custom2_tid_false_survives_save_load(self):
        """Custom 2 with throughput settings (read_tid=False) persists correctly."""
        import tempfile, os

        cfg = Config()
        cfg.read_tid = False
        self._capture_slot(cfg, 1, 'Throughput Slot 2')

        with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as f:
            tmp = f.name
        try:
            cfg.save(tmp)
            cfg2 = Config.load(tmp)

            app = _make_app(cfg2)
            app.cfg.read_tid = True
            app._apply_preset('Custom 2')
            self.assertFalse(
                app.cfg.read_tid,
                "Custom 2 (read_tid=False) must remain False after save → load → apply",
            )
        finally:
            os.unlink(tmp)

    def test_custom2_tid_true_survives_save_load(self):
        """Custom 2 with TID-on settings persists correctly."""
        import tempfile, os

        cfg = Config()
        cfg.read_tid = True
        self._capture_slot(cfg, 1, 'TID Slot 2')

        with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as f:
            tmp = f.name
        try:
            cfg.save(tmp)
            cfg2 = Config.load(tmp)

            app = _make_app(cfg2)
            app.cfg.read_tid = False
            app._apply_preset('Custom 2')
            self.assertTrue(
                app.cfg.read_tid,
                "Custom 2 (read_tid=True) must remain True after save → load → apply",
            )
        finally:
            os.unlink(tmp)

    def test_throughput_derived_settings_capture_tid_false(self):
        """Capturing settings after applying a throughput preset stores read_tid=False.

        Simulates the user workflow: apply Dense (which sets cfg.read_tid=False),
        then save to a custom slot.  The saved slot must record read_tid=False,
        and re-applying the slot must keep it False.
        """
        import tempfile, os

        cfg = Config()
        # Replicate what _apply_preset('Dense') would do to cfg
        for attr, val in _main.FACTORY_PRESETS['Dense'].items():
            if attr != 'label' and hasattr(cfg, attr):
                setattr(cfg, attr, val)
        self.assertFalse(cfg.read_tid, "sanity: Dense sets read_tid=False")

        self._capture_slot(cfg, 0, 'Dense Clone')

        with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as f:
            tmp = f.name
        try:
            cfg.save(tmp)
            cfg2 = Config.load(tmp)

            app = _make_app(cfg2)
            app.cfg.read_tid = True   # force opposite so we prove the preset
            app._apply_preset('Custom 1')
            self.assertFalse(
                app.cfg.read_tid,
                "Custom slot derived from Dense must apply read_tid=False after save/load",
            )
        finally:
            os.unlink(tmp)


if __name__ == '__main__':
    unittest.main(verbosity=2)
