"""Tests: buzzer_level round-trips through JSON save/load, including back-compat migration."""
import json
import os
import tempfile
import unittest

from config import Config


class TestBuzzerLevelRoundTrip(unittest.TestCase):

    def _tmp(self):
        """Return a temp file path that does not yet exist."""
        fd, path = tempfile.mkstemp(suffix='.json')
        os.close(fd)
        os.unlink(path)
        return path

    # ------------------------------------------------------------------
    # 1. Fresh Config round-trips correctly (HIGH is the default)
    # ------------------------------------------------------------------
    def test_fresh_config_default_high(self):
        path = self._tmp()
        try:
            cfg = Config()
            self.assertEqual(cfg.buzzer_level, 'HIGH')
            cfg.save(path)
            reloaded = Config.load(path)
            self.assertEqual(reloaded.buzzer_level, 'HIGH')
        finally:
            if os.path.exists(path):
                os.unlink(path)

    # ------------------------------------------------------------------
    # 1b. save() + load() round-trips for every supported level
    # ------------------------------------------------------------------
    def test_save_reload_low(self):
        path = self._tmp()
        try:
            cfg = Config()
            cfg.buzzer_level = 'LOW'
            cfg.save(path)
            reloaded = Config.load(path)
            self.assertEqual(reloaded.buzzer_level, 'LOW')
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_save_reload_off(self):
        path = self._tmp()
        try:
            cfg = Config()
            cfg.buzzer_level = 'OFF'
            cfg.save(path)
            reloaded = Config.load(path)
            self.assertEqual(reloaded.buzzer_level, 'OFF')
        finally:
            if os.path.exists(path):
                os.unlink(path)

    # ------------------------------------------------------------------
    # 2. Old JSON with "buzzer": true  →  buzzer_level = "HIGH"
    # ------------------------------------------------------------------
    def test_old_buzzer_true_maps_to_high(self):
        path = self._tmp()
        try:
            old = {'port': 'COM1', 'buzzer': True}
            with open(path, 'w') as f:
                json.dump(old, f)
            cfg = Config.load(path)
            self.assertEqual(cfg.buzzer_level, 'HIGH')
        finally:
            if os.path.exists(path):
                os.unlink(path)

    # ------------------------------------------------------------------
    # 3. Old JSON with "buzzer": false  →  buzzer_level = "OFF"
    # ------------------------------------------------------------------
    def test_old_buzzer_false_maps_to_off(self):
        path = self._tmp()
        try:
            old = {'port': 'COM1', 'buzzer': False}
            with open(path, 'w') as f:
                json.dump(old, f)
            cfg = Config.load(path)
            self.assertEqual(cfg.buzzer_level, 'OFF')
        finally:
            if os.path.exists(path):
                os.unlink(path)

    # ------------------------------------------------------------------
    # 4. New JSON with "buzzer_level": "LOW"  loads correctly
    # ------------------------------------------------------------------
    def test_new_buzzer_level_low(self):
        path = self._tmp()
        try:
            new = {'port': 'COM2', 'buzzer_level': 'LOW'}
            with open(path, 'w') as f:
                json.dump(new, f)
            cfg = Config.load(path)
            self.assertEqual(cfg.buzzer_level, 'LOW')
        finally:
            if os.path.exists(path):
                os.unlink(path)

    # ------------------------------------------------------------------
    # 5. Extra: "buzzer_level" wins when both old and new keys present
    # ------------------------------------------------------------------
    def test_new_key_wins_over_old_key(self):
        """If a JSON somehow has both keys, buzzer_level should take precedence."""
        path = self._tmp()
        try:
            both = {'buzzer': False, 'buzzer_level': 'LOW'}
            with open(path, 'w') as f:
                json.dump(both, f)
            cfg = Config.load(path)
            # back-compat branch only fires when buzzer_level is absent
            self.assertEqual(cfg.buzzer_level, 'LOW')
        finally:
            if os.path.exists(path):
                os.unlink(path)


if __name__ == '__main__':
    unittest.main()
