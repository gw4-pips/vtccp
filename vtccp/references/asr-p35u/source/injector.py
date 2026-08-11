"""
Copyright © 2026 VCCS. All rights reserved.
RFID FlexWedge™ Pro — proprietary software of VCCS.

Keyboard injection via pynput.
Types text into whatever window currently has focus — works with
Excel, VTCCP, Notepad, or any other Windows application.
"""
import time
import logging
from pynput.keyboard import Controller, Key

logger = logging.getLogger(__name__)

_keyboard = Controller()

_KEY_MAP = {
    'Enter':   Key.enter,
    'Tab':     Key.tab,
    'Nothing': None,
}


def inject(text: str, append_key: str = 'Enter', pre_delay_ms: int = 0):
    """
    Type `text` into the focused application, then press the terminator key.

    pre_delay_ms: optional settle time before typing (useful when the
                  caller switches focus immediately before injecting).
    """
    if pre_delay_ms > 0:
        time.sleep(pre_delay_ms / 1000.0)
    try:
        _keyboard.type(text)
        terminator = _KEY_MAP.get(append_key)
        if terminator is not None:
            time.sleep(0.02)
            _keyboard.press(terminator)
            _keyboard.release(terminator)
    except Exception as exc:
        logger.error('Keyboard injection error: %s', exc)
