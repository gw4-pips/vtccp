"""
Copyright © 2026 VCCS. All rights reserved.
RFID FlexWedge™ Pro — proprietary software of VCCS.

AsReader P35U UHF RFID Reader — C# SDK handler via pythonnet.
Replaces the defunct Winnix/E310 serial protocol handler.

Requires:
  - pythonnet (pip install pythonnet)
  - AsReaderP3xU.dll (v1.3.0) copied next to reader.py / the .exe

Public interface is identical to the old E310Reader so main.py is unchanged.
"""
import os
import sys
import time
import logging
from dataclasses import dataclass, field
from typing import Optional, Callable
import serial.tools.list_ports

logger = logging.getLogger(__name__)

# Debug log written next to reader.py / the .exe
_LOG_PATH     = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'debug.log')
_LOG_PATH_BAK = _LOG_PATH + '.1'
_MAX_LOG_BYTES = 500 * 1024   # 500 KB

import threading
_log_lock = threading.Lock()


def _rotate_log() -> None:
    """Rename debug.log → debug.log.1 when the live file is over the size cap.

    Called once at module import (catches growth from previous sessions) and
    again inside _dbg() after each write (caps growth within a single session).
    Runs under _log_lock so it is safe to call from multiple threads.
    """
    try:
        if os.path.exists(_LOG_PATH) and os.path.getsize(_LOG_PATH) > _MAX_LOG_BYTES:
            if os.path.exists(_LOG_PATH_BAK):
                os.remove(_LOG_PATH_BAK)
            os.rename(_LOG_PATH, _LOG_PATH_BAK)
    except OSError:
        pass   # non-fatal — logging should never crash the app


# Rotate once at import time so a new session always starts with a small file
_rotate_log()


def _dbg(msg: str):
    """Append a timestamped line to debug.log, rotating when the cap is hit."""
    line = f'{time.strftime("%H:%M:%S")} {msg}\n'
    with _log_lock:
        with open(_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line)
        # Rotate mid-session if this session alone exceeds the cap
        _rotate_log()


# ---------------------------------------------------------------------------
# USB VID/PID constants for the AsReader P35U (CP210x USB-to-serial bridge)
# ---------------------------------------------------------------------------
# VID 0x339C / PID 0x271B — confirmed on actual AsReader P35U hardware
# (Windows reports this as "USB Serial Device" on COM4)
# Note: NOT the Silicon Labs CP210x values (0x10C4/0xEA60) that the datasheet
# suggested — the real unit uses a different USB bridge chip.
_ASREADER_VID: int = 0x339C
_ASREADER_PID: int = 0x271B

# Keywords used as a fallback when VID/PID matching finds nothing
# (covers unlisted PIDs, third-party CP210x devices with different firmware,
# or when the OS reports None for VID/PID)
_PORT_KEYWORDS = ('asreader', 'silicon labs', 'cp210', 'usb serial', 'usb-serial')


@dataclass
class TagRead:
    """Represents a single tag read returned from the reader.

    rssi_raw stores the RSSI dBm value as an integer (negative, e.g. -65).
    freq_khz is None for the P35U (SDK does not expose channel frequency directly).
    tid is the manufacturer-assigned Tag ID (read-only, permanent); empty string
    when the SDK returns None (some tags / firmware builds omit it).
    """
    raw_epc:  str
    pc:       int
    rssi_raw: int                # dBm integer from SDK (stored as int for compat)
    antenna:  int
    freq_khz: Optional[int] = None
    timestamp: float = field(default_factory=time.time)
    tid:      str = ''

    @property
    def rssi_dbm(self) -> float:
        return float(self.rssi_raw)

    @property
    def freq_mhz(self) -> Optional[float]:
        return self.freq_khz / 1000.0 if self.freq_khz is not None else None


# ---------------------------------------------------------------------------
# DLL loading
# ---------------------------------------------------------------------------

def _dll_dir() -> str:
    """Return the directory that should contain AsReaderP3xU.dll."""
    if getattr(sys, 'frozen', False):          # running as PyInstaller .exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def pythonnet_present() -> bool:
    """Return True if pythonnet (clr) can be imported."""
    try:
        import clr  # noqa: F401
        return True
    except ImportError:
        return False
def dll_present() -> bool:
    """Return True if AsReaderP3xU.dll exists next to the .exe / script."""
    return os.path.exists(os.path.join(_dll_dir(), 'AsReaderP3xU.dll'))


def _load_dll():
    """Load AsReaderP3xU.dll via pythonnet.

    Returns:
        tuple[AsReader class, Types module]

    Raises:
        RuntimeError  if pythonnet is not installed
        FileNotFoundError  if the DLL is missing
    """
    try:
        import clr  # pythonnet
    except ImportError:
        raise RuntimeError(
            "pythonnet is not installed. Install it with:  pip install pythonnet")

    base_dir = _dll_dir()
    dll_path = os.path.join(base_dir, 'AsReaderP3xU.dll')
    if not os.path.exists(dll_path):
        raise FileNotFoundError(
            f"AsReaderP3xU.dll not found at:\n  {dll_path}\n"
            "Copy the DLL from the SDK zip next to reader.py (or next to the .exe).")

    # Add the folder to sys.path so the CLR resolver finds the DLL by name
    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)

    clr.AddReference('AsReaderP3xU')

    from AsReaderP3xU import AsReader, Types  # type: ignore[import]
    return AsReader, Types


# ---------------------------------------------------------------------------
# Port auto-detection
# ---------------------------------------------------------------------------

def _find_asreader_ports() -> list[str]:
    """Return COM ports that match the AsReader P35U, in priority order.

    Priority 1 — VID/PID match (0x10C4:0xEA60, CP210x).
        Most reliable: uniquely identifies the Silicon Labs USB bridge used
        by the P35U regardless of how the OS labels the port description.

    Priority 2 — Description/manufacturer keyword match ('Silicon Labs',
        'CP210', 'AsReader', etc.).
        Kept as a fallback for units whose OS driver reports None for VID/PID,
        or future hardware revisions with a different PID.

    Priority 3 — All ports.
        Last resort so the user can still pick a port manually when nothing
        auto-matches.
    """
    all_ports = list(serial.tools.list_ports.comports())

    # --- Primary: VID/PID ---
    vid_pid_matched = [
        p.device for p in all_ports
        if p.vid == _ASREADER_VID and p.pid == _ASREADER_PID
    ]
    if vid_pid_matched:
        _dbg(f'_find_asreader_ports: VID/PID match → {vid_pid_matched}')
        return sorted(vid_pid_matched)

    # --- Secondary: keyword match ---
    keyword_matched = [
        p.device for p in all_ports
        if any(kw in (p.description or '').lower() or
               kw in (p.manufacturer or '').lower()
               for kw in _PORT_KEYWORDS)
    ]
    if keyword_matched:
        _dbg(f'_find_asreader_ports: keyword match → {keyword_matched}')
        return sorted(keyword_matched)

    # --- Tertiary: return all ports so the user can pick manually ---
    _dbg('_find_asreader_ports: no match — returning all ports')
    return sorted(p.device for p in all_ports)


# ---------------------------------------------------------------------------
# Reader class
# ---------------------------------------------------------------------------

class P35UReader:
    """
    AsReader P35U UHF reader, driven by the manufacturer's C# SDK via pythonnet.

    The public API (list_ports, connect, disconnect, is_connected,
    start_inventory, stop_inventory, set_power, reading, TagRead) is identical
    to the old E310Reader so main.py requires no changes.
    """

    def __init__(self,
                 on_tag:        Callable[[TagRead], None],
                 on_status:     Callable[[str], None],
                 on_raw:        Optional[Callable[[bytes], None]] = None,
                 on_disconnect: Optional[Callable[[], None]] = None,
                 on_stopped:    Optional[Callable[[], None]] = None,
                 on_trigger:    Optional[Callable[[int], None]] = None):
        self._on_tag        = on_tag
        self._on_status     = on_status
        # on_raw kept for API compatibility — the SDK doesn't expose raw bytes
        self._on_raw        = on_raw
        # Called when an unexpected hardware disconnect is detected from a
        # callback thread.  The UI layer uses this to trigger a full teardown.
        self._on_disconnect = on_disconnect
        # Called when the hardware auto-stops due to mtnu limit (e.g. mtnu=1).
        # Lets the UI update dot/button state without polling.
        self._on_stopped    = on_stopped
        # Called when the hardware SCAN button is pressed (trigger==1) or
        # released (trigger==0).  Drives trigger-mode presets in the UI layer.
        self._on_trigger    = on_trigger

        self._device              = None   # AsReader instance (C# object)
        self._Types               = None   # cached Types module from DLL load
        self._connected           = False
        self.reading              = False
        # Set True by stop_inventory() so _cb_complete can distinguish an
        # intentional stop from an unexpected hardware pull.
        self._intentional_stop    = False
        # Set True when start_inventory is called with max_tags > 0 so
        # _cb_complete knows to treat the next complete_status=True as a
        # hardware-triggered stop rather than a normal continuous-mode round.
        self._hw_stop_requested   = False
        # One-shot callback set before a ReadMemory / other command call;
        # cleared and invoked by _cb_command when the DLL delivers the response.
        self._pending_command_cb  = None
        # Alternative hook: if ReadMemory fires _cb_tag instead of _cb_command
        # (observed on FW 1.8.0), this captures the result and suppresses the
        # normal inventory-tag pipeline for that one callback.
        self._pending_memory_cb   = None

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    @staticmethod
    def list_ports() -> list[str]:
        """Return candidate COM ports (AsReader-matched first, all ports as fallback)."""
        return _find_asreader_ports()

    def connect(self, port: str, baud: int = 115200) -> bool:
        """Load the DLL, instantiate AsReader, connect via VCP, set US region.

        baud is ignored (the SDK manages baud internally) but kept for signature compat.
        """
        _dbg(f'--- connect {port} ---')
        try:
            AsReaderCls, Types = _load_dll()

            dev = AsReaderCls()
            self._device = dev

            # Register all six callbacks before connecting
            self._setup_delegates(AsReaderCls, dev)

            ret = dev.ConnectWithVCP(port)
            if ret != 0:
                _dbg(f'ConnectWithVCP returned {ret} — connection refused')
                self._on_status(f'Connection failed (code {ret})')
                self._device = None
                return False

            # Cache Types for later use in advanced-settings methods
            self._Types = Types

            # Region MUST be set before StartInventory
            dev.SetRegion(Types.RegionType.REGION_US)
            _dbg('SetRegion REGION_US OK')

            self._connected = True
            self._on_status(f'Connected — {port}')
            _dbg('connected OK')

            # Enumerate public DLL methods once so we can see what's available
            try:
                attrs = sorted(a for a in dir(dev) if not a.startswith('_'))
                _dbg(f'DLL public members: {attrs}')
            except Exception as _e:
                _dbg(f'DLL member enum failed: {_e}')

            # Log every type/enum in the DLL namespace — reveals HIDInventoryMode,
            # InventoryType variants, and other enum values we can pass to Set* calls
            try:
                type_names = sorted(a for a in dir(Types) if not a.startswith('_'))
                _dbg(f'DLL Types namespace: {type_names}')
                # For each enum type, log its members so we know valid values
                for tname in type_names:
                    try:
                        t = getattr(Types, tname)
                        members = [a for a in dir(t) if not a.startswith('_')
                                   and a not in ('Equals','GetHashCode','GetType',
                                                 'GetTypeCode','MemberwiseClone',
                                                 'ReferenceEquals','ToString',
                                                 'CompareTo','HasFlag','value__')]
                        if members:
                            _dbg(f'  Types.{tname}: {members}')
                    except Exception:
                        pass
            except Exception as _e:
                _dbg(f'DLL Types enum failed: {_e}')

            # Probe GetHIDInventoryMode — may return current TID/EPC mode setting
            try:
                hid_mode = dev.GetHIDInventoryMode()
                _dbg(f'GetHIDInventoryMode() → {hid_mode!r} (type={type(hid_mode).__name__})')
                if hid_mode is not None:
                    try:
                        mode_attrs = {a: getattr(hid_mode, a)
                                      for a in dir(hid_mode) if not a.startswith('_')}
                        _dbg(f'  HIDInventoryMode fields: {mode_attrs}')
                    except Exception:
                        pass
            except Exception as _e:
                _dbg(f'GetHIDInventoryMode() failed: {_e}')

            # GetHIDWorkParams requires at least one argument but its expected type
            # is unknown.  Passing wrong managed types crashes pythonnet at the
            # binding layer before Python sees the exception, so we do not probe it
            # blindly.  Once Types.HIDWorkParams (or similar) is identified from the
            # Types namespace log above, a typed probe can be added here.
            _dbg('GetHIDWorkParams: skipping blind probe (type unknown — see Types log)')

            return True

        except FileNotFoundError as exc:
            _dbg(f'DLL not found: {exc}')
            self._on_status(f'DLL missing: {exc}')
            return False
        except Exception as exc:
            _dbg(f'connect error: {exc}')
            self._on_status(f'Connection failed: {exc}')
            return False

    def disconnect(self):
        self.stop_inventory()
        if self._device is not None:
            try:
                self._device.DisConnect()
            except Exception as exc:
                _dbg(f'DisConnect error: {exc}')
            self._device = None
        self._connected = False
        _dbg('disconnected')
        self._on_status('Disconnected')

    def is_connected(self) -> bool:
        return self._connected

    # ------------------------------------------------------------------
    # Inventory
    # ------------------------------------------------------------------

    def set_inventory_type(self, use_tid: bool) -> bool:
        """Switch to PC_EPC_TID inventory mode so td.tid is populated in _cb_tag.

        Tries three paths in order, logging each attempt:
          1. SetInventoryType(InventoryType.PC_EPC_TID)  — SDK method, not on all DLLs
          2. SetHIDInventoryMode(InventoryType.PC_EPC_TID) — alternate entry point
          3. SetHIDInventoryMode with every HIDInventoryMode enum value — last resort probe

        Returns True if any path was accepted; False if all fail (safe to ignore —
        ReadMemory fallback will still run).
        """
        if not self.is_connected() or self._device is None or self._Types is None:
            return False

        T = self._Types

        # ── Path 1: SetInventoryType (confirmed absent on current DLL) ──────
        try:
            inv_type = T.InventoryType.PC_EPC_TID if use_tid else T.InventoryType.PC_EPC_RSSI
            ret = self._device.SetInventoryType(inv_type)
            _dbg(f'SetInventoryType {"PC_EPC_TID" if use_tid else "PC_EPC_RSSI"} → {ret}')
            return ret == 0
        except AttributeError:
            _dbg('SetInventoryType: not on this DLL')
        except Exception as exc:
            _dbg(f'SetInventoryType error: {exc}')

        if not use_tid:
            return False   # nothing to revert to for non-TID path

        # SetHIDInventoryMode — cannot be probed safely without knowing the
        # exact C# parameter type.  Passing enum values of the wrong managed
        # type causes System.ArgumentException in pythonnet's MethodBinder
        # BEFORE Python sees the exception, terminating the process.
        # Omitted pending ASReader tech-support response on correct signature.
        _dbg('SetHIDInventoryMode: skipped (type signature unknown — see defect report)')
        return False

    def apply_select_filter(self, filter_values: list) -> None:
        """Issue Gen2 Select commands for each packaging-level filter value.

        Each filter value (0–7) maps to an EPC bank Gen2 Select command:
          - Bank:    EPC (01)
          - Pointer: 21  bits from bank start (PC bits 10–8 = EPC bank bits 21–23)
          - Length:  3   bits
          - Mask:    filter_value << 5  (3 bits left-aligned into a byte)

        One command is issued per value.  The reader accumulates them and will
        respond to any tag whose EPC filter field matches at least one command
        (OR semantics per Gen2 spec §6.3.2.11.2.2, Action=000/001 combination).

        If the SDK does not expose a SetSelect / SetEPCFilter method, the call
        is logged and skipped silently — inventory will proceed unfiltered rather
        than crashing.
        """
        if not self.is_connected() or self._device is None:
            return
        if not filter_values:
            _dbg('apply_select_filter: no filter values — skipping')
            return

        for val in filter_values:
            mask_byte = (int(val) & 0x07) << 5   # 3-bit filter → bits 7-5 of byte
            mask_bytes = bytes([mask_byte])
            _dbg(f'apply_select_filter: value={val} mask_byte=0x{mask_byte:02X}')
            try:
                # Most AsReader SDKs expose SetSelect(bank, ptr_bits, len_bits, mask_bytes)
                ret = self._device.SetSelect(1, 21, 3, mask_bytes)
                _dbg(f'  SetSelect → {ret}')
            except AttributeError:
                # Try alternate names used in some SDK revisions
                try:
                    ret = self._device.SetEPCFilter(21, 3, mask_bytes)
                    _dbg(f'  SetEPCFilter → {ret}')
                except AttributeError:
                    _dbg('  SetSelect/SetEPCFilter: not on this DLL version — filter skipped')
                    return   # no point looping if neither method exists
                except Exception as exc:
                    _dbg(f'  SetEPCFilter error: {exc}')
            except Exception as exc:
                _dbg(f'  SetSelect error: {exc}')

    def start_inventory(self, max_tags: int = 0, use_tid: bool = False,
                        filter_values: list | None = None):
        """Start inventory.

        StartInventory args (from SDK):
          rssiEn  – True      → RSSI included in each InventoryResult
          mtnu    – max_tags  → 0 = unlimited; 1 = hardware stops after first tag
          mtime   – 0         → no time limit per round
          rc      – 0         → continuous (unlimited rounds)
          ant1    – True      → antenna 1 enabled

        Pass max_tags=1 for stop-after-first presets: the hardware stops itself
        after one tag is read, eliminating the Python-timer race that causes a
        second inventory cycle (and a second beep) to fire before StopInventory
        takes effect.

        Pass use_tid=True to attempt PC_EPC_TID mode so td.tid is populated
        in _cb_tag without a separate ReadMemory command.

        Pass filter_values as a list of ints (0–7) to issue Gen2 Select
        commands before starting inventory.  Empty list or None = no filter.
        """
        if not self.is_connected() or self._device is None:
            return
        try:
            if use_tid:
                self.set_inventory_type(True)
            # Apply EPC packaging-level Gen2 Select filter before starting
            if filter_values:
                self.apply_select_filter(filter_values)
            self._hw_stop_requested = max_tags > 0
            ret = self._device.StartInventory(True, max_tags, 0, 0, True)
            if ret == 0:
                self.reading = True
                self._on_status('Reading…')
                _dbg(f'inventory started (max_tags={max_tags})')
            else:
                self._hw_stop_requested = False
                _dbg(f'StartInventory returned {ret}')
                self._on_status(f'Start failed (code {ret})')
        except Exception as exc:
            self._hw_stop_requested = False
            _dbg(f'start_inventory error: {exc}')
            logger.error('start_inventory error: %s', exc)

    def stop_inventory(self):
        if not self.is_connected() or not self.reading or self._device is None:
            return
        self._intentional_stop = True
        try:
            self._device.StopInventory()
            _dbg('inventory stopped (intentional)')
        except Exception as exc:
            _dbg(f'StopInventory error: {exc}')
        finally:
            self.reading = False

    # ------------------------------------------------------------------
    # Power
    # ------------------------------------------------------------------

    def set_power(self, antenna: int, read_dbm: int, write_dbm: int = 27):
        """Set TX power. AsReader P35U valid range for REGION_US is 13–27 dBm."""
        if not self.is_connected() or self._device is None:
            return
        dbm = max(13, min(27, read_dbm))
        try:
            self._device.SetTxPower(dbm)
            _dbg(f'SetTxPower {dbm} dBm')
            self._on_status(f'Power set to {dbm} dBm')
        except Exception as exc:
            _dbg(f'SetTxPower error: {exc}')
            logger.error('SetTxPower error: %s', exc)

    # ------------------------------------------------------------------
    # Internal — DLL delegate setup and callbacks
    # ------------------------------------------------------------------

    def _setup_delegates(self, AsReaderCls, dev):
        """Wire up the six SetDelegate callbacks.

        pythonnet automatically wraps Python callables into .NET delegates
        when passed to a method that expects a specific delegate type.
        """
        dev.SetDelegate(
            AsReaderCls.CallBackReadTagData(self._cb_tag),
            AsReaderCls.CallBackErrorCode(self._cb_error),
            AsReaderCls.CallBackSuccessCode(self._cb_success),
            AsReaderCls.CallBackCommandData(self._cb_command),
            AsReaderCls.CallBackReadComplete(self._cb_complete),
            AsReaderCls.CallBackTriggerHandler(self._cb_trigger),
        )
        _dbg('SetDelegate registered')

    def _cb_tag(self, result) -> None:
        """Fired by the DLL for every tag read. result is AsReader.InventoryResult."""
        try:
            # If read_tid() registered a one-shot memory-result hook, consume this
            # callback as a ReadMemory response and do NOT process it as an inventory
            # tag.  ReadMemory on FW 1.8.0 fires _cb_tag (not _cb_command).
            mem_cb = self._pending_memory_cb
            if mem_cb is not None:
                self._pending_memory_cb = None
                try:
                    mem_cb(result)
                except Exception as exc:
                    _dbg(f'_cb_tag memory hook error: {exc}')
                return

            td = result.tagdata

            epc_raw = td.epc
            if epc_raw is None:
                _dbg('cb_tag: epc is None, skipping')
                return
            epc = str(epc_raw).strip().upper().replace(' ', '')
            if not epc:
                _dbg('cb_tag: empty EPC after strip, skipping')
                return

            # PC: SDK gives it as a hex string (e.g. "3000"), convert to int
            pc_str = str(td.pc) if td.pc is not None else '0'
            try:
                pc = int(pc_str, 16)
            except ValueError:
                try:
                    pc = int(pc_str)
                except ValueError:
                    pc = 0

            # RSSI: SDK delivers a float string (e.g. '-31.2') or an unsigned
            # byte (0–255 where 128–255 = negative via two's complement).
            try:
                rssi_raw = float(str(result.rssi))
                # Two's complement correction for unsigned-byte encoding
                rssi = int(rssi_raw - 256) if 128 <= rssi_raw <= 255 else int(rssi_raw)
            except Exception as exc:
                _dbg(f'cb_tag: RSSI cast failed ({exc}), defaulting to 0')
                rssi = 0

            # Antenna
            try:
                antenna = int(result.antenna)
            except Exception:
                antenna = 1

            # TID: permanent manufacturer-assigned identifier.
            # SDK may return None for tags that do not expose TID in this mode.
            tid = ''
            try:
                tid_raw = td.tid
                if tid_raw is not None:
                    tid = str(tid_raw).strip().upper().replace(' ', '')
            except Exception as exc:
                _dbg(f'cb_tag: TID read failed ({exc}), leaving empty')

            tag = TagRead(
                raw_epc=epc,
                pc=pc,
                rssi_raw=rssi,
                antenna=antenna,
                tid=tid,
                # freq_khz not available from this SDK
            )
            _dbg(f'TAG epc={tag.raw_epc} tid={tid or "(none)"} rssi={rssi} ant={antenna}')
            self._on_tag(tag)

        except Exception as exc:
            _dbg(f'_cb_tag error: {exc}')
            logger.error('_cb_tag error: %s', exc)

    def _cb_error(self, error) -> None:
        """DLL error callback.

        Only treat as an unexpected disconnect when inventory is actually running
        (self.reading is True).  Errors that arrive while the reader is idle are
        almost always spurious responses to QC commands (StopInventory on an
        already-stopped device, ReadMemory, CheckTagStatus) — log them and let
        the QC result surface the failure instead of tearing down the connection.
        """
        _dbg(f'DLL error callback: {error}')
        if self._connected and not self._intentional_stop:
            if self.reading:
                _dbg('cb_error: inventory running — triggering hardware-disconnect sequence')
                self._on_status(f'Reader error: {error}')
                self._handle_unexpected_disconnect()
            else:
                _dbg(f'cb_error: reader idle (likely QC command) — logging only, no disconnect')

    def _cb_success(self, code) -> None:
        _dbg(f'DLL success callback: {code}')

    def _cb_command(self, data) -> None:
        """DLL fires this with raw response bytes after ReadMemory / other commands."""
        cb = self._pending_command_cb
        _dbg(f'_cb_command: fired (cb_pending={cb is not None}, '
             f'data_type={type(data).__name__}, data_is_none={data is None})')
        if cb is not None:
            self._pending_command_cb = None
            try:
                if data is None:
                    raw = b''
                else:
                    # .NET byte[] → Python bytes.  Try several strategies because
                    # pythonnet version and array type affect which one works.
                    try:
                        raw = bytes(bytearray(data))        # works for most sequences
                    except TypeError:
                        try:
                            raw = bytes(data)               # pythonnet 3.x direct
                        except TypeError:
                            # Last resort: explicit iteration via .NET Length attr
                            raw = bytes(int(data[i]) for i in range(data.Length))
                _dbg(f'_cb_command: {len(raw)} bytes → {raw.hex().upper()!r}')
                cb(raw)
            except Exception as exc:
                _dbg(f'_cb_command invoke error: {exc!r}')

    def _cb_complete(self, complete_status: bool) -> None:
        """SDK fires this when an inventory cycle ends.

        Normal continuous-mode flow:
          True  — one round finished; SDK immediately starts the next.
                  No action needed; self._intentional_stop is False here.
        Intentional stop (user pressed Stop):
          True or False — self._intentional_stop was set by stop_inventory();
                          just reset the flag; reading flag already cleared.
        Unexpected stop (hardware pull / SDK error):
          False — documented convention: inventory terminated unexpectedly.
          True  — some SDK builds invert the convention on hardware pull;
                  distinguished from a normal-round True by self._intentional_stop.
                  However, True also fires after every healthy round, so we ONLY
                  treat a non-intentional True as an unexpected disconnect when
                  self.reading is already False (stop_inventory set it) or when
                  self._connected is False — i.e. the device was already torn down
                  by another path.  In all other non-intentional-True cases we
                  leave the inventory running (normal continuous operation).
        """
        _dbg(f'DLL read complete: complete_status={complete_status} '
             f'intentional={self._intentional_stop} reading={self.reading}')

        if self._intentional_stop:
            # Normal user-initiated stop — reset the guard flag and return.
            _dbg('cb_complete: intentional stop confirmed')
            self._intentional_stop = False
            return

        if self._hw_stop_requested and complete_status:
            # Hardware reached its mtnu limit (e.g. mtnu=1) and stopped itself.
            # This is a clean, expected stop — NOT an unexpected disconnect.
            _dbg('cb_complete: hardware auto-stop (mtnu limit reached)')
            self._hw_stop_requested = False
            self.reading = False
            if self._on_stopped is not None:
                try:
                    self._on_stopped()
                except Exception as exc:
                    _dbg(f'on_stopped callback error: {exc}')
            return

        if not complete_status:
            # Documented path: False means inventory stopped unexpectedly.
            _dbg('cb_complete: unexpected stop (complete_status=False) — '
                 'triggering hardware-disconnect sequence')
            self.reading = False
            self._handle_unexpected_disconnect()
        # complete_status=True without _intentional_stop or _hw_stop_requested:
        # normal continuous-mode round completion — inventory is still running.

    def _handle_unexpected_disconnect(self) -> None:
        """Tear down the connection state and notify the UI layer.

        Safe to call from any thread (DLL callback threads included).
        """
        if not self._connected:
            return  # already disconnected — avoid double-notify
        _dbg('hardware disconnect detected — cleaning up')
        self.reading = False
        self._connected = False
        if self._device is not None:
            try:
                self._device.DisConnect()
            except Exception as exc:
                _dbg(f'DisConnect (unexpected) error: {exc}')
            self._device = None
        self._on_status('Disconnected')
        if self._on_disconnect is not None:
            try:
                self._on_disconnect()
            except Exception as exc:
                _dbg(f'on_disconnect callback error: {exc}')

    def _cb_trigger(self, trigger: int) -> None:
        _dbg(f'DLL trigger: {trigger}')
        if self._on_trigger is not None:
            try:
                self._on_trigger(trigger)
            except Exception as exc:
                _dbg(f'_cb_trigger callback error: {exc}')

    # ------------------------------------------------------------------
    # Advanced settings — applied after connect via main.py
    # ------------------------------------------------------------------

    def get_tx_power_range(self) -> tuple[int, int, int]:
        """Return (current_dbm, min_dbm, max_dbm).  Falls back to (20,13,27)."""
        if not self.is_connected() or self._device is None:
            return (20, 13, 27)
        try:
            result = self._device.GetTxPower(0, 0, 0)
            # pythonnet 3: ref params come back as (ret_code, val1, val2, val3)
            if isinstance(result, tuple) and len(result) >= 4:
                return (int(result[1]), int(result[2]), int(result[3]))
        except Exception as exc:
            _dbg(f'GetTxPower error: {exc}')
        return (20, 13, 27)

    def set_read_time(self, ms: int):
        """Set how long the antenna transmits per inventory cycle (ms)."""
        if not self.is_connected() or self._device is None:
            return
        try:
            self._device.SetReadTime(ms)
            _dbg(f'SetReadTime {ms} ms')
        except Exception as exc:
            _dbg(f'SetReadTime error: {exc}')

    def set_idle_time(self, ms: int):
        """Set pause between inventory cycles (ms)."""
        if not self.is_connected() or self._device is None:
            return
        try:
            self._device.SetIdelTime(ms)
            _dbg(f'SetIdelTime {ms} ms')
        except Exception as exc:
            _dbg(f'SetIdelTime error: {exc}')

    def set_session(self, session: int):
        """session: 0=S0, 1=S1, 2=S2, 3=S3"""
        if not self.is_connected() or self._device is None or self._Types is None:
            return
        try:
            session_map = {
                0: self._Types.SessionType.SESSION_S0,
                1: self._Types.SessionType.SESSION_S1,
                2: self._Types.SessionType.SESSION_S2,
                3: self._Types.SessionType.SESSION_S3,
            }
            s = session_map.get(session, self._Types.SessionType.SESSION_S1)
            ret = self._device.SetSession(s)
            _dbg(f'SetSession S{session} → {ret}')
        except Exception as exc:
            _dbg(f'SetSession error: {exc}')

    def set_anti_collision(self, mode: str, start_q: int, min_q: int, max_q: int):
        """mode: 'DynamicQ' or 'FixedQ'; Q values 0-8."""
        if not self.is_connected() or self._device is None or self._Types is None:
            return
        try:
            T = self._Types
            ac_mode = (T.AntiCollisionMode.DynamicQ
                       if mode == 'DynamicQ'
                       else T.AntiCollisionMode.FixedQ)
            q_vals = [T.QType.Q0, T.QType.Q1, T.QType.Q2, T.QType.Q3,
                      T.QType.Q4, T.QType.Q5, T.QType.Q6, T.QType.Q7, T.QType.Q8]
            def qv(n): return q_vals[max(0, min(8, n))]
            ret = self._device.SetAntiCollisionMode(ac_mode, qv(start_q), qv(min_q), qv(max_q))
            _dbg(f'SetAntiCollisionMode {mode} sq={start_q} min={min_q} max={max_q} → {ret}')
        except Exception as exc:
            _dbg(f'SetAntiCollisionMode error: {exc}')

    def set_rssi_threshold(self, dbm: int):
        """Filter out tags below dbm signal strength.  -99 = accept everything."""
        if not self.is_connected() or self._device is None:
            return
        try:
            val = max(-99, min(0, dbm))
            self._device.SetRSSIThreshold(val)
            _dbg(f'SetRSSIThreshold {val} dBm')
        except Exception as exc:
            _dbg(f'SetRSSIThreshold error: {exc}')

    def set_buzzer(self, on: bool, volume: str = 'HIGH'):
        """Set buzzer level.  on=False → OFF; on=True → volume (LOW or HIGH)."""
        if not self.is_connected() or self._device is None or self._Types is None:
            return
        try:
            if not on:
                bval = self._Types.Buzzer.OFF
            elif volume == 'LOW':
                bval = self._Types.Buzzer.LOW
            else:
                bval = self._Types.Buzzer.HIGH
            self._device.SetBuzzer(bval)
            _dbg(f'SetBuzzer {"OFF" if not on else volume}')
        except Exception as exc:
            _dbg(f'SetBuzzer error: {exc}')

    def set_basic_target(self, mode: str = 'TOGGLE'):
        """Set inventory target mode.

        mode:
          'TARGET_A' — inventory flag A only (good for stop-after-first presets)
          'TARGET_B' — inventory flag B only
          'TOGGLE'   — alternate A↔B each round; tags naturally fall out of
                       repeated singulation without Python-side dedup timers.
                       Use for continuous multi-read presets.

        SDK enum name has a typo: TOGGLE_INVENTORY_ROUBD (not ROUND).
        """
        if not self.is_connected() or self._device is None or self._Types is None:
            return
        try:
            T = self._Types
            target_map = {
                'TARGET_A': T.TargetABType.TARGET_A,
                'TARGET_B': T.TargetABType.TARGET_B,
                'TOGGLE':   T.TargetABType.TOGGLE_INVENTORY_ROUBD,
            }
            tval = target_map.get(mode, T.TargetABType.TOGGLE_INVENTORY_ROUBD)
            ret = self._device.SetBasicTarget(tval)
            _dbg(f'SetBasicTarget {mode} → {ret}')
        except Exception as exc:
            _dbg(f'SetBasicTarget error: {exc}')

    def read_tid(self, epc_hex: str) -> str:
        """Attempt to read TID memory bank for the tag in the RF field.

        Uses the SDK ReadMemory() wrapper.  Two callback paths are tried:

        Path A (_cb_tag hook):  On FW 1.8.0 the DLL fires CallBackReadTagData
        (our _cb_tag) with td.data containing the raw memory bytes.  A one-shot
        _pending_memory_cb hook intercepts that call before it reaches the
        inventory pipeline.

        Path B (_cb_command hook):  Older DLL behaviour; ReadMemory fires
        CallBackCommandData (our _cb_command) with raw BB…7E bytes.  Kept as
        a fallback — has not been observed to fire on SDK 1.3.0 / FW 1.8.0.

        Both hooks share the same threading.Event so whichever fires first wins.

        Returns the TID as an uppercase hex string, or '' on failure / timeout.
        """
        if not self.is_connected() or self._device is None or self._Types is None:
            return ''
        try:
            epc_clean = epc_hex.strip().replace(' ', '')
            epc_bytes = bytes.fromhex(epc_clean)

            evt           = threading.Event()
            result_holder: list[str] = ['']

            # ── Path A: result arrives via _cb_tag (td.data or td.tid) ──────
            def _on_tag_data(inv_result) -> None:
                try:
                    td = inv_result.tagdata
                    # ReadMemory puts the memory contents in td.data;
                    # fall back to td.tid in case DLL maps it there.
                    for attr in ('data', 'tid'):
                        raw = getattr(td, attr, None)
                        if raw is not None:
                            s = str(raw).strip().upper().replace(' ', '')
                            if s:
                                result_holder[0] = s
                                _dbg(f'read_tid _cb_tag path: attr={attr} → {s!r}')
                                break
                except Exception as exc:
                    _dbg(f'read_tid _on_tag_data error: {exc}')
                finally:
                    evt.set()

            # ── Path B: _cb_command fallback (dead code — confirmed by ASR engineering
        #    2026-08-17: CallBackCommandData is reserved for firmware update
        #    packets only; it NEVER fires for ReadMemory() by design, regardless
        #    of firmware version.  Kept as a defensive no-op so evt.set() is
        #    always called if the behaviour ever changes. ────────────────────
            def _on_cmd(data: bytes) -> None:
                raw = bytes(data) if data is not None else b''
                _dbg(f'read_tid _cb_command path (unexpected): {len(raw)}B → {raw.hex().upper()!r}')
                if raw:
                    result_holder[0] = raw.hex().upper()
                evt.set()

            # Stop inventory only if still running (see comment in earlier version).
            if self.reading:
                try:
                    self._device.StopInventory()
                    _dbg('read_tid: StopInventory before ReadMemory')
                except Exception as exc:
                    _dbg(f'read_tid: pre-StopInventory failed ({exc}), continuing')
                time.sleep(0.05)

            # Register both hooks before issuing the command.
            self._pending_memory_cb  = _on_tag_data
            self._pending_command_cb = _on_cmd

            ret = self._device.ReadMemory(
                self._Types.MemBankType.MEM_TID,
                0,   # startAddressWord
                4,   # lengthWord — 4 words = 8 bytes (64-bit TID)
                0,   # accessPassword (none)
                epc_bytes,
            )
            if ret != 0:
                self._pending_memory_cb  = None
                self._pending_command_cb = None
                _dbg(f'read_tid: ReadMemory returned {ret} (non-zero = rejected)')
                return ''

            if evt.wait(timeout=2.0):
                tid = result_holder[0]
                _dbg(f'read_tid {epc_hex[:12]}… → {tid or "(empty)"}')
                return tid
            else:
                self._pending_memory_cb  = None
                self._pending_command_cb = None
                _dbg('read_tid: timeout — neither _cb_tag nor _cb_command fired')
                return ''

        except Exception as exc:
            self._pending_memory_cb  = None
            self._pending_command_cb = None
            _dbg(f'read_tid error: {exc!r}')
            return ''

    _CHECK_TAG_TIMEOUT: float = 3.0   # seconds — matches read_tid timeout

    def check_tag_status(self, epc_hex: str) -> str:
        """Return lock state of the tag currently in the RF field.

        Calls CheckTagStatus(epcData) → Types.TagStatus in a daemon thread
        so that a tag moving out of RF field mid-QC (which causes the DLL to
        block internally for 3-5 s) does not freeze the UI.  If the DLL call
        does not return within _CHECK_TAG_TIMEOUT seconds, 'Timeout' is
        returned immediately and the background thread is abandoned.

        Returns one of:
          'PermaLock' | 'Lock' | 'UnLock' | 'Unknown' | 'Error' | 'Timeout'
        """
        if not self.is_connected() or self._device is None:
            return 'Error'
        try:
            # Convert hex string (e.g. "3034...") to byte array
            epc_clean = epc_hex.strip().replace(' ', '')
            epc_bytes = bytes.fromhex(epc_clean)

            result_holder: list = [None]   # [status_str | Exception]
            evt = threading.Event()

            def _run() -> None:
                try:
                    raw = self._device.CheckTagStatus(epc_bytes)
                    val = int(raw)
                    result_holder[0] = {0: 'UnLock', 1: 'Lock', 2: 'PermaLock',
                                        3: 'Unknown', 4: 'Error'}.get(val, 'Error')
                except Exception as exc:
                    result_holder[0] = exc
                finally:
                    evt.set()

            t = threading.Thread(target=_run, daemon=True,
                                 name='check_tag_status')
            t.start()

            if not evt.wait(timeout=self._CHECK_TAG_TIMEOUT):
                _dbg(f'CheckTagStatus {epc_hex[:12]}… → Timeout '
                     f'(>{self._CHECK_TAG_TIMEOUT:.0f}s — tag left field?)')
                return 'Timeout'

            outcome = result_holder[0]
            if isinstance(outcome, Exception):
                _dbg(f'CheckTagStatus error: {outcome}')
                return 'Error'

            _dbg(f'CheckTagStatus {epc_hex[:12]}… → {outcome}')
            return outcome

        except Exception as exc:
            _dbg(f'CheckTagStatus error: {exc}')
            return 'Error'

    # ── Region / frequency ────────────────────────────────────────────

    _REGION_FREQ_MAP: dict = {
        'REGION_US': '902–928 MHz  (US, 50-ch FHSS)',
        'REGION_EU': '865–868 MHz  (EU)',
        'REGION_KR': '917–920.8 MHz  (KR)',
        'REGION_CN': '920–925 MHz  (CN)',
        'REGION_IN': '865–867 MHz  (IN)',
        'REGION_AU': '920–928 MHz  (AU)',
        'REGION_JP': '916.8–920.8 MHz  (JP)',
    }

    def get_frequency_string(self) -> str:
        """Return a human-readable frequency / region string read live from device."""
        if not self.is_connected() or self._device is None:
            return '—'
        try:
            region = self._device.GetRegion()
            # RegionType enum → name like 'REGION_US'
            name = str(region).split('.')[-1].strip()
            return self._REGION_FREQ_MAP.get(name, str(region))
        except Exception as exc:
            _dbg(f'GetRegion error: {exc}')
            return '902–928 MHz  (US, 50-ch FHSS)'

    # ── Device info ───────────────────────────────────────────────────

    def get_device_info(self) -> dict:
        """Return dict with fw, hw, rfid_fw, sn, sdk, frequency strings."""
        info = {'fw': '—', 'hw': '—', 'rfid_fw': '—',
                'sn': '—', 'sdk': '—', 'frequency': '—'}
        if not self.is_connected() or self._device is None:
            return info
        for key, method_name in [
            ('fw',      'GetFwVersion'),
            ('hw',      'GetHwVersion'),
            ('rfid_fw', 'GetRFIDFwVersion'),
            ('sn',      'GetProductSN'),
            ('sdk',     'GetSdkVersion'),
        ]:
            try:
                result = getattr(self._device, method_name)('')
                # pythonnet 3: ref string → (ret_code, value)
                if isinstance(result, tuple) and len(result) >= 2:
                    info[key] = str(result[1]) or '—'
                elif result is not None:
                    info[key] = str(result) or '—'
            except Exception as exc:
                _dbg(f'{method_name} error: {exc}')
        info['frequency'] = self.get_frequency_string()
        return info

    def factory_reset(self) -> bool:
        """Restore all reader settings to factory defaults."""
        if not self.is_connected() or self._device is None:
            return False
        try:
            result = self._device.DefaultSetting()
            _dbg(f'DefaultSetting → {result}')
            return bool(result)
        except Exception as exc:
            _dbg(f'DefaultSetting error: {exc}')
            return False


# ---------------------------------------------------------------------------
# Backward-compat alias — main.py imports: from reader import E310Reader, TagRead
# ---------------------------------------------------------------------------
E310Reader = P35UReader
