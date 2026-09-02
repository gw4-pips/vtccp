"""
TC Image Scraper - Step 4 (v2): FULLY WIRED, full-frame output
Rev 2.0 - 2026-08-06

Aligned to KNOWLEDGE_BASE.txt Rev 4.1:
  - Section 17: no cropping. Output is the full 1280x960 frame.
  - Section 16: per-model h-flip handled in step3 (v2).
  - Section 4 + FlatField addendum: geometry-gated, clipping-guarded
    flat-field correction handled in step3 (v2).

This file owns: continuous capture + rotate-on-trigger, last-content-frame
selection, retry/recovery, race-fixed health monitor, and unit/calibration
auto-detection. It hands a FULL reconstructed frame plus a detector callback
to step3's review_and_save.

Usage:  python step4_full_pipeline_v2.py
Then run verification scans as normal (PDF-only output config). Each scan
opens a full-frame review; accept or reject in the console. Ctrl+C to stop.
"""

import os
import sys
import re
import struct
import subprocess
import statistics
import threading
import queue
import time
import datetime
import xml.etree.ElementTree as ET

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import numpy as np

from step3_localization_review_v2 import review_and_save

USBPCAP_EXE = r"C:\Program Files\USBPcap\USBPcapCMD.exe"
USBPCAP_DEVICE = r"\\.\USBPcap1"

WATCH_FOLDER = r"C:\CapRawWTCImage"
PCAP_SUBFOLDER = "WS-pcap-Files"

# Calibration / settings sources for auto-detection (KB Section 3/6).
SETTINGS_FOLDER = r"C:\ProgramData\Webscan Inc\TruCheck\Settings"
CALIBRATION_LOG = r"C:\ProgramData\Webscan Inc\TruCheck\CalibrationLog.xml"

FULL_FRAME_INCL_LEN = 1228827
USBPCAP_HEADER_LEN = 27
CONTENT_MEAN_THRESHOLD = 30.0
GLOBAL_HEADER_LEN = 24
PACKET_HEADER_LEN = 16
FRAME_W, FRAME_H = 1280, 960

lock = threading.Lock()
state = {"proc": None, "path": None}
trigger_in_progress = threading.Event()
review_queue = queue.Queue()


# ---------------------------------------------------------------------------
# Unit / calibration / geometry auto-detection
# ---------------------------------------------------------------------------
def _read_latest_calibration_log(board_serial):
    """Return the most recent CalibrationLog entry for a board serial, or {}."""
    try:
        root = ET.parse(CALIBRATION_LOG).getroot()
    except Exception:
        return {}
    matches = [e for e in root.findall("CalibrationLogValue")
               if e.findtext("Serial") == board_serial]
    if not matches:
        return {}
    e = matches[-1]
    def _int(tag):
        try:
            return int(e.findtext(tag))
        except Exception:
            return None
    return {"rmin": _int("RMin"), "rmax": _int("RMax"),
            "rmin_pix": _int("RMinPix"), "rmax_pix": _int("RMaxPix")}


def detect_current_unit_and_calibration(pdf_path=None):
    """Best-effort detection of unit model, board serial, current geometry,
    and calibration anchors. Returns a dict with detection_notes describing
    any gaps. Geometry is read from the per-scan PDF/report when available
    (that is the authoritative source of the geometry actually used)."""
    notes = []
    unit_model = None
    board_serial = None
    geometry = None

    # geometry + unit from the triggering report PDF, if present
    if pdf_path and os.path.isfile(pdf_path):
        try:
            import pypdf
            txt = "".join(p.extract_text() for p in pypdf.PdfReader(pdf_path).pages)
            m = re.search(r"Unit Serial:\s*([A-Za-z0-9\-]+)", txt)
            if m:
                unit_model = m.group(1)
            mg = re.search(r"\b(45Q|30Q|30T|30S|90)\b", txt)
            if mg:
                geometry = mg.group(1)
        except Exception as e:
            notes.append(f"PDF parse failed: {e!r}")

    # board serial from the Settings _s.dat/_z.dat via nrbf parser, if present
    try:
        from nrbf_parser2 import parse_full_stream
        cand = None
        if os.path.isdir(SETTINGS_FOLDER):
            for fn in os.listdir(SETTINGS_FOLDER):
                if fn.endswith("_s.dat"):
                    cand = os.path.join(SETTINGS_FOLDER, fn)
                    break
        if cand:
            data = open(cand, "rb").read()
            parsed = parse_full_stream(data)
            root = parsed["resolved_objects"].get(parsed["root_ref"])
            vals = root["values"] if root else {}
            board_serial = vals.get("_boardSerial") or board_serial
            unit_model = unit_model or vals.get("_serial")
    except Exception as e:
        notes.append(f"settings decode failed: {e!r}")

    calib = _read_latest_calibration_log(board_serial) if board_serial else {}
    if not calib:
        notes.append("no CalibrationLog anchors for this board serial")
    if not geometry:
        notes.append("scan geometry not detected (flat-field gate cannot confirm match)")
    if not unit_model:
        notes.append("unit model not detected (h-flip will default to no-flip)")

    return {
        "unit_model": unit_model,
        "board_serial": board_serial,
        "geometry": geometry,
        "rmin": calib.get("rmin"), "rmax": calib.get("rmax"),
        "rmin_pix": calib.get("rmin_pix"), "rmax_pix": calib.get("rmax_pix"),
        "detection_notes": "; ".join(notes) if notes else "",
    }


# ---------------------------------------------------------------------------
# Capture management
# ---------------------------------------------------------------------------
def start_new_capture():
    pcap_dir = os.path.join(WATCH_FOLDER, PCAP_SUBFOLDER)
    os.makedirs(pcap_dir, exist_ok=True)
    path = os.path.join(
        pcap_dir, f"_tc_active_capture_{datetime.datetime.now():%Y%m%d_%H%M%S_%f}.pcap")
    cmd = [USBPCAP_EXE, "-d", USBPCAP_DEVICE, "-o", path,
           "-b", "16777216", "-A", "-s", "2000000"]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return proc, path


def start_new_capture_with_retry(max_attempts=6, initial_delay=0.15):
    delay = initial_delay
    for attempt in range(1, max_attempts + 1):
        proc, path = start_new_capture()
        time.sleep(delay)
        if proc.poll() is None:
            if attempt > 1:
                print(f"  Capture launch succeeded on attempt {attempt}.")
            return proc, path
        err = proc.stderr.read().decode(errors="replace")
        print(f"  Capture launch attempt {attempt}/{max_attempts} died "
              f"(rc={proc.returncode}). stderr: {err.strip()!r}")
        delay *= 2
    print("  ALERT: capture could not be (re)started. Capture is DOWN.")
    return None, None


def parse_full_frames(path):
    frames = []
    with open(path, "rb") as f:
        data = f.read()
    offset = GLOBAL_HEADER_LEN
    while offset + PACKET_HEADER_LEN <= len(data):
        _, _, incl_len, _ = struct.unpack("<IIII", data[offset:offset + PACKET_HEADER_LEN])
        offset += PACKET_HEADER_LEN
        pkt = data[offset:offset + incl_len]
        offset += incl_len
        if incl_len == FULL_FRAME_INCL_LEN:
            payload = pkt[USBPCAP_HEADER_LEN:]
            if len(payload) == FRAME_W * FRAME_H:
                frames.append((statistics.fmean(payload), payload))
    return frames


def handle_trigger(pdf_path):
    ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
    print(f"[{ts}] TRIGGER: {os.path.basename(pdf_path)}")
    trigger_in_progress.set()
    try:
        _handle_trigger_inner(pdf_path)
    finally:
        trigger_in_progress.clear()


def _handle_trigger_inner(pdf_path):
    with lock:
        old_proc, old_path = state["proc"], state["path"]

    if old_proc.poll() is not None:
        print(f"  WARNING: old capture had already exited (rc={old_proc.returncode}).")
    t0 = time.time()
    old_proc.terminate()
    old_proc.wait(timeout=5)
    print(f"  Old capture closed in {time.time()-t0:.3f}s (rc={old_proc.returncode}).")

    new_proc, new_path = start_new_capture_with_retry()
    with lock:
        state["proc"], state["path"] = new_proc, new_path
    if new_proc:
        print(f"  New capture running: {new_path}")

    frames = parse_full_frames(old_path)
    print(f"  Parsed {len(frames)} full frame(s).")
    content = [(i, m, p) for i, (m, p) in enumerate(frames) if m >= CONTENT_MEAN_THRESHOLD]
    if not content:
        print("  NO CONTENT FRAME FOUND. Flagging - not silently skipping.")
        return
    last_idx, last_mean, last_payload = content[-1]
    print(f"  Using frame {last_idx} (mean={last_mean:.2f}) as the verified image.")

    # De-serialize row-major, no flip here (KB Section 9). Display-orientation
    # flip is applied downstream per-model in step3 (Section 16).
    reconstructed = np.frombuffer(last_payload, dtype=np.uint8).reshape(FRAME_H, FRAME_W)
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    print(f"  Queued for review: {stem}")
    review_queue.put((reconstructed, pdf_path, stem))


class PdfHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".pdf"):
            return
        threading.Thread(target=handle_trigger, args=(event.src_path,), daemon=True).start()


def health_monitor():
    known = None
    while True:
        with lock:
            current = state["proc"]
        if current is not None and current is not known:
            known = current
        if known is not None and known.poll() is not None:
            if trigger_in_progress.is_set():
                time.sleep(0.5); continue
            ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
            print(f"[{ts}] HEALTH: capture exited (rc={known.returncode}). Recovering...")
            new_proc, new_path = start_new_capture_with_retry()
            with lock:
                if state["proc"] is current:
                    state["proc"], state["path"] = new_proc, new_path
            if new_proc:
                print(f"  RECOVERED: {new_path}")
            known = new_proc
        time.sleep(0.5)


def main():
    print("=" * 70)
    print("STEP 4 v2 - FULL-FRAME PIPELINE (no cropping; per-model flip; gated FF)")
    print("=" * 70)
    if not os.path.isdir(WATCH_FOLDER):
        print(f"ERROR: WATCH_FOLDER does not exist: {WATCH_FOLDER}")
        sys.exit(1)

    print(f"Watching: {WATCH_FOLDER}")
    proc, path = start_new_capture_with_retry()
    state["proc"], state["path"] = proc, path
    time.sleep(0.5)
    print(f"Initial capture running: {path}")

    observer = Observer()
    observer.schedule(PdfHandler(), WATCH_FOLDER, recursive=False)
    observer.start()
    threading.Thread(target=health_monitor, daemon=True).start()

    print("-" * 70)
    print("Ready. Run verification scans (PDF-only output). Full-frame review")
    print("opens here on the main thread per scan. Ctrl+C to stop.")
    print("-" * 70)

    try:
        while True:
            try:
                reconstructed, pdf_path, stem = review_queue.get(timeout=1)
            except queue.Empty:
                continue
            review_and_save(
                reconstructed, pdf_path, stem, WATCH_FOLDER,
                detect_fn=lambda p=pdf_path: detect_current_unit_and_calibration(p))
    except KeyboardInterrupt:
        print("Stopping...")
        observer.stop(); observer.join()
        with lock:
            if state["proc"]:
                state["proc"].terminate()
        print("Done.")


if __name__ == "__main__":
    main()
