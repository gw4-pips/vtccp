"""
TC Image Scraper - Step 3 (v2): full-frame review + save/reject
Rev 2.0 - 2026-08-06

Aligned to KNOWLEDGE_BASE.txt Rev 4.1:
  - Section 17: cropping / symbol-localization RETIRED for all FlexHite units.
    Output is ALWAYS the full 1280x960 frame. No ROI selection, no crop, no
    localization, no crop-boundary review UI. The name "localization_review"
    is kept only for import compatibility with step4; there is no localization.
  - Section 16: per-model h-flip. TC-850 / TC-853 (predicted) flip; TC-861 /
    TC-863 do not. Unknown models default to NO flip WITH a loud warning.
  - Section 4 + FlatField addendum (2026-08-06): flat-field correction is
    per-board-serial, applied FLIP-THEN-CORRECT, geometry-gated (45Q by
    current inference) and clipping-guarded (refuse + fall back to raw if
    correction would clip too much of the frame).

Operator flow per scan: a full-frame review image opens; operator accepts
(save) or rejects. No crop step exists.
"""

import os
import re
import numpy as np
import cv2
from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# Config / constants
# ---------------------------------------------------------------------------
FRAME_W, FRAME_H = 1280, 960
STRIP_HEIGHT = 90                 # label strip appended below the full frame
MAX_DISPLAY_W, MAX_DISPLAY_H = 1400, 1100
OCRB_FONT_PATH = "OCRB.ttf"

REJECTED_SUBFOLDER = "Rejected Scans"
RAW_SUBFOLDER = "RAW Images-NOT Shared"
SAVE_RAW_DEV_COPY = True          # dev: also save the uncorrected raw frame
                                  # (same orientation, same strip) when a
                                  # corrected image is the shared output.
                                  # Turn OFF for beta/production builds.

FLATFIELD_MAP_FOLDER = r"C:\CapRawWTCImage\FlatFieldMaps"
EXPECTED_MAP_SHAPE = (FRAME_H, FRAME_W)   # (960, 1280)

# Geometry the shading map is calibrated for (KB FlatField addendum Finding A;
# strong inference, pending designer confirmation). Correction is applied only
# when the detected scan geometry matches this. If detection is unavailable,
# the clipping guard is the backstop.
MAP_CALIBRATION_GEOMETRY = "45Q"

# Refuse correction if more than this fraction of the frame would clip to 255
# (KB FlatField addendum Finding B). 0.10% ~= 1230 px of a 1.23 Mpx frame.
CLIP_FRACTION_LIMIT = 0.0010

# --- Per-model h-flip (KB Rev 4.1 Section 16) -------------------------------
# True  -> raw wire image is mirrored vs TruCheck display; apply h-flip.
# False -> raw already matches display orientation; do NOT flip.
# None  -> present-but-unconfirmed; default to no-flip with a loud warning.
MODEL_HFLIP = {
    "TC-850": True,     # confirmed 2026-07-18 (FlyCap2 matched check)
    "TC-853": True,     # PREDICTED per user domain authority; not yet measured
    "TC-861": False,    # confirmed 2026-08-06 (corr 0.953 as-is vs 0.289 flipped)
    "TC-863": False,    # confirmed 2026-08-06 (matched report vs FlyCap2)
}

# --- Known FlexHite px/mil (KB Section 8.2) - still used for the label strip
FLEXHITE_PX_PER_MIL = {
    "TC-850": 0.9563,
    "TC-863": 0.9563,
    "TC-861": 1.6257,
    "TC-853": 1.6257,
}


def get_model_hflip(unit_model):
    """Return (needs_flip, note) for a unit model string, matched by prefix.
    Unconfirmed/unknown models return False plus a LOUD note the caller must
    surface. Never a silent guess."""
    if not unit_model:
        return False, ("UNIT MODEL UNKNOWN - cannot determine sensor mirror. "
                       "Defaulting to NO FLIP. VERIFY ORIENTATION BY EYE: a Data "
                       "Matrix solid-L should sit on the LEFT and BOTTOM edges.")
    um = unit_model.upper()
    for prefix, value in MODEL_HFLIP.items():
        if um.startswith(prefix):
            if prefix == "TC-853":
                return value, ("TC-853 h-flip is PREDICTED (Rule 6), not yet "
                               "directly measured. Applying flip per prediction; "
                               "verify by eye when possible.")
            return value, None
    return False, (f"Model {unit_model} is NOT in the h-flip table. Defaulting to "
                   f"NO FLIP. VERIFY ORIENTATION BY EYE before saving.")


def get_known_px_per_mil(unit_model):
    if not unit_model:
        return None
    um = unit_model.upper()
    for prefix, value in FLEXHITE_PX_PER_MIL.items():
        if um.startswith(prefix):
            return value
    return None


# ---------------------------------------------------------------------------
# Flat-field gain map
# ---------------------------------------------------------------------------
def get_flatfield_gain_map(board_serial):
    """Load this unit's Advanced Calibration gain map by BOARD SERIAL.
    Returns a (960,1280) uint16 array or None. Never raises; every rejection
    path prints why, so 'correction unavailable' is always explainable."""
    if not board_serial:
        print("  Flat-field: no board serial detected - cannot look up a map.")
        return None
    path = os.path.join(FLATFIELD_MAP_FOLDER, f"{board_serial}.npy")
    if not os.path.isfile(path):
        print(f"  Flat-field: no map on file for board serial {board_serial}")
        print(f"              (looked for: {path})")
        return None
    try:
        gain_map = np.load(path)
    except Exception as e:
        print(f"  Flat-field: FAILED to load {path}: {e!r}")
        return None
    if gain_map.shape != EXPECTED_MAP_SHAPE:
        print(f"  Flat-field: REJECTED {path} - shape {gain_map.shape}, "
              f"expected {EXPECTED_MAP_SHAPE}.")
        return None
    if gain_map.dtype != np.uint16:
        print(f"  Flat-field: REJECTED {path} - dtype {gain_map.dtype}, expected uint16.")
        return None
    lo, hi = int(gain_map.min()), int(gain_map.max())
    centre = int(gain_map[EXPECTED_MAP_SHAPE[0] // 2, EXPECTED_MAP_SHAPE[1] // 2])
    if lo < 200 or hi > 2000 or not (200 <= centre <= 400):
        print(f"  Flat-field: REJECTED {path} - implausible values "
              f"(min={lo}, max={hi}, centre={centre}; expect min~255, centre~256).")
        return None
    print(f"  Flat-field: loaded map for {board_serial} "
          f"(min={lo}, max={hi}, centre={centre}, max boost {hi/256.0:.2f}x)")
    return gain_map


def apply_flatfield_correction(display_gray, gain_map):
    """Apply the gain map to a full frame that is ALREADY in display
    orientation (flip first - KB FlatField addendum Finding C). Returns a new
    uint8 array, same shape."""
    corrected = display_gray.astype(np.float32) * (gain_map.astype(np.float32) / 256.0)
    return np.clip(corrected, 0, 255).astype(np.uint8)


def would_clip_fraction(display_gray, gain_map):
    """Fraction of pixels that would hit 255 after correction, without
    committing to the corrected array."""
    corrected = display_gray.astype(np.float32) * (gain_map.astype(np.float32) / 256.0)
    return float((corrected >= 255.0).mean())


# ---------------------------------------------------------------------------
# Label strip (full-frame; KB Section 17 - fixed canvas extension)
# ---------------------------------------------------------------------------
def _load_font(size=22):
    try:
        return ImageFont.truetype(OCRB_FONT_PATH, size)
    except Exception:
        return ImageFont.load_default()


def _split_stem(stem):
    """Best-effort split of '{name}_{date}_{time}' into (name, date-time)."""
    m = re.match(r"(.*?)[_ ](\d{2}[-_]\d{2}[-_]\d{2}.*)$", stem)
    if m:
        return m.group(1), m.group(2)
    return stem, ""


def compose_with_strip(full_gray, stem, calib_line=None, name_suffix=""):
    """Append a 3-line label strip below the FULL 1280x960 frame.
      line 1: name portion (+ name_suffix, e.g. '_RAW')
      line 2: date/time portion
      line 3: px/mil + calibration anchors (or placeholder)"""
    img = Image.fromarray(full_gray)
    w, h = img.size
    canvas = Image.new('L', (w, h + STRIP_HEIGHT), color=255)
    canvas.paste(img, (0, 0))
    font = _load_font()
    name_part, date_time_part = _split_stem(stem)
    if name_suffix:
        name_part = f"{name_part}{name_suffix}"
    line3 = calib_line if calib_line else "(px/mil, calibration - pending)"
    d = ImageDraw.Draw(canvas)
    d.text((6, h + 4),  name_part,      fill=0, font=font)
    d.text((6, h + 32), date_time_part, fill=0, font=font)
    d.text((6, h + 60), line3,          fill=0, font=font)
    return canvas


# ---------------------------------------------------------------------------
# Console/window helpers (Windows). Kept minimal; safe no-ops off-Windows.
# ---------------------------------------------------------------------------
def bring_console_to_front():
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW("TC Image Scraper - review")
    except Exception:
        pass


def send_window_to_back(_name):
    pass  # placeholder; OpenCV window stacking handled by the WM


def handle_reject(pdf_path, watch_folder):
    """Move the triggering PDF into the Rejected subfolder; save no image."""
    try:
        rej = os.path.join(watch_folder, REJECTED_SUBFOLDER)
        os.makedirs(rej, exist_ok=True)
        base = os.path.basename(pdf_path)
        os.replace(pdf_path, os.path.join(rej, base))
        print(f"  REJECTED. Moved report to {REJECTED_SUBFOLDER}\\{base}")
    except Exception as e:
        print(f"  REJECTED (note: could not move PDF: {e!r})")


# ---------------------------------------------------------------------------
# Main review entry point
# ---------------------------------------------------------------------------
def review_and_save(reconstructed_gray, pdf_path, stem, watch_folder,
                    detect_fn=None):
    """Full-frame review + save/reject. No cropping (KB Section 17).

    Pipeline (KB Section 16 + FlatField addendum):
      1. Per-model h-flip to display orientation (if required).
      2. Flat-field correction in display orientation, IF: a map exists for
         this board serial AND the scan geometry matches the map's
         calibration geometry AND correction would not clip too much.
         Otherwise the (flipped) raw frame is the output.
      3. Compose label strip below the full frame; operator accept/reject.

    detect_fn() -> dict with keys: unit_model, board_serial, geometry,
      rmin, rmin_pix, rmax, rmax_pix, detection_notes. Injected by step4 so
      this module has no hard dependency on the detector's internals.
    """
    detected = detect_fn() if detect_fn else {}
    if detected.get("detection_notes"):
        print(f"  Auto-detection notes: {detected['detection_notes']}")
    unit_model  = detected.get("unit_model")
    board_serial = detected.get("board_serial")
    geometry    = detected.get("geometry")
    rmin, rmin_pix = detected.get("rmin"), detected.get("rmin_pix")
    rmax, rmax_pix = detected.get("rmax"), detected.get("rmax_pix")

    # --- STEP 1: per-model sensor mirror -> display orientation -------------
    needs_flip, flip_note = get_model_hflip(unit_model)
    if flip_note:
        print(f"  *** ORIENTATION: {flip_note}")
    if needs_flip:
        oriented = np.fliplr(reconstructed_gray)
        print(f"  H-FLIP applied for model {unit_model}.")
    else:
        oriented = reconstructed_gray
        print(f"  No flip for model {unit_model} (raw matches display).")

    # --- STEP 2: flat-field correction in display orientation --------------
    gain_map = get_flatfield_gain_map(board_serial)
    corrected_applied = False
    display_frame = oriented
    if gain_map is not None:
        # geometry gate
        if geometry and geometry.upper() != MAP_CALIBRATION_GEOMETRY.upper():
            print(f"  Flat-field: SKIPPED - scan geometry {geometry} != map "
                  f"calibration geometry {MAP_CALIBRATION_GEOMETRY}. Using raw. "
                  f"(One map does not correct all geometries - KB FlatField A.)")
        else:
            clip_frac = would_clip_fraction(oriented, gain_map)
            if clip_frac > CLIP_FRACTION_LIMIT:
                print(f"  Flat-field: REFUSED - correction would clip "
                      f"{clip_frac*100:.2f}% of the frame at 255 (> "
                      f"{CLIP_FRACTION_LIMIT*100:.2f}% limit). Using raw to avoid "
                      f"silent data loss (KB FlatField B).")
            else:
                display_frame = apply_flatfield_correction(oriented, gain_map)
                corrected_applied = True
                geo_txt = geometry if geometry else "geometry unknown - guard-passed"
                print(f"  Flat-field correction APPLIED ({geo_txt}; "
                      f"clip {clip_frac*100:.3f}%).")
    # if no map, display_frame is just the oriented raw

    # --- label strip line 3 (px/mil + calibration anchors) -----------------
    ppm = get_known_px_per_mil(unit_model)
    ppm_txt = f"{ppm:.4f}px/mil" if ppm else "px/mil n/a"
    if None not in (rmin, rmin_pix, rmax, rmax_pix):
        calib_line = f"{ppm_txt}  RMin:{rmin}@raw{rmin_pix} RMax:{rmax}@raw{rmax_pix}"
    else:
        calib_line = f"{ppm_txt}  RMin/RMax: not available this session"
    if corrected_applied:
        calib_line += "  [FF]"   # discrete marker: correction applied

    composited = compose_with_strip(display_frame, stem, calib_line=calib_line)

    # --- STEP 3: full-frame operator review --------------------------------
    preview = np.array(composited)
    ph, pw = preview.shape[:2]
    scale = min(MAX_DISPLAY_W / pw, MAX_DISPLAY_H / ph)   # may enlarge or shrink
    disp = cv2.resize(preview, (int(pw * scale), int(ph * scale)))
    win = "Full-frame review - image + ID strip as it will be saved. See console."
    cv2.namedWindow(win, cv2.WINDOW_NORMAL)
    cv2.resizeWindow(win, int(pw * scale), int(ph * scale))
    cv2.imshow(win, disp)
    cv2.waitKey(1)
    send_window_to_back(win)
    bring_console_to_front()

    resp = input("  [s]ave / [x] reject: ").strip().lower()
    cv2.destroyWindow(win)
    cv2.waitKey(1)

    if resp not in ("s", "save"):
        handle_reject(pdf_path, watch_folder)
        return

    out_path = os.path.join(watch_folder, f"{stem}.tiff")
    composited.save(out_path, format='TIFF', compression=None)
    print(f"  ACCEPTED. Saved: {out_path}")

    # dev: save the uncorrected (but same-orientation) raw counterpart
    if SAVE_RAW_DEV_COPY and corrected_applied:
        raw_dir = os.path.join(watch_folder, RAW_SUBFOLDER)
        os.makedirs(raw_dir, exist_ok=True)
        raw_strip = compose_with_strip(oriented, stem, calib_line=calib_line,
                                       name_suffix="_RAW")
        raw_path = os.path.join(raw_dir, f"{stem}_RAW.tiff")
        raw_strip.save(raw_path, format='TIFF', compression=None)
        print(f"  DEV: raw copy saved: {raw_path}")
