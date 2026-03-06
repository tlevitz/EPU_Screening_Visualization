# app.py

import base64
import io
import re
import os
import threading
from collections import OrderedDict

from datetime import datetime
from functools import lru_cache

from flask import Flask, render_template, send_file, abort, jsonify, request, url_for
from PIL import Image

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module=r"numpy\._core\.getlimits")

from epu.epu_stats import (
    load_calibration_table,
    process_directory_screening,
    process_directory_collection,
)
from session_layout import (
    detect_atlas_root,
    find_latest_atlas_jpg,
    find_fallback_atlas_jpgs,
    build_session_nodes,
)
from epu.annotate_atlas import annotate_atlas_pair
from epu.annotate_foilhole import annotate_foilhole_template
from epu.annotate_gridsquare import annotate_gridsquare_image_or_pair, annotate_single_gridsquare_image, _parse_ts_yyyymmdd_hhmmss, find_latest_gridsquare_support_and_nonsupport
from epu.report_scale_bars import add_scale_bar_by_xml
from epu.report_style import FONT_SIZES

# ------------------

### CHANGE THIS IF YOU ARE NOT LOCATED AT /mnt/z ###

BASE_ROOT = "/mnt/z"
PIXEL_TABLE_PATH = os.path.join(BASE_ROOT, "pixelsizes.txt")

### ----------------


app = Flask(__name__)

# Cache pixel table and annotated atlas / template / gridsquares in memory
_pix_dict = None
_beamsize_dict = None
_caldate_dict = None

class ImageCache:
    """
    Simple thread-safe LRU cache for image bytes.
    Limits by number of entries.
    """
    def __init__(self, max_entries=256):
        self.max_entries = max_entries
        self._lock = threading.Lock()
        self._data = OrderedDict()  # key -> bytes

    def get(self, key):
        with self._lock:
            if key not in self._data:
                return None
            self._data.move_to_end(key)  # mark as recently used
            return self._data[key]

    def set(self, key, value):
        with self._lock:
            if key in self._data:
                self._data.move_to_end(key)
            self._data[key] = value
            while len(self._data) > self.max_entries:
                self._data.popitem(last=False)  # evict oldest

    def clear(self):
        with self._lock:
            self._data.clear()

image_cache = ImageCache(max_entries=256)

def get_cached_image_bytes(cache_key, generator_func, fmt="JPEG"):
    """
    cache_key: hashable key (e.g., tuple)
    generator_func: function that returns a PIL.Image or None
    Returns: bytes or None
    """
    data = image_cache.get(cache_key)
    if data is not None:
        return data


    img = generator_func()
    if img is None:
        return None


    buf = io.BytesIO()
    img.save(buf, format=fmt)
    data = buf.getvalue()
    image_cache.set(cache_key, data)
    return data

def get_pixel_table():
    global _pix_dict, _beamsize_dict, _caldate_dict
    if _pix_dict is None:
        _pix_dict, _beamsize_dict, _caldate_dict = load_calibration_table(PIXEL_TABLE_PATH)
    return _pix_dict, _beamsize_dict, _caldate_dict

def encode_path(path: str) -> str:
    return base64.urlsafe_b64encode(path.encode("utf-8")).decode("ascii")

def decode_path(token: str) -> str:
    try:
        return base64.urlsafe_b64decode(token.encode("ascii")).decode("utf-8")
    except Exception:
        raise ValueError("Invalid session id")

def is_collection_session(session_dir: str) -> bool:
    """
    Auto-detect mode: if any *Fractions files exist under Images-Disc1, treat as collection.
    """
    images_root = os.path.join(session_dir, "Images-Disc1")
    if not os.path.isdir(images_root):
        return False
    for root, dirs, files in os.walk(images_root):
        for f in files:
            if "fractions" in f.lower():
                return True
    return False

def find_sessions():
    sessions = []
    if not os.path.isdir(BASE_ROOT):
        return sessions
    for entry in os.scandir(BASE_ROOT):
        if not entry.is_dir():
            continue
        images_disc1 = os.path.join(entry.path, "Images-Disc1")
        if os.path.isdir(images_disc1):
            mtime = datetime.fromtimestamp(entry.stat().st_mtime)
            sessions.append(
                {
                    "name": entry.name,
                    "path": entry.path,
                    "mtime": mtime,
                    "id": encode_path(entry.path),
                }
            )
    sessions.sort(key=lambda s: s["mtime"], reverse=True)
    return sessions

def _format_defocus_values(val):
    """
    Expect val like [[-1.0, -1.5], [-2.0, -2.5]].
    Convert each inner list to its string form, then join with ", ".
    """
    if not isinstance(val, list):
        return val
    inner_strs = [str(inner) for inner in val]
    return ", ".join(inner_strs)

def build_summary_rows(df_all, instrument_model, mode):
    """
    Mirror epu_stats.write_table:
    - Use the same column lists and order.
    - Apply Defocus Values (um) formatting.
    """

    row = df_all.iloc[0].copy()
    camera_string = str(row.get("Camera", "")).strip()

    if mode == "screening":
        if instrument_model and "TUNDRA" in instrument_model.upper() and camera_string == "Ceta-F":
            cols = [
                "Date", "Folder", "Atlas Path", "Start Time", "End Time", "Total Time (hrs)",
                "Grid Squares Screened", "Total Micrographs",
                "Average Micrographs per Grid Square",
                "Microscope", "Acceleration Voltage (kV)", "Extractor Voltage (V)",
                "Spherical Aberration (mm)", "Gun Lens", "Spot Size", "Intensity",
                "EPU Version", "C2 Aperture (um)", "Objective Aperture (um)",
                "Camera", "Image Dimensions (pixels)", "Nominal Magnification",
                "EPU Pixel Size (A/pix)", "Calibrated Pixel Size (A/pix)",
                "Calibrated Beam Diameter (um)", "Pixel and Beam Size Calibration Date",
                "Exposure Time (s)", "Approx. Total Dose (e/pix)",
                "Approx. Total Dose (e/A2)", "Approx. Dose Rate (e/pix/s)",
                "Grid Type", "Grid Geometry", "EPU Measured Hole Size (um)",
                "EPU Measured Hole Center-to-Center Distance (um)",
                "Best Guess Hole Size and Spacing (um)",
                "Number of Acquisition Areas (Shots Per Hole)",
                "AFIS", "AFIS Clustering Distance (um)",
                "Number of Fractions", "Defocus Values (um)",
            ]
        elif instrument_model and "TUNDRA" in instrument_model.upper():
            cols = [
                "Date", "Folder", "Atlas Path", "Gain Reference File", "Start Time", "End Time", 
                "Total Time (hrs)", "Grid Squares Screened", "Total Micrographs",
                "Average Micrographs per Grid Square",
                "Microscope", "Acceleration Voltage (kV)", "Extractor Voltage (V)",
                "Spherical Aberration (mm)", "Gun Lens", "Spot Size", "Intensity",
                "EPU Version", "C2 Aperture (um)", "Objective Aperture (um)",
                "Camera", "Camera Mode", "Image Dimensions (pixels)", "Nominal Magnification",
                "EPU Pixel Size (A/pix)", "Calibrated Pixel Size (A/pix)",
                "Calibrated Beam Diameter (um)", "Pixel and Beam Size Calibration Date",
                "Exposure Time (s)", "Approx. Total Dose (e/pix)",
                "Approx. Total Dose (e/A2)", "Approx. Dose Rate (e/pix/s)",
                "Grid Type", "Grid Geometry", "EPU Measured Hole Size (um)",
                "EPU Measured Hole Center-to-Center Distance (um)",
                "Best Guess Hole Size and Spacing (um)",
                "Number of Acquisition Areas (Shots Per Hole)",
                "AFIS", "AFIS Clustering Distance (um)",
                "Number of Fractions", "Defocus Values (um)",
            ]
        else:
            cols = [
               "Date", "Folder", "Atlas Path", "Start Time", "End Time", "Total Time (hrs)",
                "Grid Squares Screened", "Total Micrographs",
                "Average Micrographs per Grid Square", "Gain Reference File",
                "EPU Version", "Start Time", "End Time", "Total Time (hrs)",
                "Grid Squares Collected", "Total Movies",
                "Average Movies per Grid Square", "Movies per Hour",
                "Stage Tilt (Degrees)", "Microscope",
                "Acceleration Voltage (kV)", "Extractor Voltage (V)",
                "Spherical Aberration (mm)", "Gun Lens", "Spot Size",
                "Beam Diameter (um)", "C2 Aperture (um)", "C3 Aperture (um)",
                "Objective Aperture (um)", "Energy Filter",
                "Energy Filter Slit Width (eV)", "Illumination Mode",
                "Camera", "Camera Mode", "Image Dimensions (pixels)",
                "Nominal Magnification", "EPU Pixel Size (A/pix)",
                "Calibrated Pixel Size (A/pix)", "Pixel Size Calibration Date",
                "Exposure Time (s)", "Approx. Total Dose (e/pix)",
                "Approx. Total Dose (e/A2)", "Approx. Dose Rate (e/pix/s)",
                "Grid Type", "Grid Geometry", "EPU Measured Hole Size (um)",
                "EPU Measured Hole Center-to-Center Distance (um)",
                "Best Guess Hole Size and Spacing (um)",
                "Number of Acquisition Areas (Shots Per Hole)",
                "AFIS", "AFIS Clustering Distance (um)",
                "Number of Fractions", "Defocus Values (um)",
            ]
    else:
        if instrument_model and "TUNDRA" in instrument_model.upper() and camera_string == "Ceta-F":
            cols = [
                "Date", "Folder", "Atlas Path", "Start Time", "End Time", "Total Time (hrs)",
                "Grid Squares Collected", "Total Movies",
                "Average Movies per Grid Square", "Movies per Hour",
                "Microscope", "Acceleration Voltage (kV)", "Extractor Voltage (V)",
                "Spherical Aberration (mm)", "Gun Lens", "Spot Size", "Intensity",
                "EPU Version", "C2 Aperture (um)", "Objective Aperture (um)",
                "Camera", "Image Dimensions (pixels)", "Nominal Magnification",
                "EPU Pixel Size (A/pix)", "Calibrated Pixel Size (A/pix)",
                "Beam Size (um)", "Pixel and Beam Size Calibration Date",
                "Exposure Time (s)", "Approx. Total Dose (e/pix)",
                "Approx. Total Dose (e/A2)", "Approx. Dose Rate (e/pix/s)",
                "Grid Type", "Grid Geometry", "EPU Measured Hole Size (um)",
                "EPU Measured Hole Center-to-Center Distance (um)",
                "Best Guess Hole Size and Spacing (um)",
                "Number of Acquisition Areas (Shots Per Hole)",
                "AFIS", "AFIS Clustering Distance (um)",
                "Number of Fractions", "Defocus Values (um)",
            ]
        elif instrument_model and "TUNDRA" in instrument_model.upper():
            cols = [
                "Date", "Folder", "Atlas Path", "Gain Reference File", "Start Time", "End Time", 
                "Total Time (hrs)", "Grid Squares Collected", "Total Movies",
                "Average Movies per Grid Square", "Movies per Hour",
                "Microscope", "Acceleration Voltage (kV)", "Extractor Voltage (V)",
                "Spherical Aberration (mm)", "Gun Lens", "Spot Size", "Intensity",
                "EPU Version", "C2 Aperture (um)", "Objective Aperture (um)",
                "Camera", "Camera Mode", "Image Dimensions (pixels)", "Nominal Magnification",
                "EPU Pixel Size (A/pix)", "Calibrated Pixel Size (A/pix)",
                "Beam Size (um)", "Pixel and Beam Size Calibration Date",
                "Exposure Time (s)", "Approx. Total Dose (e/pix)",
                "Approx. Total Dose (e/A2)", "Approx. Dose Rate (e/pix/s)",
                "Grid Type", "Grid Geometry", "EPU Measured Hole Size (um)",
                "EPU Measured Hole Center-to-Center Distance (um)",
                "Best Guess Hole Size and Spacing (um)",
                "Number of Acquisition Areas (Shots Per Hole)",
                "AFIS", "AFIS Clustering Distance (um)",
                "Number of Fractions", "Defocus Values (um)",
            ]
        else:
            cols = [
                "Date", "Folder", "Atlas Path", "Gain Reference File",
                "EPU Version", "Start Time", "End Time", "Total Time (hrs)",
                "Grid Squares Collected", "Total Movies",
                "Average Movies per Grid Square", "Movies per Hour",
                "Stage Tilt (Degrees)", "Microscope",
                "Acceleration Voltage (kV)", "Extractor Voltage (V)",
                "Spherical Aberration (mm)", "Gun Lens", "Spot Size",
                "Beam Diameter (um)", "C2 Aperture (um)", "C3 Aperture (um)",
                "Objective Aperture (um)", "Energy Filter",
                "Energy Filter Slit Width (eV)", "Illumination Mode",
                "Camera", "Camera Mode", "Image Dimensions (pixels)",
                "Nominal Magnification", "EPU Pixel Size (A/pix)",
                "Calibrated Pixel Size (A/pix)", "Pixel Size Calibration Date",
                "Exposure Time (s)", "Approx. Total Dose (e/pix)",
                "Approx. Total Dose (e/A2)", "Approx. Dose Rate (e/pix/s)",
                "Grid Type", "Grid Geometry", "EPU Measured Hole Size (um)",
                "EPU Measured Hole Center-to-Center Distance (um)",
                "Best Guess Hole Size and Spacing (um)",
                "Number of Acquisition Areas (Shots Per Hole)",
                "AFIS", "AFIS Clustering Distance (um)",
                "Number of Fractions", "Defocus Values (um)",
            ]

    # Keep only columns that actually exist
    cols = [c for c in cols if c in df_all.columns]

    # Apply Defocus Values (um) formatting if present
    if "Defocus Values (um)" in row.index:
        row["Defocus Values (um)"] = _format_defocus_values(row["Defocus Values (um)"])

    summary_rows = [(c, str(row[c])) for c in cols]
    return summary_rows

def get_session_version(session_dir: str) -> float:
    """
    Return a version token for this session based on the newest file
    mtime under the session directory.

    Any new image/XML/etc. written anywhere inside the session folder
    will bump this version and invalidate cached stats/nodes/images.
    """
    latest = 0.0
    try:
        for root, dirs, files in os.walk(session_dir):
            for name in files:
                path = os.path.join(root, name)
                try:
                    mtime = os.path.getmtime(path)
                    if mtime > latest:
                        latest = mtime
                except OSError:
                    continue
    except OSError:
        pass

    # Fallback to directory mtime if no files
    if latest == 0.0:
        try:
            latest = os.path.getmtime(session_dir)
        except OSError:
            latest = 0.0

    return latest


from functools import lru_cache

@lru_cache(maxsize=128)
def get_session_stats_cached(session_dir: str, version: float):
    """
    Internal cached function keyed by (session_dir, version).
    Do not call directly; use get_session_stats().
    """
    pix_dict, beamsize_dict, caldate_dict = get_pixel_table()
    mode = "collection" if is_collection_session(session_dir) else "screening"


    if mode == "screening":
        df_all, atlas_path, instrument_model = process_directory_screening(
            session_dir, pix_dict, beamsize_dict, caldate_dict
        )
    else:
        df_all, atlas_path, instrument_model, cam_name = process_directory_collection(
            session_dir, pix_dict, beamsize_dict, caldate_dict
        )


    return df_all, atlas_path, instrument_model, mode

def get_session_stats(session_dir: str):
    """
    Public helper: compute or retrieve cached stats for a session.
    Automatically invalidates when the session_dir mtime changes.
    """
    version = get_session_version(session_dir)
    return get_session_stats_cached(session_dir, version)

@lru_cache(maxsize=128)
def get_session_nodes_cached(session_dir: str, version: float):
    """
    Internal cached function keyed by (session_dir, version).
    Do not call directly; use get_session_nodes().
    """
    df_all, atlas_path, instrument_model, mode = get_session_stats(session_dir)
    atlas_root, atlas_source = detect_atlas_root(session_dir, atlas_arg=None, summary_text="")
    nodes = build_session_nodes(session_dir, atlas_root)
    return nodes, atlas_root, atlas_source

def get_session_nodes(session_dir: str):
    """
    Public helper: compute or retrieve cached nodes for a session.
    Automatically invalidates when the session_dir mtime changes.
    """
    version = get_session_version(session_dir)
    return get_session_nodes_cached(session_dir, version)

@app.route("/")
def index():
    sessions = find_sessions()
    return render_template("index.html", sessions=sessions, base_root=BASE_ROOT)

@app.route("/session/<session_id>/summary.json")
def session_summary_json(session_id):
    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    df_all, atlas_path, instrument_model, mode = get_session_stats(session_dir)
    summary_rows = build_summary_rows(df_all, instrument_model, mode)

    # Notes (same logic as session_view)
    if mode == "screening":
        notes = [
            "These statistics are for the first image taken in the screening set. "
            "If you took images at different microscope settings, this will not be correct for all images.",
            "The dose is approximated from the first micrograph. The total dose on specimen is slightly higher.",
            "The hole size and spacing is guessed based on the measure hole size function in EPU. "
            "If you are using an uncommon hole size/spacing, it may misidentify it.",
            "Pixel size is listed both as the pixel size automatically coded in EPU as well as the experimentally-calibrated pixel size. "
            "I advise that you use the calibrated pixel size in processing.",
            "Please contact Talya if any of these numbers appear to be incorrect! The script may need updating.",
        ]
    else:
        notes = [
            "If you took images at different microscope settings, this will not be correct for all images.",
            "The dose is approximated from the first movie. The total dose on specimen is slightly higher.",
            "The hole size and spacing is guessed based on the measure hole size function in EPU. "
            "If you are using an uncommon hole size/spacing, it may misidentify it.",
            "Pixel size is listed both as the pixel size automatically coded in EPU as well as the experimentally-calibrated pixel size. "
            "I advise that you use the calibrated pixel size in processing.",
            "Please contact Talya if any of these numbers appear to be incorrect! The script may need updating.",
        ]

    return jsonify({
        "mode": mode,
        "summary_rows": summary_rows,  # list of [key, value]
        "notes": notes,
    })

@app.route("/session/<session_id>/nodes.json")
def session_nodes_json(session_id):
    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    # Use cached nodes (includes versioning via get_session_version)
    nodes, atlas_root, atlas_source = get_session_nodes(session_dir)

    out_nodes = []
    for gs in nodes:
        gs_index = gs.get("index")

        # Plasmon availability + URL
        has_plasmon = False
        plasmon_url = None
        try:
            gs_dir = gs.get("gs_dir")
            if gs_dir and gs_index:
                support_path, nonsupport_path = find_latest_gridsquare_support_and_nonsupport(gs_dir)
                main_base = support_path or nonsupport_path
                if nonsupport_path and os.path.isfile(nonsupport_path):
                    if not (main_base and os.path.isfile(main_base) and os.path.realpath(nonsupport_path) == os.path.realpath(main_base)):
                        has_plasmon = True
                        plasmon_url = url_for(
                            "session_gridsquare_plasmon",
                            session_id=session_id,
                            gs_index=gs_index,
                        )
        except Exception:
            pass

        children = []
        for ch in gs["children"]:
            ch_index = ch.get("index")

            foilhole_url = None
            micro_url = None  # important: define it

            has_fh = bool(ch.get("foilhole_img_path"))

            paths = ch.get("micrograph_img_paths") or []
            has_mg = bool(paths) or bool(ch.get("micrograph_img_path"))

            acq_areas = ch.get("micrograph_acq_areas") or []
            n_acq_areas = ch.get("n_acq_areas")

            if has_fh and gs_index and ch_index:
                foilhole_url = url_for(
                    "session_foilhole_hole",
                    session_id=session_id,
                    gs_index=gs_index,
                    child_index=ch_index,
                )

            micro_urls = []
            if gs_index and ch_index:
                for i in range(len(paths)):
                    micro_urls.append(url_for(
                        "session_foilhole_micro_index",
                        session_id=session_id,
                        gs_index=gs_index,
                        child_index=ch_index,
                        micro_index=i
                    ))

            children.append({
                "index": ch_index,
                "key": ch.get("key"),
                "has_foilhole": has_fh,
                "has_micrograph": has_mg,
                "foilhole_url": foilhole_url,
                "micro_urls": micro_urls,
                "micro_url": micro_url,
                "n_acq_areas": n_acq_areas,
                "micro_acq_areas": acq_areas,
            })

        out_nodes.append({
            "index": gs_index, 
            "name": gs.get("name"),
            "epu": gs.get("epu"), 
            "has_plasmon": has_plasmon,
            "plasmon_url": plasmon_url,
            "children": children,
        }) 

    return jsonify({
        "nodes": out_nodes
    })

@app.route("/session/<session_id>")
def session_view(session_id):
    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    if session_dir.startswith(BASE_ROOT + os.sep):
        display_name = session_dir[len(BASE_ROOT) + 1:]
    else:
        display_name = os.path.basename(session_dir)
    refresh = request.args.get("refresh") == "1"
    if refresh:
        get_session_stats_cached.cache_clear()
        get_session_nodes_cached.cache_clear()
        image_cache.clear()
    return render_template(
        "session.html",
        session_id=session_id,
        display_name=display_name,
    )

def _pil_to_response(img, fmt="PNG"):
    if img is None:
        abort(404)
    buf = io.BytesIO()
    img.save(buf, format=fmt)
    buf.seek(0)
    return send_file(buf, mimetype=f"image/{fmt.lower()}")

def warm_recent_sessions_background():
    """
    Background loop: periodically warm caches for the most recent sessions.
    """
    while True:
        try:
            sessions = find_sessions()[:WARMUP_SESSIONS_LIMIT]
            for s in sessions:
                session_dir = s["path"]
                print(f"[warmup] Warming session: {session_dir}")
                warm_session(session_dir)
        except Exception as e:
            print(f"[warmup] Error in warmup loop: {e}")
        time.sleep(WARMUP_INTERVAL_SECONDS)

def start_warmup_thread():
    t = threading.Thread(target=warm_recent_sessions_background, daemon=True)
    t.start()

@app.route("/session/<session_id>/atlas")
def session_atlas(session_id):
    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    version = get_session_version(session_dir)
    cache_key = ("atlas", session_dir, version)

    def generate():
        _, atlas_root, atlas_source = get_session_nodes(session_dir)

        # 1) Try annotated atlas (but don't hide failures)
        if atlas_root and annotate_atlas_pair is not None and atlas_source in ("dm_atlasid", "dm_hint", "cli"):
            try:
                return annotate_atlas_pair(session_dir, atlas_root)
            except Exception:
                app.logger.exception("annotate_atlas_pair failed for session=%s atlas_root=%s", session_dir, atlas_root)

        # 2) NEW: fall back to latest atlas JPG inside the atlas_root
        if atlas_root:
            jpg = find_latest_atlas_jpg(atlas_root, session_dir=session_dir)
            if jpg:
                try:
                    return Image.open(jpg).convert("RGB")
                except Exception:
                    app.logger.exception("Failed to open atlas jpg: %s", jpg)

        # 3) Existing fallback: look for atlas jpgs in the session folder itself
        fallbacks = find_fallback_atlas_jpgs(session_dir)
        if fallbacks:
            try:
                return Image.open(fallbacks[0]).convert("RGB")
            except Exception:
                return None

        return None

    data = get_cached_image_bytes(cache_key, generate, fmt="JPEG")
    if data is None:
        abort(404)
    return send_file(io.BytesIO(data), mimetype="image/jpeg")


@app.route("/session/<session_id>/template")
def session_template(session_id):
    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    version = get_session_version(session_dir)
    cache_key = ("template", session_dir, version)

    def generate():
        df_all, _, _, mode = get_session_stats(session_dir)
        row_series = df_all.iloc[0].copy()
        beam_diameter_stats_m = None
        if "Beam Size (um)" in row_series.index:
            beam_um_val = row_series["Beam Size (um)"]
            try:
                if beam_um_val is not None and beam_um_val != "Beam size not calibrated":
                    beam_diameter_stats_m = float(beam_um_val) * 1e-6
            except (TypeError, ValueError):
                beam_diameter_stats_m = None
        return annotate_foilhole_template(session_dir, beam_diameter_stats_m=beam_diameter_stats_m)

    data = get_cached_image_bytes(cache_key, generate, fmt="JPEG")
    if data is None:
        abort(404)
    return send_file(io.BytesIO(data), mimetype="image/jpeg")


@app.route("/session/<session_id>/gridsquare/<int:gs_index>")
def session_gridsquare(session_id, gs_index):
    import re
    from datetime import datetime

    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    nodes, atlas_root, _ = get_session_nodes(session_dir)
    node = next((n for n in nodes if n.get("index") == gs_index), None)
    if node is None:
        abort(404)

    gs_dir = node["gs_dir"]

    # --- Determine cutoff_dt based on "first micrograph on Grid Square 1" ---
    cutoff_dt = None

    # Rule 1: if only 1 gridsquare, proceed as usual
    if len(nodes) > 1:
        gs1_node = next((n for n in nodes if n.get("index") == 1), None)

        # Only attempt filtering if we can identify GS1 by index
        if gs1_node is not None:
            gs1_dir = gs1_node["gs_dir"]
            data_dir = os.path.join(gs1_dir, "Data")

            # Match your micrograph JPG naming scheme
            MICROGRAPH_JPG_RE = re.compile(
                r"^FoilHole_([A-Za-z0-9]+)_Data_[^_]+_[^_]+_(\d{8})_(\d{6})\.jpg$",
                re.IGNORECASE,
            )

            earliest = None
            if os.path.isdir(data_dir):
                try:
                    for name in os.listdir(data_dir):
                        m = MICROGRAPH_JPG_RE.match(name)
                        if not m:
                            continue
                        date_str, time_str = m.group(2), m.group(3)
                        try:
                            dt = datetime.strptime(date_str + time_str, "%Y%m%d%H%M%S")
                        except Exception:
                            continue
                        if earliest is None or dt < earliest:
                            earliest = dt
                except Exception:
                    earliest = None

            # Rule 2: if no successful micrographs on GS1, proceed as usual
            cutoff_dt = earliest

    # Only filter OTHER gridsquares (Rule 3)
    min_ts_for_this_gs = None
    if gs_index != 1 and cutoff_dt is not None:
        min_ts_for_this_gs = cutoff_dt

    version = get_session_version(session_dir)
    # include cutoff in cache key so filtered/unfiltered renders don't collide
    cutoff_key = min_ts_for_this_gs.isoformat() if min_ts_for_this_gs else None
    cache_key = ("gs", session_dir, version, gs_index, cutoff_key)

    def generate():
        img = None

        if annotate_gridsquare_image_or_pair is not None:
            try:
                # requires your annotate_gridsquare.py patch: annotate_gridsquare_image_or_pair(gs_dir, min_ts=...)
                img = annotate_gridsquare_image_or_pair(gs_dir, min_ts=min_ts_for_this_gs)
            except TypeError:
                # backwards compatible if not yet patched
                img = annotate_gridsquare_image_or_pair(gs_dir)
            except Exception:
                img = None

        if img is None and annotate_single_gridsquare_image is not None:
            try:
                img = annotate_single_gridsquare_image(gs_dir, min_ts=min_ts_for_this_gs)
            except TypeError:
                img = annotate_single_gridsquare_image(gs_dir)
            except Exception:
                img = None

        return img

    data = get_cached_image_bytes(cache_key, generate, fmt="JPEG")
    if data is None:
        abort(404)
    return send_file(io.BytesIO(data), mimetype="image/jpeg")

@app.route("/session/<session_id>/gridsquare/<int:gs_index>/plasmon")
def session_gridsquare_plasmon(session_id, gs_index):
    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    nodes, _, _ = get_session_nodes(session_dir)
    node = next((n for n in nodes if n.get("index") == gs_index), None)
    if node is None:
        abort(404)

    gs_dir = node["gs_dir"]
    support_path, nonsupport_path = find_latest_gridsquare_support_and_nonsupport(gs_dir)

    # Plasmon = non-support, but only show it if it differs from the main base (support if present)
    main_base = support_path or nonsupport_path
    if not nonsupport_path or not os.path.isfile(nonsupport_path):
        abort(404)
    if main_base and os.path.isfile(main_base):
        try:
            if os.path.realpath(nonsupport_path) == os.path.realpath(main_base):
                abort(404)
        except Exception:
            pass

    version = get_session_version(session_dir)
    cache_key = ("gs_plasmon", session_dir, version, gs_index)

    def generate():
        img = Image.open(nonsupport_path).convert("RGB")
        # Reuse the caption helper you already have in epu.annotate_gridsquare
        try:
            from epu.annotate_gridsquare import add_plasmon_caption
            img = add_plasmon_caption(img, "Energy filter plasmon image: black = empty hole")
        except Exception:
            pass
        return img

    data = get_cached_image_bytes(cache_key, generate, fmt="JPEG")
    if data is None:
        abort(404)
    return send_file(io.BytesIO(data), mimetype="image/jpeg")

@app.route("/session/<session_id>/foilhole/<int:gs_index>/<int:child_index>/hole")
def session_foilhole_hole(session_id, gs_index, child_index):
    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    # Get latest nodes for this session
    nodes, atlas_root, _ = get_session_nodes(session_dir)
    node = next((n for n in nodes if n.get("index") == gs_index), None)
    if node is None:
        abort(404)

    children = [c for c in node["children"] if c.get("index") == child_index]
    if not children:
        abort(404)
    child = children[0]

    path = child.get("foilhole_img_path")
    if not path or not os.path.isfile(path):
        abort(404)

    img = Image.open(path).convert("RGB")
    img = add_scale_bar_by_xml(
        img,
        path,
        bar_um=1.0,
        align="left",
        font_size=FONT_SIZES["defocus"],
    )
    return _pil_to_response(img, fmt="JPEG")


@app.route("/session/<session_id>/foilhole/<int:gs_index>/<int:child_index>/micro/<int:micro_index>")
def session_foilhole_micro_index(session_id, gs_index, child_index, micro_index):
    try:
        session_dir = decode_path(session_id)
    except ValueError:
        abort(404)
    if not os.path.isdir(session_dir):
        abort(404)

    nodes, atlas_root, _ = get_session_nodes(session_dir)
    node = next((n for n in nodes if n.get("index") == gs_index), None)
    if node is None:
        abort(404)

    children = [c for c in node["children"] if c.get("index") == child_index]
    if not children:
        abort(404)
    child = children[0]

    paths = child.get("micrograph_img_paths") or []
    if micro_index < 0 or micro_index >= len(paths):
        abort(404)
    path = paths[micro_index]

    img = Image.open(path).convert("RGB")
    img = add_scale_bar_by_xml(
        img,
        path,
        bar_nm=50.0,
        align="left",
        add_defocus=True,
        font_size=FONT_SIZES["defocus"],
    )
    return _pil_to_response(img, fmt="JPEG")


if __name__ == "__main__":
    app.run(debug=True)