# session_layout.py

import os
import re
from datetime import datetime
from typing import Optional, List, Dict, Tuple

import xml.etree.ElementTree as ET

from epu.report_utils import open_image_or_none
from epu.report_style import FONT_SIZES
from epu.report_scale_bars import add_scale_bar_by_xml

# Imports from annotators
try:
    from epu.annotate_atlas import map_grids_to_atlas, square_type_and_mtime
except Exception:
    map_grids_to_atlas = None
    square_type_and_mtime = None

try:
    import epu.annotate_gridsquare as ag
    annotate_gridsquare_image_or_pair = getattr(ag, "annotate_gridsquare_image_or_pair", None)
    annotate_single_gridsquare_image = getattr(ag, "annotate_single_gridsquare_image", None)
    find_unique_foilhole_xmls_earliest_latest = getattr(ag, "find_unique_foilhole_xmls_earliest_latest", None)
    get_selected_holes_for_gridsquare = getattr(ag, "get_selected_holes_for_gridsquare", None)
    add_plasmon_caption = getattr(ag, "add_plasmon_caption", None)
except Exception:
    annotate_gridsquare_image_or_pair = None
    annotate_single_gridsquare_image = None
    find_unique_foilhole_xmls_earliest_latest = None
    get_selected_holes_for_gridsquare = None
    add_plasmon_caption = None

# ---------- Patterns and helpers (from generate_report.py, without reportlab) ----------

GRID_IMG_RE = re.compile(r"^GridSquare_(\d{8})_(\d{6})\.jpg$", re.IGNORECASE)
GRID_SUPPORT_IMG_RE = re.compile(r"^GridSquare_Support_(\d{8})_(\d{6})\.jpg$", re.IGNORECASE)
FOILHOLE_RE = re.compile(r"^FoilHole_([A-Za-z0-9]+)_(\d{8})_(\d{6})\.jpg$", re.IGNORECASE)
MICROGRAPH_RE = re.compile(
    r"^FoilHole_([A-Za-z0-9]+)_Data_[^_]+_[^_]+_(\d{8})_(\d{6})\.jpg$", re.IGNORECASE
)
GS_ID_RE = re.compile(r"grid\s*square[_\s-]*([0-9]+)", re.IGNORECASE)

def _dt_from_foilhole_filename(name: str) -> Optional[datetime]:
    m = FOILHOLE_RE.match(name)
    if not m:
        return None
    date_str, time_str = m.group(2), m.group(3)
    dt = parse_datetime_tokens(date_str, time_str)
    return dt if isinstance(dt, datetime) else None

def _dt_from_micrograph_filename(name: str) -> Optional[datetime]:
    m = MICROGRAPH_RE.match(name)
    if not m:
        return None
    date_str, time_str = m.group(2), m.group(3)
    dt = parse_datetime_tokens(date_str, time_str)
    return dt if isinstance(dt, datetime) else None

def _first_micrograph_dt_in_gridsquare(gs_dir: str) -> Optional[datetime]:
    data_dir = os.path.join(gs_dir, "Data")
    if not os.path.isdir(data_dir):
        return None
    earliest = None
    for name in os.listdir(data_dir):
        dt = _dt_from_micrograph_filename(name)
        if dt is None:
            continue
        if earliest is None or dt < earliest:
            earliest = dt
    return earliest

def parse_datetime_tokens(date_str, time_str):
    try:
        return datetime.strptime(date_str + time_str, "%Y%m%d%H%M%S")
    except Exception:
        return (date_str, time_str)

def gridsquare_images(gs_dir: str):
    """
    Return (latest_support_path, latest_non_support_path) for this GridSquare directory.
    Each may be None if not present.
    """
    support = []
    nonsupport = []
    try:
        for name in os.listdir(gs_dir):
            m = GRID_SUPPORT_IMG_RE.match(name)
            if m:
                date_str, time_str = m.group(1), m.group(2)
                support.append((os.path.join(gs_dir, name), date_str, time_str))
                continue
            m = GRID_IMG_RE.match(name)
            if m:
                date_str, time_str = m.group(1), m.group(2)
                nonsupport.append((os.path.join(gs_dir, name), date_str, time_str))
    except Exception:
        pass

    def pick_latest(lst):
        if not lst:
            return None
        lst.sort(key=lambda tup: parse_datetime_tokens(tup[1], tup[2]), reverse=True)
        return lst[0][0]

    return pick_latest(support), pick_latest(nonsupport)

def latest_foilholes_per_key(gs_dir: str):
    holes_dir = os.path.join(gs_dir, "FoilHoles")
    if not os.path.isdir(holes_dir):
        return []
    groups = {}
    for name in os.listdir(holes_dir):
        m = FOILHOLE_RE.match(name)
        if not m:
            continue
        key, date_str, time_str = m.group(1), m.group(2), m.group(3)
        path = os.path.join(holes_dir, name)
        dt = parse_datetime_tokens(date_str, time_str)
        prev = groups.get(key)
        if prev is None or dt > prev[0]:
            groups[key] = (dt, path, date_str, time_str)
    out = []
    for key, (_, path, date_str, time_str) in groups.items():
        out.append((key, path, date_str, time_str))
    out.sort(key=lambda x: x[0])
    return out

def find_matching_micrograph(gs_dir: str, foilhole_key: str) -> Optional[str]:
    data_dir = os.path.join(gs_dir, "Data")
    if not os.path.isdir(data_dir):
        return None
    candidates = []
    for name in os.listdir(data_dir):
        m = MICROGRAPH_RE.match(name)
        if not m:
            continue
        key, date_str, time_str = m.group(1), m.group(2), m.group(3)
        if key != foilhole_key:
            continue
        candidates.append((os.path.join(data_dir, name), date_str, time_str))
    if not candidates:
        return None
    candidates.sort(key=lambda tup: parse_datetime_tokens(tup[1], tup[2]), reverse=True)
    return candidates[0][0]

def find_gridsquares(base_folder: str) -> List[str]:
    gs_root = os.path.join(base_folder, "Images-Disc1")
    if not os.path.isdir(gs_root):
        return []
    gs_dirs = [
        os.path.join(gs_root, d)
        for d in os.listdir(gs_root)
        if d.startswith("GridSquare") and os.path.isdir(os.path.join(gs_root, d))
    ]
    gs_dirs.sort()
    return gs_dirs

def extract_epu_from_gridsquare_name(gs_name: str) -> Optional[str]:
    m = GS_ID_RE.search(gs_name or "")
    return m.group(1) if m else None

def _ln(tag: str) -> str:
    return tag.split("}")[-1] if isinstance(tag, str) else tag

def extract_sample_and_root_from_atlas_path(p: str) -> Optional[Tuple[str, str]]:
    """
    Given a path like ...\\Sample4\\Atlas\\Atlas.dm (or ...\\Sample4\\Atlas),
    return (sample_dir, atlas_root_name).
    """
    if not p:
        return None
    p = p.strip().strip('"').strip("'")
    if re.search(r"(?i)\batlas\.dm$", p):
        atlas_dir = os.path.dirname(p)
    else:
        atlas_dir = p
    last = os.path.basename(atlas_dir)
    if last.lower() != "atlas":
        return None
    sample_dir = os.path.basename(os.path.dirname(atlas_dir))
    if not re.match(r"(?i)^sample\d+$", sample_dir):
        return None
    atlas_root_dir = os.path.dirname(os.path.dirname(atlas_dir))
    atlas_root_name = os.path.basename(atlas_root_dir)
    return sample_dir, atlas_root_name

def atlas_root_is_valid(root: str) -> bool:
    if not os.path.isdir(root):
        return False
    try:
        for name in os.listdir(root):
            if re.match(r"(?i)^sample\d+$", name):
                adir = os.path.join(root, name, "Atlas")
                if os.path.isfile(os.path.join(adir, "Atlas.dm")):
                    return True
    except Exception:
        pass
    return False

def atlas_id_from_epu_dm(session_dir: str) -> Optional[str]:
    dm_path = os.path.join(session_dir, "EpuSession.dm")
    if not os.path.isfile(dm_path):
        return None
    try:
        root = ET.parse(dm_path).getroot()
    except Exception:
        return None
    for elem in root.iter():
        if _ln(elem.tag).lower() == "atlasid":
            txt = (elem.text or "").strip()
            return txt if txt else None
    return None

def atlas_name_from_epu_dm_path(session_dir: str) -> Optional[str]:
    """
    Robustly extract the atlas root folder name from EpuSession.dm by locating the
    <AtlasId> element text (a path like ...\\<atlas_root>\\Sample0\\Atlas\\Atlas.dm).

    Returns the atlas root folder name (the folder right before SampleN), or None.
    """
    dm_path = os.path.join(session_dir, "EpuSession.dm")
    if not os.path.isfile(dm_path):
        return None

    try:
        root = ET.parse(dm_path).getroot()
    except Exception:
        return None

    atlas_path = None
    for elem in root.iter():
        if _ln(elem.tag).lower() == "atlasid":
            txt = (elem.text or "").strip()
            if txt:
                atlas_path = txt
                break

    if not atlas_path:
        return None

    # Split Windows or POSIX paths safely
    parts = [p for p in re.split(r"[\\/]+", atlas_path.strip().strip('"').strip("'")) if p]

    # Find "SampleN" and return the part immediately before it (atlas root folder name)
    for i, part in enumerate(parts):
        if re.match(r"(?i)^sample\d+$", part):
            if i - 1 >= 0:
                name = parts[i - 1]
                return name if name else None
            return None

    return None

def normalize_atlas_arg(a: str) -> str:
    a_abs = os.path.abspath(a)
    if os.path.basename(a_abs).lower() == "atlas":
        if os.path.isfile(os.path.join(a_abs, "Atlas.dm")):
            return os.path.dirname(os.path.dirname(a_abs))
    elif os.path.basename(a_abs).lower() == "atlas.dm" and os.path.isfile(a_abs):
        adir = os.path.dirname(a_abs)
        if os.path.basename(adir).lower() == "atlas":
            return os.path.dirname(os.path.dirname(adir))
    return a_abs

def detect_atlas_root(
    session_dir: str,
    atlas_arg: Optional[str],
    summary_text: str = "",
) -> Tuple[Optional[str], Optional[str]]:
    """
    Detect atlas root. Returns (atlas_root_path, atlas_source),
    where atlas_source is one of: 'cli', 'dm_atlasid', 'dm_hint', or None.
    """
    session_dir = os.path.abspath(session_dir)
    parent_dir = os.path.dirname(session_dir)

    if atlas_arg:
        chosen = normalize_atlas_arg(atlas_arg)
        return chosen, "cli"


    atlas_id_text = atlas_id_from_epu_dm(session_dir)
    if atlas_id_text:
        parsed = extract_sample_and_root_from_atlas_path(atlas_id_text)
        if parsed:
            sample_dir, atlas_root_name = parsed
            for base in (session_dir, parent_dir):
                candidate = os.path.join(base, atlas_root_name)
                if atlas_root_is_valid(candidate):
                    return candidate, "dm_atlasid"
        else:
            for base in (session_dir, parent_dir):
                candidate = os.path.join(base, atlas_id_text)
                if atlas_root_is_valid(candidate):
                    return candidate, "dm_atlasid"


    dm_name = atlas_name_from_epu_dm_path(session_dir)
    if dm_name:
        for base in (session_dir, parent_dir):
            candidate = os.path.join(base, dm_name)
            if atlas_root_is_valid(candidate):
                return candidate, "dm_hint"

    return None, None

def find_latest_atlas_jpg(atlas_root: str, session_dir: Optional[str] = None) -> Optional[str]:
    def collect_from_sample(sample: str) -> List[str]:
        adir = os.path.join(atlas_root, sample, "Atlas")
        if os.path.isdir(adir):
            try:
                return [
                    os.path.join(adir, n)
                    for n in os.listdir(adir)
                    if n.lower().startswith("atlas") and n.lower().endswith(".jpg")
                ]
            except Exception:
                return []
        return []

    preferred_sample = None
    if session_dir:
        txt = atlas_id_from_epu_dm(session_dir)
        if txt:
            parsed = extract_sample_and_root_from_atlas_path(txt)
            if parsed:
                preferred_sample, _ = parsed

    candidates: List[str] = []
    tried_samples: List[str] = []

    if preferred_sample:
        candidates = collect_from_sample(preferred_sample)
        tried_samples.append(preferred_sample)

    if not candidates:
        if "Sample0" not in tried_samples:
            candidates = collect_from_sample("Sample0")
            tried_samples.append("Sample0")

    if not candidates:
        try:
            sample_dirs = [n for n in os.listdir(atlas_root) if re.match(r"(?i)^sample\d+$", n)]
            for s in sample_dirs:
                if s in tried_samples:
                    continue
                files = collect_from_sample(s)
                if files:
                    candidates = files
                    break
        except Exception:
            pass

    if not candidates:
        return None

    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]

def find_fallback_atlas_jpgs(session_dir: str) -> List[str]:
    results: List[str] = []
    if not os.path.isdir(session_dir):
        return results

    try:
        for n in os.listdir(session_dir):
            if n.lower().endswith(".jpg") and "atlas" in n.lower():
                results.append(os.path.join(session_dir, n))
    except Exception:
        pass

    if os.path.basename(session_dir).lower() == "epu_out":
        parent = os.path.dirname(session_dir)
        if os.path.isdir(parent):
            try:
                for n in os.listdir(parent):
                    if n.lower().endswith(".jpg") and "atlas" in n.lower():
                        results.append(os.path.join(parent, n))
            except Exception:
                pass

    seen = set()
    deduped = []
    for p in results:
        if p not in seen:
            deduped.append(p)
            seen.add(p)
    return deduped

def compute_gridsquare_index_map(session_dir: str, atlas_root: Optional[str]) -> Dict[str, int]:
    if not atlas_root or map_grids_to_atlas is None or square_type_and_mtime is None:
        return {}
    try:
        df, _, _, _ = map_grids_to_atlas(atlas_root, session_dir, check_node_center=True, fill_rotation='median')
        if df is None or df.empty:
            return {}
        types, colors, mtimes = [], [], []
        for _, row in df.iterrows():
            color, typ, mt = square_type_and_mtime(row['folder'])
            types.append(typ)
            colors.append(color)
            mtimes.append(mt)
        df['square_type'] = types
        df['color'] = colors
        df['square_first_mtime'] = mtimes
        df = df.sort_values(by='square_first_mtime', ascending=True, na_position='last')
        df['grid_square_index'] = range(1, len(df) + 1)
        mapping = {}
        for _, row in df.iterrows():
            mapping[os.path.realpath(row['folder'])] = int(row['grid_square_index'])
        return mapping
    except Exception:
        return {}

def build_session_nodes(session_dir: str, atlas_root: Optional[str]):
    """
    Build a list of GridSquare nodes with children (FoilHoles).

    Filtering:
      1) If there is only 1 grid square: do nothing special.
      2) If multiple grid squares: find the first micrograph timestamp on "Grid Square 1".
         - If GS1 has no micrographs: do nothing special.
      3) For all OTHER grid squares: drop any FoilHole JPG whose timestamp is earlier
         than the first GS1 micrograph timestamp. These dropped FoilHoles:
            - are not shown as thumbnails
            - are not used for "selected" indexing
    """
    def _first_micrograph_dt_in_gridsquare(gs_dir: str) -> Optional[datetime]:
        data_dir = os.path.join(gs_dir, "Data")
        if not os.path.isdir(data_dir):
            return None
        earliest = None
        try:
            for name in os.listdir(data_dir):
                m = MICROGRAPH_RE.match(name)
                if not m:
                    continue
                dt = parse_datetime_tokens(m.group(2), m.group(3))
                if not isinstance(dt, datetime):
                    continue
                if earliest is None or dt < earliest:
                    earliest = dt
        except Exception:
            return None
        return earliest

    gs_index_map = compute_gridsquare_index_map(session_dir, atlas_root)
    gs_dirs = find_gridsquares(session_dir)
    nodes = []

    # --- Determine GS1 and cutoff timestamp to avoid showing template definition FoilHole images ---
    cutoff_dt: Optional[datetime] = None
    gs1_dir: Optional[str] = None

    if len(gs_dirs) > 1:
        # Prefer the grid square that atlas mapping assigns index 1
        if gs_index_map:
            for d in gs_dirs:
                if gs_index_map.get(os.path.realpath(d)) == 1:
                    gs1_dir = d
                    break

        # Fallback: first directory in sorted order
        if gs1_dir is None and gs_dirs:
            gs1_dir = gs_dirs[0]

        if gs1_dir is not None:
            cutoff_dt = _first_micrograph_dt_in_gridsquare(gs1_dir)
            # If cutoff_dt is None (no micrographs in GS1), proceed as usual.

    for gs_dir in gs_dirs:
        gs_name = os.path.basename(gs_dir)
        support_img_path, nonsupport_img_path = gridsquare_images(gs_dir)
        gs_img_path = support_img_path or nonsupport_img_path
        gs_epu = extract_epu_from_gridsquare_name(gs_name)
        gs_index = gs_index_map.get(os.path.realpath(gs_dir))

        foilholes_latest = latest_foilholes_per_key(gs_dir)

        # Build key -> (path, dt)
        fh_latest_map = {}
        for (key, path, date_str, time_str) in foilholes_latest:
            dt = parse_datetime_tokens(date_str, time_str)
            dt = dt if isinstance(dt, datetime) else None
            fh_latest_map[key] = (path, dt)

        # --- filtering: only for non-GS1 grid squares, only if cutoff exists (to avoid showing template definition foilhole images) ---
        if cutoff_dt is not None and gs1_dir is not None and gs_dir != gs1_dir:
            fh_latest_map = {
                k: (p, dt)
                for k, (p, dt) in fh_latest_map.items()
                if not (dt is not None and dt < cutoff_dt)
            }

        # Convenience: key -> path only (after filtering)
        fh_path_map = {k: v[0] for k, v in fh_latest_map.items()}

        # Selection + indexing (filtered FoilHoles must not be indexed)
        if get_selected_holes_for_gridsquare is not None:
            try:
                sel_keys_order, _sel_idx_map = get_selected_holes_for_gridsquare(gs_dir, max_show=12)
                keys_selected = [k for k in sel_keys_order if k in fh_path_map]
            except Exception:
                keys_selected = list(fh_path_map.keys())
        else:
            keys_selected = list(fh_path_map.keys())

        # Re-number sequentially after filtering (prevents gaps)
        idx_map = {k: i + 1 for i, k in enumerate(keys_selected)}

        children = []
        for key in keys_selected:
            fh_path = fh_path_map.get(key)
            micro = find_matching_micrograph(gs_dir, key)

            child = {
                "key": key,
                "index": idx_map.get(key),
                "foilhole_img_path": fh_path if fh_path and os.path.isfile(fh_path) else None,
                "micrograph_img_path": micro if micro and os.path.isfile(micro) else None,
            }
            children.append(child)

        nodes.append(
            {
                "gs_dir": gs_dir,
                "name": gs_name,
                "epu": gs_epu,
                "index": gs_index,
                "latest_img_path": gs_img_path if gs_img_path and os.path.isfile(gs_img_path) else None,
                "support_img_path": support_img_path if support_img_path and os.path.isfile(support_img_path) else None,
                "nonsupport_img_path": nonsupport_img_path if nonsupport_img_path and os.path.isfile(nonsupport_img_path) else None,
                "children": children,
            }
        )

    try:
        nodes.sort(key=lambda n: (n.get("index") is None, n.get("index", 10**9), n.get("name", "")))
    except Exception:
        pass

    return nodes


# session_layout.py

import os
import re
from datetime import datetime
from typing import Optional, List, Dict, Tuple

import xml.etree.ElementTree as ET

from epu.report_utils import open_image_or_none
from epu.report_style import FONT_SIZES
from epu.report_scale_bars import add_scale_bar_by_xml

# Imports from annotators
try:
    from epu.annotate_atlas import map_grids_to_atlas, square_type_and_mtime
except Exception:
    map_grids_to_atlas = None
    square_type_and_mtime = None

try:
    import epu.annotate_gridsquare as ag
    annotate_gridsquare_image_or_pair = getattr(ag, "annotate_gridsquare_image_or_pair", None)
    annotate_single_gridsquare_image = getattr(ag, "annotate_single_gridsquare_image", None)
    find_unique_foilhole_xmls_earliest_latest = getattr(ag, "find_unique_foilhole_xmls_earliest_latest", None)
    get_selected_holes_for_gridsquare = getattr(ag, "get_selected_holes_for_gridsquare", None)
    add_plasmon_caption = getattr(ag, "add_plasmon_caption", None)
except Exception:
    annotate_gridsquare_image_or_pair = None
    annotate_single_gridsquare_image = None
    find_unique_foilhole_xmls_earliest_latest = None
    get_selected_holes_for_gridsquare = None
    add_plasmon_caption = None

# ---------- Patterns and helpers (from generate_report.py, without reportlab) ----------

GRID_IMG_RE = re.compile(r"^GridSquare_(\d{8})_(\d{6})\.jpg$", re.IGNORECASE)
GRID_SUPPORT_IMG_RE = re.compile(r"^GridSquare_Support_(\d{8})_(\d{6})\.jpg$", re.IGNORECASE)
FOILHOLE_RE = re.compile(r"^FoilHole_([A-Za-z0-9]+)_(\d{8})_(\d{6})\.jpg$", re.IGNORECASE)
MICROGRAPH_RE = re.compile(
    r"^FoilHole_([A-Za-z0-9]+)_Data_[^_]+_[^_]+_(\d{8})_(\d{6})\.jpg$", re.IGNORECASE
)
GS_ID_RE = re.compile(r"grid\s*square[_\s-]*([0-9]+)", re.IGNORECASE)

def _dt_from_foilhole_filename(name: str) -> Optional[datetime]:
    m = FOILHOLE_RE.match(name)
    if not m:
        return None
    date_str, time_str = m.group(2), m.group(3)
    dt = parse_datetime_tokens(date_str, time_str)
    return dt if isinstance(dt, datetime) else None

def _dt_from_micrograph_filename(name: str) -> Optional[datetime]:
    m = MICROGRAPH_RE.match(name)
    if not m:
        return None
    date_str, time_str = m.group(2), m.group(3)
    dt = parse_datetime_tokens(date_str, time_str)
    return dt if isinstance(dt, datetime) else None

def _first_micrograph_dt_in_gridsquare(gs_dir: str) -> Optional[datetime]:
    data_dir = os.path.join(gs_dir, "Data")
    if not os.path.isdir(data_dir):
        return None
    earliest = None
    for name in os.listdir(data_dir):
        dt = _dt_from_micrograph_filename(name)
        if dt is None:
            continue
        if earliest is None or dt < earliest:
            earliest = dt
    return earliest

def parse_datetime_tokens(date_str, time_str):
    try:
        return datetime.strptime(date_str + time_str, "%Y%m%d%H%M%S")
    except Exception:
        return (date_str, time_str)

def gridsquare_images(gs_dir: str):
    """
    Return (latest_support_path, latest_non_support_path) for this GridSquare directory.
    Each may be None if not present.
    """
    support = []
    nonsupport = []
    try:
        for name in os.listdir(gs_dir):
            m = GRID_SUPPORT_IMG_RE.match(name)
            if m:
                date_str, time_str = m.group(1), m.group(2)
                support.append((os.path.join(gs_dir, name), date_str, time_str))
                continue
            m = GRID_IMG_RE.match(name)
            if m:
                date_str, time_str = m.group(1), m.group(2)
                nonsupport.append((os.path.join(gs_dir, name), date_str, time_str))
    except Exception:
        pass

    def pick_latest(lst):
        if not lst:
            return None
        lst.sort(key=lambda tup: parse_datetime_tokens(tup[1], tup[2]), reverse=True)
        return lst[0][0]

    return pick_latest(support), pick_latest(nonsupport)

def latest_foilholes_per_key(gs_dir: str):
    holes_dir = os.path.join(gs_dir, "FoilHoles")
    if not os.path.isdir(holes_dir):
        return []
    groups = {}
    for name in os.listdir(holes_dir):
        m = FOILHOLE_RE.match(name)
        if not m:
            continue
        key, date_str, time_str = m.group(1), m.group(2), m.group(3)
        path = os.path.join(holes_dir, name)
        dt = parse_datetime_tokens(date_str, time_str)
        prev = groups.get(key)
        if prev is None or dt > prev[0]:
            groups[key] = (dt, path, date_str, time_str)
    out = []
    for key, (_, path, date_str, time_str) in groups.items():
        out.append((key, path, date_str, time_str))
    out.sort(key=lambda x: x[0])
    return out

def find_matching_micrograph(gs_dir: str, foilhole_key: str) -> Optional[str]:
    data_dir = os.path.join(gs_dir, "Data")
    if not os.path.isdir(data_dir):
        return None
    candidates = []
    for name in os.listdir(data_dir):
        m = MICROGRAPH_RE.match(name)
        if not m:
            continue
        key, date_str, time_str = m.group(1), m.group(2), m.group(3)
        if key != foilhole_key:
            continue
        candidates.append((os.path.join(data_dir, name), date_str, time_str))
    if not candidates:
        return None
    candidates.sort(key=lambda tup: parse_datetime_tokens(tup[1], tup[2]), reverse=True)
    return candidates[0][0]

def find_gridsquares(base_folder: str) -> List[str]:
    gs_root = os.path.join(base_folder, "Images-Disc1")
    if not os.path.isdir(gs_root):
        return []
    gs_dirs = [
        os.path.join(gs_root, d)
        for d in os.listdir(gs_root)
        if d.startswith("GridSquare") and os.path.isdir(os.path.join(gs_root, d))
    ]
    gs_dirs.sort()
    return gs_dirs

def extract_epu_from_gridsquare_name(gs_name: str) -> Optional[str]:
    m = GS_ID_RE.search(gs_name or "")
    return m.group(1) if m else None

def _ln(tag: str) -> str:
    return tag.split("}")[-1] if isinstance(tag, str) else tag

def extract_sample_and_root_from_atlas_path(p: str) -> Optional[Tuple[str, str]]:
    """
    Given a path like ...\\Sample4\\Atlas\\Atlas.dm (or ...\\Sample4\\Atlas),
    return (sample_dir, atlas_root_name).
    """
    if not p:
        return None
    p = p.strip().strip('"').strip("'")
    if re.search(r"(?i)\batlas\.dm$", p):
        atlas_dir = os.path.dirname(p)
    else:
        atlas_dir = p
    last = os.path.basename(atlas_dir)
    if last.lower() != "atlas":
        return None
    sample_dir = os.path.basename(os.path.dirname(atlas_dir))
    if not re.match(r"(?i)^sample\d+$", sample_dir):
        return None
    atlas_root_dir = os.path.dirname(os.path.dirname(atlas_dir))
    atlas_root_name = os.path.basename(atlas_root_dir)
    return sample_dir, atlas_root_name

def atlas_root_is_valid(root: str) -> bool:
    if not os.path.isdir(root):
        return False
    try:
        for name in os.listdir(root):
            if re.match(r"(?i)^sample\d+$", name):
                adir = os.path.join(root, name, "Atlas")
                if os.path.isfile(os.path.join(adir, "Atlas.dm")):
                    return True
    except Exception:
        pass
    return False

def atlas_id_from_epu_dm(session_dir: str) -> Optional[str]:
    dm_path = os.path.join(session_dir, "EpuSession.dm")
    if not os.path.isfile(dm_path):
        return None
    try:
        root = ET.parse(dm_path).getroot()
    except Exception:
        return None
    for elem in root.iter():
        if _ln(elem.tag).lower() == "atlasid":
            txt = (elem.text or "").strip()
            return txt if txt else None
    return None

def atlas_name_from_epu_dm_path(session_dir: str) -> Optional[str]:
    """
    Robustly extract the atlas root folder name from EpuSession.dm by locating the
    <AtlasId> element text (a path like ...\\<atlas_root>\\Sample0\\Atlas\\Atlas.dm).

    Returns the atlas root folder name (the folder right before SampleN), or None.
    """
    dm_path = os.path.join(session_dir, "EpuSession.dm")
    if not os.path.isfile(dm_path):
        return None

    try:
        root = ET.parse(dm_path).getroot()
    except Exception:
        return None

    atlas_path = None
    for elem in root.iter():
        if _ln(elem.tag).lower() == "atlasid":
            txt = (elem.text or "").strip()
            if txt:
                atlas_path = txt
                break

    if not atlas_path:
        return None

    # Split Windows or POSIX paths safely
    parts = [p for p in re.split(r"[\\/]+", atlas_path.strip().strip('"').strip("'")) if p]

    # Find "SampleN" and return the part immediately before it (atlas root folder name)
    for i, part in enumerate(parts):
        if re.match(r"(?i)^sample\d+$", part):
            if i - 1 >= 0:
                name = parts[i - 1]
                return name if name else None
            return None

    return None

def normalize_atlas_arg(a: str) -> str:
    a_abs = os.path.abspath(a)
    if os.path.basename(a_abs).lower() == "atlas":
        if os.path.isfile(os.path.join(a_abs, "Atlas.dm")):
            return os.path.dirname(os.path.dirname(a_abs))
    elif os.path.basename(a_abs).lower() == "atlas.dm" and os.path.isfile(a_abs):
        adir = os.path.dirname(a_abs)
        if os.path.basename(adir).lower() == "atlas":
            return os.path.dirname(os.path.dirname(adir))
    return a_abs

def detect_atlas_root(
    session_dir: str,
    atlas_arg: Optional[str],
    summary_text: str = "",
) -> Tuple[Optional[str], Optional[str]]:
    """
    Detect atlas root. Returns (atlas_root_path, atlas_source),
    where atlas_source is one of: 'cli', 'dm_atlasid', 'dm_hint', or None.
    """
    session_dir = os.path.abspath(session_dir)
    parent_dir = os.path.dirname(session_dir)

    if atlas_arg:
        chosen = normalize_atlas_arg(atlas_arg)
        return chosen, "cli"


    atlas_id_text = atlas_id_from_epu_dm(session_dir)
    if atlas_id_text:
        parsed = extract_sample_and_root_from_atlas_path(atlas_id_text)
        if parsed:
            sample_dir, atlas_root_name = parsed
            for base in (session_dir, parent_dir):
                candidate = os.path.join(base, atlas_root_name)
                if atlas_root_is_valid(candidate):
                    return candidate, "dm_atlasid"
        else:
            for base in (session_dir, parent_dir):
                candidate = os.path.join(base, atlas_id_text)
                if atlas_root_is_valid(candidate):
                    return candidate, "dm_atlasid"


    dm_name = atlas_name_from_epu_dm_path(session_dir)
    if dm_name:
        for base in (session_dir, parent_dir):
            candidate = os.path.join(base, dm_name)
            if atlas_root_is_valid(candidate):
                return candidate, "dm_hint"

    return None, None

def find_latest_atlas_jpg(atlas_root: str, session_dir: Optional[str] = None) -> Optional[str]:
    def collect_from_sample(sample: str) -> List[str]:
        adir = os.path.join(atlas_root, sample, "Atlas")
        if os.path.isdir(adir):
            try:
                return [
                    os.path.join(adir, n)
                    for n in os.listdir(adir)
                    if n.lower().startswith("atlas") and n.lower().endswith(".jpg")
                ]
            except Exception:
                return []
        return []

    preferred_sample = None
    if session_dir:
        txt = atlas_id_from_epu_dm(session_dir)
        if txt:
            parsed = extract_sample_and_root_from_atlas_path(txt)
            if parsed:
                preferred_sample, _ = parsed

    candidates: List[str] = []
    tried_samples: List[str] = []

    if preferred_sample:
        candidates = collect_from_sample(preferred_sample)
        tried_samples.append(preferred_sample)

    if not candidates:
        if "Sample0" not in tried_samples:
            candidates = collect_from_sample("Sample0")
            tried_samples.append("Sample0")

    if not candidates:
        try:
            sample_dirs = [n for n in os.listdir(atlas_root) if re.match(r"(?i)^sample\d+$", n)]
            for s in sample_dirs:
                if s in tried_samples:
                    continue
                files = collect_from_sample(s)
                if files:
                    candidates = files
                    break
        except Exception:
            pass

    if not candidates:
        return None

    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]

def find_fallback_atlas_jpgs(session_dir: str) -> List[str]:
    results: List[str] = []
    if not os.path.isdir(session_dir):
        return results

    try:
        for n in os.listdir(session_dir):
            if n.lower().endswith(".jpg") and "atlas" in n.lower():
                results.append(os.path.join(session_dir, n))
    except Exception:
        pass

    if os.path.basename(session_dir).lower() == "epu_out":
        parent = os.path.dirname(session_dir)
        if os.path.isdir(parent):
            try:
                for n in os.listdir(parent):
                    if n.lower().endswith(".jpg") and "atlas" in n.lower():
                        results.append(os.path.join(parent, n))
            except Exception:
                pass

    seen = set()
    deduped = []
    for p in results:
        if p not in seen:
            deduped.append(p)
            seen.add(p)
    return deduped

def compute_gridsquare_index_map(session_dir: str, atlas_root: Optional[str]) -> Dict[str, int]:
    if not atlas_root or map_grids_to_atlas is None or square_type_and_mtime is None:
        return {}
    try:
        df, _ = map_grids_to_atlas(atlas_root, session_dir, check_node_center=True, fill_rotation='median')
        if df is None or df.empty:
            return {}
        types, colors, mtimes = [], [], []
        for _, row in df.iterrows():
            color, typ, mt = square_type_and_mtime(row['folder'])
            types.append(typ)
            colors.append(color)
            mtimes.append(mt)
        df['square_type'] = types
        df['color'] = colors
        df['square_first_mtime'] = mtimes
        df = df.sort_values(by='square_first_mtime', ascending=True, na_position='last')
        df['grid_square_index'] = range(1, len(df) + 1)
        mapping = {}
        for _, row in df.iterrows():
            mapping[os.path.realpath(row['folder'])] = int(row['grid_square_index'])
        return mapping
    except Exception:
        return {}

def build_session_nodes(session_dir: str, atlas_root: Optional[str]):
    """
    Build a list of GridSquare nodes with children (FoilHoles).

    Filtering:
      1) If there is only 1 grid square: do nothing special.
      2) If multiple grid squares: find the first micrograph timestamp on "Grid Square 1".
         - If GS1 has no micrographs: do nothing special.
      3) For all OTHER grid squares: drop any FoilHole JPG whose timestamp is earlier
         than the first GS1 micrograph timestamp. These dropped FoilHoles:
            - are not shown as thumbnails
            - are not used for "selected" indexing
    """
    def _first_micrograph_dt_in_gridsquare(gs_dir: str) -> Optional[datetime]:
        data_dir = os.path.join(gs_dir, "Data")
        if not os.path.isdir(data_dir):
            return None
        earliest = None
        try:
            for name in os.listdir(data_dir):
                m = MICROGRAPH_RE.match(name)
                if not m:
                    continue
                dt = parse_datetime_tokens(m.group(2), m.group(3))
                if not isinstance(dt, datetime):
                    continue
                if earliest is None or dt < earliest:
                    earliest = dt
        except Exception:
            return None
        return earliest

    gs_index_map = compute_gridsquare_index_map(session_dir, atlas_root)
    gs_dirs = find_gridsquares(session_dir)
    nodes = []

    # --- Determine GS1 and cutoff timestamp to avoid showing template definition FoilHole images ---
    cutoff_dt: Optional[datetime] = None
    gs1_dir: Optional[str] = None

    if len(gs_dirs) > 1:
        # Prefer the grid square that atlas mapping assigns index 1
        if gs_index_map:
            for d in gs_dirs:
                if gs_index_map.get(os.path.realpath(d)) == 1:
                    gs1_dir = d
                    break

        # Fallback: first directory in sorted order
        if gs1_dir is None and gs_dirs:
            gs1_dir = gs_dirs[0]

        if gs1_dir is not None:
            cutoff_dt = _first_micrograph_dt_in_gridsquare(gs1_dir)
            # If cutoff_dt is None (no micrographs in GS1), proceed as usual.

    for gs_dir in gs_dirs:
        gs_name = os.path.basename(gs_dir)
        support_img_path, nonsupport_img_path = gridsquare_images(gs_dir)
        gs_img_path = support_img_path or nonsupport_img_path
        gs_epu = extract_epu_from_gridsquare_name(gs_name)
        gs_index = gs_index_map.get(os.path.realpath(gs_dir))

        foilholes_latest = latest_foilholes_per_key(gs_dir)

        # Build key -> (path, dt)
        fh_latest_map = {}
        for (key, path, date_str, time_str) in foilholes_latest:
            dt = parse_datetime_tokens(date_str, time_str)
            dt = dt if isinstance(dt, datetime) else None
            fh_latest_map[key] = (path, dt)

        # --- filtering: only for non-GS1 grid squares, only if cutoff exists (to avoid showing template definition foilhole images) ---
        if cutoff_dt is not None and gs1_dir is not None and gs_dir != gs1_dir:
            fh_latest_map = {
                k: (p, dt)
                for k, (p, dt) in fh_latest_map.items()
                if not (dt is not None and dt < cutoff_dt)
            }

        # Convenience: key -> path only (after filtering)
        fh_path_map = {k: v[0] for k, v in fh_latest_map.items()}

        # Selection + indexing (filtered FoilHoles must not be indexed)
        if get_selected_holes_for_gridsquare is not None:
            try:
                sel_keys_order, _sel_idx_map = get_selected_holes_for_gridsquare(gs_dir, max_show=12)
                keys_selected = [k for k in sel_keys_order if k in fh_path_map]
            except Exception:
                keys_selected = list(fh_path_map.keys())
        else:
            keys_selected = list(fh_path_map.keys())

        # Re-number sequentially after filtering (prevents gaps)
        idx_map = {k: i + 1 for i, k in enumerate(keys_selected)}

        children = []
        for key in keys_selected:
            fh_path = fh_path_map.get(key)
            micro = find_matching_micrograph(gs_dir, key)

            child = {
                "key": key,
                "index": idx_map.get(key),
                "foilhole_img_path": fh_path if fh_path and os.path.isfile(fh_path) else None,
                "micrograph_img_path": micro if micro and os.path.isfile(micro) else None,
            }
            children.append(child)

        nodes.append(
            {
                "gs_dir": gs_dir,
                "name": gs_name,
                "epu": gs_epu,
                "index": gs_index,
                "latest_img_path": gs_img_path if gs_img_path and os.path.isfile(gs_img_path) else None,
                "support_img_path": support_img_path if support_img_path and os.path.isfile(support_img_path) else None,
                "nonsupport_img_path": nonsupport_img_path if nonsupport_img_path and os.path.isfile(nonsupport_img_path) else None,
                "children": children,
            }
        )

    try:
        nodes.sort(key=lambda n: (n.get("index") is None, n.get("index", 10**9), n.get("name", "")))
    except Exception:
        pass

    return nodes
