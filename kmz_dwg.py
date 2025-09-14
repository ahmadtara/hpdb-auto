import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
import math
from shapely.geometry import Point, LineString
from shapely.ops import nearest_points
from statistics import mean

# ----------------------------
# Config / Transformer
# ----------------------------
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

target_folders = {
    'FDT', 'FAT', 'HP COVER', 'HP UNCOVER', 'NEW POLE 7-3', 'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY', 'DISTRIBUTION CABLE', 'SLING WIRE', 'KOTAK', 'JALAN'
}

# ----------------------------
# KML / KMZ parsing
# ----------------------------
def extract_kmz(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kmz_file.extractall(extract_dir)
    # default produced file usually doc.kml
    return os.path.join(extract_dir, "doc.kml")

def parse_kml(kml_path):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    with open(kml_path, 'rb') as f:
        tree = ET.parse(f)
    root = tree.getroot()
    folders = root.findall('.//kml:Folder', ns)
    items = []
    for folder in folders:
        folder_name_tag = folder.find('kml:name', ns)
        if folder_name_tag is None:
            continue
        folder_name = folder_name_tag.text.strip().upper()
        if folder_name not in target_folders:
            continue
        placemarks = folder.findall('.//kml:Placemark', ns)
        for pm in placemarks:
            name = pm.find('kml:name', ns)
            name_text = name.text.strip() if name is not None else ""

            point_coord = pm.find('.//kml:Point/kml:coordinates', ns)
            if point_coord is not None and point_coord.text and point_coord.text.strip():
                lon, lat, *_ = point_coord.text.strip().split(',')
                items.append({
                    'type': 'point',
                    'name': name_text,
                    'latitude': float(lat),
                    'longitude': float(lon),
                    'folder': folder_name
                })
                continue

            line_coord = pm.find('.//kml:LineString/kml:coordinates', ns)
            if line_coord is not None and line_coord.text and line_coord.text.strip():
                coords = []
                for c in line_coord.text.strip().split():
                    lon, lat, *_ = c.split(',')
                    coords.append((float(lat), float(lon)))
                items.append({
                    'type': 'path',
                    'name': name_text,
                    'coords': coords,
                    'folder': folder_name
                })
                continue

            poly_coord = pm.find('.//kml:Polygon//kml:coordinates', ns)
            if poly_coord is not None and poly_coord.text and poly_coord.text.strip():
                coords = []
                for c in poly_coord.text.strip().split():
                    lon, lat, *_ = c.split(',')
                    coords.append((float(lat), float(lon)))
                items.append({
                    'type': 'path',
                    'name': name_text,
                    'coords': coords,
                    'folder': folder_name
                })
    return items

# ----------------------------
# Utilities
# ----------------------------
def latlon_to_xy(lat, lon):
    # transformer.transform(longitude, latitude)
    return transformer.transform(lon, lat)

def apply_offset(points_xy):
    xs = [x for x, y in points_xy]
    ys = [y for x, y in points_xy]
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    return [(x - cx, y - cy) for x, y in points_xy], (cx, cy)

def classify_items(items):
    classified = {name: [] for name in [
        "FDT", "FAT", "HP_COVER", "HP_UNCOVER", "NEW_POLE", "EXISTING_POLE", "POLE",
        "BOUNDARY", "DISTRIBUTION_CABLE", "SLING_WIRE", "KOTAK", "JALAN"
    ]}
    for it in items:
        folder = it['folder']
        if "FDT" in folder:
            classified["FDT"].append(it)
        elif "FAT" in folder and folder != "FAT AREA":
            classified["FAT"].append(it)
        elif "HP COVER" in folder:
            classified["HP_COVER"].append(it)
        elif "HP UNCOVER" in folder:
            classified["HP_UNCOVER"].append(it)
        elif "NEW POLE" in folder:
            classified["NEW_POLE"].append(it)
        elif "EXISTING" in folder or "EMR" in folder:
            classified["EXISTING_POLE"].append(it)
        elif "BOUNDARY" in folder:
            classified["BOUNDARY"].append(it)
        elif "DISTRIBUTION CABLE" in folder:
            classified["DISTRIBUTION_CABLE"].append(it)
        elif "SLING WIRE" in folder:
            classified["SLING_WIRE"].append(it)
        elif "KOTAK" in folder:
            classified["KOTAK"].append(it)
        elif "JALAN" in folder:
            classified["JALAN"].append(it)
        else:
            classified["POLE"].append(it)
    return classified

def segment_angle_xy(p1, p2):
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    ang = math.degrees(math.atan2(dy, dx))
    # normalize to [-180,180)
    if ang <= -180:
        ang += 360
    if ang > 180:
        ang -= 360
    return ang

# ----------------------------
# Cable helpers (work in projected XY)
# ----------------------------
def build_cable_lines(classified):
    """Return list of dicts: {'orig': item, 'xy_path': [(x,y),...], 'line': LineString}"""
    cables = []
    for item in classified.get("DISTRIBUTION_CABLE", []):
        if item['type'] != 'path' or not item.get('coords'):
            continue
        xy = [latlon_to_xy(lat, lon) for lat, lon in item['coords']]
        if len(xy) >= 2:
            cables.append({'orig': item, 'xy_path': xy, 'line': LineString(xy)})
    return cables

def nearest_segment_angle_with_minlen(pt_xy, cable_line: LineString, min_seg_len):
    """
    For a given cable LineString (in XY), find the nearest segment to pt_xy that has length >= min_seg_len.
    If none found, return angle of the nearest segment regardless of length (fallback).
    """
    coords = list(cable_line.coords)
    px, py = pt_xy
    best_dist = float("inf")
    best_angle = None
    # first try to find nearest segment with seg_len >= min_seg_len
    for i in range(len(coords) - 1):
        p1 = coords[i]
        p2 = coords[i + 1]
        seg = LineString([p1, p2])
        seg_len = seg.length
        if seg_len < min_seg_len:
            continue
        dist = seg.distance(Point(px, py))
        if dist < best_dist:
            best_dist = dist
            best_angle = segment_angle_xy(p1, p2)
    if best_angle is not None:
        return best_angle
    # fallback: nearest segment ignoring min length
    best_dist = float("inf")
    for i in range(len(coords) - 1):
        p1 = coords[i]
        p2 = coords[i + 1]
        seg = LineString([p1, p2])
        dist = seg.distance(Point(px, py))
        if dist < best_dist:
            best_dist = dist
            best_angle = segment_angle_xy(p1, p2)
    return best_angle if best_angle is not None else 0.0

# ----------------------------
# Group HP by projection on cable and along-distance
# ----------------------------
def group_hp_by_cable_and_along(hp_xy_list, cables, max_gap_along=20.0):
    """
    hp_xy_list: list of dicts {'obj': original_obj, 'xy':(x,y)}
    cables: list produced by build_cable_lines
    Returns list of groups where each group is dict:
      {'cable_idx': int, 'indices': [hp indices], 'along_vals': [float], 'representative_xy': (x,y)}
    Strategy:
      - For each HP, find nearest cable index and project onto it (line.project)
      - For each cable, sort HP by along-distance and split into groups when gap > max_gap_along
    """
    # prepare structure per cable
    per_cable_hp = {i: [] for i in range(len(cables))}
    hp_meta = []  # store cable_idx and along
    for idx, hp in enumerate(hp_xy_list):
        pt = Point(hp['xy'])
        best_c = None
        best_dist = float("inf")
        best_along = None
        for c_idx, c in enumerate(cables):
            line = c['line']
            d = line.distance(pt)
            if d < best_dist:
                best_dist = d
                best_c = c_idx
                best_along = line.project(pt)  # distance along line (units same as projected coords)
        if best_c is None:
            best_c = 0
            best_along = 0.0
        per_cable_hp[best_c].append((idx, best_along))
        hp_meta.append({'cable_idx': best_c, 'along': best_along, 'dist_to_cable': best_dist})

    groups = []
    # for each cable, sort and split
    for c_idx, hp_list in per_cable_hp.items():
        if not hp_list:
            continue
        hp_list.sort(key=lambda x: x[1])  # sort by along
        # split into segments where gap > max_gap_along
        current_group = [hp_list[0][0]]
        last_along = hp_list[0][1]
        group_alongs = [last_along]
        for idx_along in hp_list[1:]:
            idx_i, along_i = idx_along
            if abs(along_i - last_along) > max_gap_along:
                # close current
                groups.append({
                    'cable_idx': c_idx,
                    'indices': current_group.copy(),
                    'along_vals': group_alongs.copy()
                })
                current_group = [idx_i]
                group_alongs = [along_i]
            else:
                current_group.append(idx_i)
                group_alongs.append(along_i)
            last_along = along_i
        # append last
        groups.append({
            'cable_idx': c_idx,
            'indices': current_group.copy(),
            'along_vals': group_alongs.copy()
        })

    return groups, hp_meta

# ----------------------------
# offset helper for text placement
# ----------------------------
def offset_xy(x, y, angle_deg, dist=2.5):
    """Geser koordinat (x,y) menjauh dari garis sesuai sudut (tegak lurus)."""
    # angle_deg adalah arah segmen (deg). Geser tegak lurus (angle + 90)
    angle_rad = math.radians(angle_deg + 90)
    return (x + dist * math.cos(angle_rad), y + dist * math.sin(angle_rad))

# ----------------------------
# Main DXF building + rotation assignment
# ----------------------------
def build_dxf_with_smart_hp(classified, template_path, output_path,
                            min_seg_len=15.0, max_gap_along=20.0):
    # load template to copy blocks / ensure layers, if not exists create new doc
    doc = ezdxf.readfile(template_path) if os.path.exists(template_path) else ezdxf.new('R2010')
    msp = doc.modelspace()

    # Prepare all XY coords and map them back
    all_xy = []
    for _, cat_items in classified.items():
        for obj in cat_items:
            if obj['type'] == 'point':
                all_xy.append(latlon_to_xy(obj['latitude'], obj['longitude']))
            elif obj['type'] == 'path':
                all_xy.extend([latlon_to_xy(lat, lon) for lat, lon in obj['coords']])

    if not all_xy:
        st.error("❌ Tidak ada data dari KMZ!")
        return None

    shifted_all, _ = apply_offset(all_xy)

    # assign shifted coords back
    idx = 0
    for _, cat_items in classified.items():
        for obj in cat_items:
            if obj['type'] == 'point':
                obj['xy'] = shifted_all[idx]
                idx += 1
            elif obj['type'] == 'path':
                obj['xy_path'] = shifted_all[idx: idx + len(obj['coords'])]
                idx += len(obj['coords'])

    # layer mapping
    layer_mapping = {
        "BOUNDARY": "FAT AREA",
        "DISTRIBUTION_CABLE": "FO 36 CORE",
        "SLING_WIRE": "STRAND UG",
        "KOTAK": "GARIS HOMEPASS",
        "JALAN": "JALAN"
    }

    # build cables list (XY)
    cables = build_cable_lines(classified)  # each has 'xy_path' and 'line'

    # prepare HP list in order of appearance (we'll map indices)
    hp_items = []
    for obj in classified.get("HP_COVER", []) + classified.get("HP_UNCOVER", []):
        if 'xy' in obj:
            hp_items.append({'obj': obj, 'xy': obj['xy']})

    # group HP by cable+along and split into deret by gap along
    groups, hp_meta = group_hp_by_cable_and_along(hp_items, cables, max_gap_along=max_gap_along)

    # for each group, determine rotation:
    # choose representative point = median along projected point, then find nearest long seg (>=min_seg_len)
    group_rotation_map = {}  # group index -> angle
    for g_idx, group in enumerate(groups):
        c_idx = group['cable_idx']
        cable = cables[c_idx]
        line = cable['line']
        alongs = group['along_vals']
        # representative along: median
        rep_along = sorted(alongs)[len(alongs)//2]
        rep_point = line.interpolate(rep_along)
        rep_xy = (rep_point.x, rep_point.y)
        # try to get nearest segment angle with min length
        angle = nearest_segment_angle_with_minlen(rep_xy, line, min_seg_len)
        group_rotation_map[g_idx] = angle
        # also assign rotation to each hp in group (store on hp_items by index)
        for hp_idx in group['indices']:
            hp_items[hp_idx]['rotation'] = angle
            hp_items[hp_idx]['group_idx'] = g_idx

    # any HP not assigned (edge case) -> fallback individually
    for idx_i, hp in enumerate(hp_items):
        if 'rotation' not in hp:
            # find nearest cable and segment angle fallback
            best_angle = 0.0
            best_dist = float("inf")
            for c in cables:
                line = c['line']
                dist = line.distance(Point(hp['xy']))
                if dist < best_dist:
                    best_dist = dist
                    best_angle = nearest_segment_angle_with_minlen(hp['xy'], line, min_seg_len)
            hp['rotation'] = best_angle

    # ----------------------------
    # Assign rotation for POLES (NEW & EXISTING), FAT, FDT
    # ----------------------------
    pole_items = []
    for obj in classified.get("NEW_POLE", []) + classified.get("EXISTING_POLE", []):
        if 'xy' in obj:
            best_angle = 0.0
            best_dist = float("inf")
            best_line = None
            for c in cables:
                line = c['line']
                dist = line.distance(Point(obj['xy']))
                if dist < best_dist:
                    best_dist = dist
                    best_line = line
            if best_line is not None:
                best_angle = nearest_segment_angle_with_minlen(obj['xy'], best_line, min_seg_len)
            pole_items.append({'obj': obj, 'rotation': best_angle})

    fat_items = []
    for obj in classified.get("FAT", []):
        if 'xy' in obj:
            best_angle = 0.0
            best_dist = float("inf")
            best_line = None
            for c in cables:
                line = c['line']
                dist = line.distance(Point(obj['xy']))
                if dist < best_dist:
                    best_dist = dist
                    best_line = line
            if best_line is not None:
                best_angle = nearest_segment_angle_with_minlen(obj['xy'], best_line, min_seg_len)
            fat_items.append({'obj': obj, 'rotation': best_angle})

    fdt_items = []
    for obj in classified.get("FDT", []):
        if 'xy' in obj:
            best_angle = 0.0
            best_dist = float("inf")
            best_line = None
            for c in cables:
                line = c['line']
                dist = line.distance(Point(obj['xy']))
                if dist < best_dist:
                    best_dist = dist
                    best_line = line
            if best_line is not None:
                best_angle = nearest_segment_angle_with_minlen(obj['xy'], best_line, min_seg_len)
            fdt_items.append({'obj': obj, 'rotation': best_angle})

    # (Optional) smoothing etc could be added here

    # --- DRAW: first draw polylines (cables + others) ---
    for layer_name, cat_items in classified.items():
        true_layer = layer_mapping.get(layer_name, layer_name)
        for obj in cat_items:
            if obj['type'] == 'path':
                if len(obj.get('xy_path', [])) >= 2:
                    msp.add_lwpolyline(obj['xy_path'], dxfattribs={"layer": true_layer})
                elif len(obj.get('xy_path', [])) == 1:
                    msp.add_circle(center=obj['xy_path'][0], radius=0.5, dxfattribs={"layer": true_layer})

    # draw non-HP points except POLE/FAT/FDT (we handle those specifically)
    for layer_name, cat_items in classified.items():
        true_layer = layer_mapping.get(layer_name, layer_name)
        if layer_name in ("HP_COVER", "HP_UNCOVER", "NEW_POLE", "EXISTING_POLE", "FAT", "FDT"):
            continue  # skip here, draw later with rotation/offset
        for obj in cat_items:
            if obj['type'] == 'point':
                x, y = obj['xy']
                # draw generic symbol (circle) and text
                msp.add_circle(center=(x, y), radius=2, dxfattribs={"layer": true_layer})
                msp.add_text(obj.get("name", ""), dxfattribs={
                    "height": 3.0,
                    "layer": "FEATURE_LABEL",
                    "color": 1,
                    "insert": (x + 2, y)
                })

    # draw POLE with rotation + offset
    for pole in pole_items:
        x, y = pole['obj']['xy']
        rot = pole['rotation']
        name = pole['obj'].get('name', '')
        # draw pole symbol
        msp.add_circle(center=(x, y), radius=2, dxfattribs={"layer": "POLE"})
        # offset text slightly away from line to avoid overlap
        tx, ty = offset_xy(x, y, rot, dist=3.5)
        msp.add_text(name, dxfattribs={
            "height": 3.0,
            "layer": "FEATURE_LABEL",
            "color": 2,
            "insert": (tx, ty),
            "rotation": rot
        })

    # draw FAT texts with rotation + offset
    for fat in fat_items:
        x, y = fat['obj']['xy']
        rot = fat['rotation']
        name = fat['obj'].get('name', '')
        # draw marker
        msp.add_circle(center=(x, y), radius=1.8, dxfattribs={"layer": "FAT"})
        tx, ty = offset_xy(x, y, rot, dist=4.0)
        msp.add_text(name, dxfattribs={
            "height": 4.0,
            "layer": "FEATURE_LABEL",
            "color": 3,
            "insert": (tx, ty),
            "rotation": rot
        })

    # draw FDT texts with rotation + offset
    for fdt in fdt_items:
        x, y = fdt['obj']['xy']
        rot = fdt['rotation']
        name = fdt['obj'].get('name', '')
        # draw marker
        msp.add_circle(center=(x, y), radius=1.8, dxfattribs={"layer": "FDT"})
        tx, ty = offset_xy(x, y, rot, dist=4.0)
        msp.add_text(name, dxfattribs={
            "height": 3.5,
            "layer": "FEATURE_LABEL",
            "color": 4,
            "insert": (tx, ty),
            "rotation": rot
        })

    # draw HP texts with assigned rotations (already computed)
    for hp in hp_items:
        x, y = hp['xy']
        rot = hp.get('rotation', 0.0)
        name = hp['obj'].get('name', '')
        layername = hp['obj']['folder']
        if "HP COVER" in layername:
            tx, ty = offset_xy(x, y, rot, dist=2.5)
            msp.add_text(name, dxfattribs={
                "height": 6,
                "layer": "FEATURE_LABEL",
                "color": 6,
                "insert": (tx, ty),
                "rotation": rot
            })
        else:
            tx, ty = offset_xy(x, y, rot, dist=2.0)
            msp.add_text(name, dxfattribs={
                "height": 3.0,
                "layer": "FEATURE_LABEL",
                "color": 7,
                "insert": (tx, ty),
                "rotation": rot
            })

    # save
    doc.saveas(output_path)
    return output_path

# ----------------------------
# Streamlit UI
# ----------------------------
def run_kmz_to_dwg():
    st.title("🏗️ KMZ → AUTOCAD (Smart HP rotation)")

    st.markdown("""
    - Rotasi HP mengikuti DISTRIBUTION CABLE.
    - Segmen kabel pendek (< min seg len) diabaikan agar HP di belokan kecil tidak belok sendiri.
    - HP digroup berdasarkan proyeksi sepanjang kabel; tiap group (deret) pakai rotasi yang sama.
    - POLE, FAT & FDT juga diberi rotasi mengikuti kabel terdekat dan teksnya digeser sedikit agar tidak menempel ke garis/block.
    """)

    uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
    uploaded_template = st.file_uploader("📀 Upload Template DXF (optional)", type=["dxf"])

    st.sidebar.header("Rotation parameters")
    min_seg_len = st.sidebar.slider("Min seg length to consider (m)", min_value=5.0, max_value=100.0, value=15.0, step=1.0)
    max_gap_along = st.sidebar.slider("Max along-gap to split deret (m)", min_value=5.0, max_value=200.0, value=20.0, step=1.0)

    if uploaded_kmz:
        extract_dir = "temp_kmz"
        os.makedirs(extract_dir, exist_ok=True)
        with open("temp_upload.kmz", "wb") as f:
            f.write(uploaded_kmz.read())
        kml_path = extract_kmz("temp_upload.kmz", extract_dir)
        items = parse_kml(kml_path)
        classified = classify_items(items)

        # compute output
        template_path = "template_ref.dxf"
        if uploaded_template:
            with open(template_path, "wb") as f:
                f.write(uploaded_template.read())
        else:
            # if template not provided, ensure we have something (will use new doc inside function)
            template_path = template_path if os.path.exists(template_path) else ""

        output_dxf = "converted_output.dxf"
        try:
            result = build_dxf_with_smart_hp(classified, template_path, output_dxf,
                                            min_seg_len=min_seg_len, max_gap_along=max_gap_along)
            if result and os.path.exists(result):
                st.success("✅ DXF berhasil dibuat.")
                with open(result, "rb") as f:
                    st.download_button("⬇️ Download DXF", f, file_name=os.path.basename(result))
            else:
                st.error("❌ Gagal membuat DXF.")
        except Exception as e:
            st.error(f"❌ Error saat memproses: {e}")

if __name__ == "__main__":
    run_kmz_to_dwg()
