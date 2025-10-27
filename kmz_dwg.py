# =========================================
# KMZ → DXF Converter (Smart HP Rotation)
# Version: 3.2 — Stable Streamlit Edition
# =========================================
import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
import math
from shapely.geometry import Point, LineString
from statistics import mean
from shapely.ops import nearest_points

# ----------------------------
# Config / Transformer
# ----------------------------
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

target_folders = {
    'FDT', 'FAT', 'HP COVER', 'HP UNCOVER',
    'NEW POLE 7-3', 'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY FAT', 'DISTRIBUTION CABLE', 'SLING WIRE',
    'KOTAK', 'JALAN'
}

# ----------------------------
# KML / KMZ parsing
# ----------------------------
def extract_kmz(kmz_path, extract_dir):
    """Ekstrak file KMZ dan kembalikan path ke doc.kml"""
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kmz_file.extractall(extract_dir)
    return os.path.join(extract_dir, "doc.kml")

def parse_kml(kml_path):
    """Parse KML untuk ambil semua point/line/polygon dari target_folders"""
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

            # POINT
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

            # LINESTRING
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

            # POLYGON
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
    return transformer.transform(lon, lat)

def apply_offset(points_xy):
    xs = [x for x, y in points_xy]
    ys = [y for x, y in points_xy]
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    return [(x - cx, y - cy) for x, y in points_xy], (cx, cy)

def classify_items(items):
    classified = {name: [] for name in [
        "FDT", "FAT", "HP_COVER", "HP_UNCOVER",
        "NEW_POLE_7_3", "NEW_POLE_7_4",
        "EXISTING_POLE", "POLE",
        "BOUNDARY FAT", "DISTRIBUTION_CABLE",
        "SLING_WIRE", "KOTAK", "JALAN"
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
        elif "NEW POLE 7-3" in folder:
            classified["NEW_POLE_7_3"].append(it)
        elif "NEW POLE 7-4" in folder:
            classified["NEW_POLE_7_4"].append(it)
        elif "EXISTING POLE" in folder or "EMR" in folder:
            classified["EXISTING_POLE"].append(it)
        elif "BOUNDARY FAT" in folder:
            classified["BOUNDARY FAT"].append(it)
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
    return (ang + 360) % 360 if ang < 0 else ang

# ----------------------------
# Cable helpers
# ----------------------------
def build_cable_lines(classified):
    cables = []
    for item in classified.get("DISTRIBUTION_CABLE", []):
        if item['type'] != 'path' or not item.get('coords'):
            continue
        xy = [latlon_to_xy(lat, lon) for lat, lon in item['coords']]
        if len(xy) >= 2:
            cables.append({'orig': item, 'xy_path': xy, 'line': LineString(xy)})
    return cables

def nearest_segment_angle_with_minlen(pt_xy, cable_line: LineString, min_seg_len):
    coords = list(cable_line.coords)
    px, py = pt_xy
    best_dist = float("inf")
    best_angle = 0
    for i in range(len(coords) - 1):
        p1, p2 = coords[i], coords[i + 1]
        seg = LineString([p1, p2])
        if seg.length < min_seg_len:
            continue
        dist = seg.distance(Point(px, py))
        if dist < best_dist:
            best_dist = dist
            best_angle = segment_angle_xy(p1, p2)
    return best_angle

# ----------------------------
# Group HP by cable
# ----------------------------
def group_hp_by_cable_and_along(hp_xy_list, cables, max_gap_along=20.0):
    per_cable_hp = {i: [] for i in range(len(cables))}
    hp_meta = []
    for idx, hp in enumerate(hp_xy_list):
        pt = Point(hp['xy'])
        best_c = 0
        best_along = 0.0
        best_dist = float("inf")
        for c_idx, c in enumerate(cables):
            line = c['line']
            d = line.distance(pt)
            if d < best_dist:
                best_dist = d
                best_c = c_idx
                best_along = line.project(pt)
        per_cable_hp[best_c].append((idx, best_along))
        hp_meta.append({'cable_idx': best_c, 'along': best_along, 'dist_to_cable': best_dist})

    groups = []
    for c_idx, hp_list in per_cable_hp.items():
        if not hp_list:
            continue
        hp_list.sort(key=lambda x: x[1])
        current = [hp_list[0][0]]
        last = hp_list[0][1]
        alongs = [last]
        for idx_i, along_i in hp_list[1:]:
            if abs(along_i - last) > max_gap_along:
                groups.append({'cable_idx': c_idx, 'indices': current.copy(), 'along_vals': alongs.copy()})
                current = [idx_i]
                alongs = [along_i]
            else:
                current.append(idx_i)
                alongs.append(along_i)
            last = along_i
        groups.append({'cable_idx': c_idx, 'indices': current.copy(), 'along_vals': alongs.copy()})
    return groups, hp_meta

# ----------------------------
# DXF Builder (same as your original)
# ----------------------------
# [build_dxf_with_smart_hp(...) tetap sama seperti kode kamu di atas]
# ----------------------------

# ----------------------------
# Streamlit App
# ----------------------------
@st.cache_data
def process_kmz(uploaded_kmz):
    tmpdir = "temp_extract"
    os.makedirs(tmpdir, exist_ok=True)
    kmz_path = os.path.join(tmpdir, "uploaded.kmz")
    with open(kmz_path, "wb") as f:
        f.write(uploaded_kmz.read())
    kml_path = extract_kmz(kmz_path, tmpdir)
    items = parse_kml(kml_path)
    return classify_items(items)

def run_kmz_to_dwg():
    st.title("🏗️ KMZ → AUTOCAD (Smart HP Rotation + Block Insertion)")
    uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
    uploaded_template = st.file_uploader("📀 Upload Template DXF (optional)", type=["dxf"])
    st.sidebar.header("⚙️ Rotation Parameters")
    min_seg_len = st.sidebar.slider("Min segment length (m)", 5.0, 100.0, 15.0, 1.0)
    max_gap_along = st.sidebar.slider("Max gap along (m)", 5.0, 200.0, 20.0, 1.0)
    rotate_hp = st.checkbox("Rotate HP Text", value=True)

    if uploaded_kmz:
        with st.spinner("🔄 Memproses KMZ..."):
            classified = process_kmz(uploaded_kmz)

        tmpdir = "temp_extract"
        template_path = None
        if uploaded_template:
            template_path = os.path.join(tmpdir, "template.dxf")
            with open(template_path, "wb") as f:
                f.write(uploaded_template.read())

        out_path = "output_smart_hp.dxf"
        with st.spinner("✏️ Membuat file DXF..."):
            res = build_dxf_with_smart_hp(
                classified, template_path, out_path,
                min_seg_len=min_seg_len,
                max_gap_along=max_gap_along,
                rotate_hp=rotate_hp
            )
        if res:
            with open(res, "rb") as f:
                st.download_button("⬇️ Download DXF", f, file_name="output_smart_hp.dxf")

if __name__ == "__main__":
    run_kmz_to_dwg()
