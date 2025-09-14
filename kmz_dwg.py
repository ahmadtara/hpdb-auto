import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
import math
from shapely.geometry import Point, LineString

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

target_folders = {
    'FDT', 'FAT', 'HP COVER', 'HP UNCOVER', 'NEW POLE 7-3', 'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY', 'DISTRIBUTION CABLE', 'SLING WIRE', 'KOTAK', 'JALAN'
}

def extract_kmz(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kmz_file.extractall(extract_dir)
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
            if point_coord is not None:
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
            if line_coord is not None:
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
            if poly_coord is not None:
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

def latlon_to_xy(lat, lon):
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

def segment_angle(p1, p2):
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    return math.degrees(math.atan2(dy, dx))

def longest_segment_angle(coords):
    max_len = 0
    angle = 0
    for i in range(len(coords) - 1):
        x1, y1 = coords[i]
        x2, y2 = coords[i + 1]
        seg_len = math.hypot(x2 - x1, y2 - y1)
        if seg_len > max_len:
            max_len = seg_len
            angle = segment_angle((x1, y1), (x2, y2))
    return angle

def cluster_hp_rotation(hp_points, homepass_lines, cables):
    if not hp_points:
        return {}
    # Cari arah utama (longest segment dari polyline terdekat)
    base_angle = None
    if homepass_lines:
        base_line = max(homepass_lines, key=lambda l: l.length)
        base_angle = longest_segment_angle(list(base_line.coords))
    elif cables:
        base_line = max([LineString(c['xy_path']) for c in cables if 'xy_path' in c], key=lambda l: l.length)
        base_angle = longest_segment_angle(list(base_line.coords))
    else:
        base_angle = 0.0
    # Semua HP di cluster pakai base_angle yang sama
    return {id(pt): base_angle for pt in hp_points}

def draw_to_template(classified, template_path):
    doc = ezdxf.readfile(template_path)
    msp = doc.modelspace()

    # --- Kumpulkan semua koordinat ---
    all_xy = []
    for layer_name, cat_items in classified.items():
        for obj in cat_items:
            if obj['type'] == 'point':
                all_xy.append(latlon_to_xy(obj['latitude'], obj['longitude']))
            elif obj['type'] == 'path':
                all_xy.extend([latlon_to_xy(lat, lon) for lat, lon in obj['coords']])

    if not all_xy:
        st.error("❌ Tidak ada data dari KMZ!")
        return None

    shifted_all, (cx, cy) = apply_offset(all_xy)
    idx = 0
    for layer_name, cat_items in classified.items():
        for obj in cat_items:
            if obj['type'] == 'point':
                obj['xy'] = shifted_all[idx]
                idx += 1
            elif obj['type'] == 'path':
                obj['xy_path'] = shifted_all[idx: idx + len(obj['coords'])]
                idx += len(obj['coords'])

    layer_mapping = {
        "BOUNDARY": "FAT AREA",
        "DISTRIBUTION_CABLE": "FO 36 CORE",  # default mapping
        "SLING_WIRE": "STRAND UG",
        "KOTAK": "GARIS HOMEPASS",
        "JALAN": "JALAN"
    }

    homepass_lines = []
    for obj in classified.get("KOTAK", []):
        if obj['type'] == 'path' and len(obj['xy_path']) >= 2:
            homepass_lines.append(LineString(obj['xy_path']))

    cables = classified.get("DISTRIBUTION_CABLE", [])
    hp_points = classified.get("HP_COVER", []) + classified.get("HP_UNCOVER", [])
    rotation_map = cluster_hp_rotation(hp_points, homepass_lines, cables)

    # --- Gambar objek ---
    for layer_name, cat_items in classified.items():
        true_layer = layer_mapping.get(layer_name, layer_name)
        for obj in cat_items:
            if obj['type'] != 'point':
                if len(obj['xy_path']) >= 2:
                    msp.add_lwpolyline(obj['xy_path'], dxfattribs={"layer": true_layer})
                elif len(obj['xy_path']) == 1:
                    msp.add_circle(center=obj['xy_path'][0], radius=0.5, dxfattribs={"layer": true_layer})
                continue

            x, y = obj['xy']
            rotation = rotation_map.get(id(obj), 0.0)

            if layer_name == "HP_COVER":
                msp.add_text(obj["name"], dxfattribs={
                    "height": 6,
                    "layer": "FEATURE_LABEL",
                    "color": 6,
                    "insert": (x, y),
                    "rotation": rotation
                })
                continue

            elif layer_name == "HP_UNCOVER":
                msp.add_text(obj["name"], dxfattribs={
                    "height": 3.0,
                    "layer": "FEATURE_LABEL",
                    "color": 7,
                    "insert": (x, y),
                    "rotation": rotation
                })
                continue

            # fallback lain untuk FAT/FDT/POLE
            msp.add_circle(center=(x, y), radius=2, dxfattribs={"layer": true_layer})
            msp.add_text(obj["name"], dxfattribs={
                "height": 3.0,
                "layer": "FEATURE_LABEL",
                "color": 1,
                "insert": (x + 2, y)
            })

    return doc

def run_kmz_to_dwg():
    st.title("🏗️ KMZ → AUTOCAD ")
    uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
    uploaded_template = st.file_uploader("📀 Upload Template DXF", type=["dxf"])

    if uploaded_kmz and uploaded_template:
        extract_dir = "temp_kmz"
        os.makedirs(extract_dir, exist_ok=True)
        output_dxf = "converted_output.dxf"

        with open("template_ref.dxf", "wb") as f:
            f.write(uploaded_template.read())

        with st.spinner("🔍 Memproses data..."):
            try:
                kml_path = extract_kmz(uploaded_kmz, extract_dir)
                items = parse_kml(kml_path)
                classified = classify_items(items)
                updated_doc = draw_to_template(classified, "template_ref.dxf")
                if updated_doc:
                    updated_doc.saveas(output_dxf)

                if os.path.exists(output_dxf):
                    st.success("✅ Konversi berhasil! DXF sudah dibuat.")
                    with open(output_dxf, "rb") as f:
                        st.download_button("⬇️ Download DXF", f, file_name="output_from_kmz.dxf")
            except Exception as e:
                st.error(f"❌ Gagal memproses: {e}")
