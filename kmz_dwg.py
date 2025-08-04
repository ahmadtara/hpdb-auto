import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
from shapely.geometry import LineString, Point
import math

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

target_folders = {
    'FDT', 'FAT', 'HP COVER', 'HP UNCOVER', 'NEW POLE 7-3', 'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY', 'DISTRIBUTION CABLE', 'SLING WIRE', 'KOTAK'
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
        "BOUNDARY", "DISTRIBUTION_CABLE", "SLING_WIRE", "KOTAK"
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
        else:
            classified["POLE"].append(it)
    return classified

def get_nearest_angle(point, lines):
    min_dist = float('inf')
    nearest_line = None
    for line in lines:
        dist = line.distance(point)
        if dist < min_dist:
            min_dist = dist
            nearest_line = line
    if nearest_line is None:
        return 0
    x1, y1 = nearest_line.coords[0]
    x2, y2 = nearest_line.coords[-1]
    angle_rad = math.atan2(y2 - y1, x2 - x1)
    angle_deg = math.degrees(angle_rad)
    return angle_deg

def draw_to_template(classified, template_path):
    doc = ezdxf.readfile(template_path)
    msp = doc.modelspace()

    distribution_lines = []
    for obj in classified.get("DISTRIBUTION_CABLE", []):
        if obj['type'] == 'path':
            coords = [latlon_to_xy(lat, lon) for lat, lon in obj['coords']]
            line = LineString(coords)
            distribution_lines.append(line)

    placed_text_positions = []
    min_distance = 2.0
    vertical_step = 1.5

    for layer_name, items in classified.items():
        for obj in items:
            if obj['type'] == 'point':
                x, y = latlon_to_xy(obj['latitude'], obj['longitude'])
                point = Point(x, y)
                rotation = get_nearest_angle(point, distribution_lines)
                adjusted_y = y
                while any(Point(x, adjusted_y).distance(p) < min_distance for p in placed_text_positions):
                    adjusted_y += vertical_step
                placed_text_positions.append(Point(x, adjusted_y))
                msp.add_text(
                    obj["name"],
                    dxfattribs={
                        "height": 6.0,
                        "layer": "FEATURE_LABEL",
                        "color": 1,
                        "insert": (x, adjusted_y),
                        "rotation": rotation
                    }
                )

    return doc

def run_kmz_to_dwg():
    st.title("\U0001F3D7️ KMZ → AUTOCAD ")
    st.markdown("""
<h2>👋 Hai, <span style='color:#0A84FF'>bro</span></h2>
✅ <span style='font-weight:bold;'>CATATAN PENTING :</span><br>
1️⃣ <span style='color:#FF6B6B;'>PASTIKAN KMZ SESUAI TEMPLATE</span>.<br>
2️⃣ FOLDER KOTAK HARUS DIBUAT MANUAL DULU DARI DALAM KMZ <code>Agar kotak rumah otoatis didalam kode</code><br><br>
""", unsafe_allow_html=True)

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
