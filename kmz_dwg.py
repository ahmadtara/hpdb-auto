import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
import math
from shapely.geometry import Point, LineString

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
    return math.degrees(math.atan2(dy, dx))

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
            cables.append({'xy_path': xy, 'line': LineString(xy)})
    return cables

def nearest_segment_angle(pt_xy, cable_line: LineString):
    coords = list(cable_line.coords)
    px, py = pt_xy
    best_dist = float("inf")
    best_angle = 0.0
    for i in range(len(coords) - 1):
        p1, p2 = coords[i], coords[i+1]
        seg = LineString([p1, p2])
        dist = seg.distance(Point(px, py))
        if dist < best_dist:
            best_dist = dist
            best_angle = segment_angle_xy(p1, p2)
    return best_angle

def get_rotation_for_point(pt_xy, cables):
    best_dist = float("inf")
    best_angle = 0.0
    for c in cables:
        dist = c['line'].distance(Point(pt_xy))
        if dist < best_dist:
            best_dist = dist
            best_angle = nearest_segment_angle(pt_xy, c['line'])
    return best_angle

# ----------------------------
# DXF Builder
# ----------------------------
def build_dxf_with_rotation(classified, template_path, output_path):
    if template_path and os.path.exists(template_path):
        doc = ezdxf.readfile(template_path)
    else:
        doc = ezdxf.new('R2010')
    msp = doc.modelspace()

    # kumpulkan semua koordinat lalu shift
    all_xy = []
    for _, cat_items in classified.items():
        for obj in cat_items:
            if obj['type'] == 'point':
                all_xy.append(latlon_to_xy(obj['latitude'], obj['longitude']))
            elif obj['type'] == 'path':
                all_xy.extend([latlon_to_xy(lat, lon) for lat, lon in obj['coords']])
    shifted_all, _ = apply_offset(all_xy)

    idx = 0
    for _, cat_items in classified.items():
        for obj in cat_items:
            if obj['type'] == 'point':
                obj['xy'] = shifted_all[idx]; idx += 1
            elif obj['type'] == 'path':
                obj['xy_path'] = shifted_all[idx: idx + len(obj['coords'])]; idx += len(obj['coords'])

    # bangun kabel
    cables = build_cable_lines(classified)

    # gambar point (FDT, FAT, POLE)
    for layer_name, cat_items in classified.items():
        if layer_name not in ["FDT", "FAT", "NEW_POLE", "EXISTING_POLE"]:
            continue
        for obj in cat_items:
            if obj['type'] != "point":
                continue
            x, y = obj['xy']
            angle = get_rotation_for_point((x, y), cables)

            # block scale kecil
            if layer_name in ["FDT", "FAT", "NEW_POLE", "EXISTING_POLE"]:
                scale = 0.0025
            else:
                scale = 1.0

            # insert block (fallback circle)
            try:
                msp.add_blockref(
                    layer_name, (x, y),
                    dxfattribs={
                        "xscale": scale,
                        "yscale": scale,
                        "zscale": scale,
                        "rotation": angle
                    }
                )
            except Exception:
                msp.add_circle(center=(x, y), radius=1, dxfattribs={"layer": layer_name})

            # teks dengan offset
            msp.add_text(
                obj.get("name", ""),
                dxfattribs={
                    "height": 5.0,
                    "layer": "FEATURE_LABEL",
                    "color": 1,
                    "insert": (x + 2, y + 2),
                    "rotation": angle
                }
            )

    doc.saveas(output_path)
    return output_path

# ----------------------------
# Streamlit UI
# ----------------------------
def run_kmz_to_dwg():
    st.title("🏗️ KMZ → AUTOCAD (Rotate FDT/FAT/POLE by cable)")
    uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
    uploaded_template = st.file_uploader("📀 Upload Template DXF (optional)", type=["dxf"])

    if uploaded_kmz:
        extract_dir = "temp_kmz"; os.makedirs(extract_dir, exist_ok=True)
        with open("temp_upload.kmz","wb") as f: f.write(uploaded_kmz.read())
        kml_path = extract_kmz("temp_upload.kmz", extract_dir)
        items = parse_kml(kml_path)
        classified = classify_items(items)
        template_path = "template_ref.dxf"
        if uploaded_template:
            with open(template_path,"wb") as f: f.write(uploaded_template.read())
        else:
            template_path = template_path if os.path.exists(template_path) else ""
        output_dxf = "converted_output.dxf"
        try:
            result = build_dxf_with_rotation(classified, template_path, output_dxf)
            if result and os.path.exists(result):
                st.success("✅ DXF berhasil dibuat.")
                with open(result,"rb") as f:
                    st.download_button("⬇️ Download DXF", f, file_name=os.path.basename(result))
            else:
                st.error("❌ Gagal membuat DXF.")
        except Exception as e:
            st.error(f"❌ Error: {e}")

if __name__=="__main__":
    run_kmz_to_dwg()
