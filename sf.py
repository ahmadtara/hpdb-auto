import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
import math

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

# hanya folder yang dipakai
target_folders = {
    'FDT', 'NEW POLE 7-4', 'NEW POLE 9-4', 'EXISTING POLE EMR 7-4',
    'EXISTING POLE EMR 9-4', 'CABLE',
    'JOINT CLOSURE', 'SLACK HANGER', 'JALAN'
}

# mapping folder ➝ layer CAD
layer_mapping = {
    "CABLE": "FO 36 CORE",
    "JALAN": "JALAN",
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
    classified = {
        "FDT": [], "NEW_POLE_74": [], "NEW_POLE_94": [],
        "EXISTING_POLE": [], "CABLE": [], "CLOSURE": [],
        "COIL": [], "JALAN": []
    }
    for it in items:
        folder = it['folder']
        if folder == "FDT":
            classified["FDT"].append(it)
        elif folder == "NEW POLE 7-4":
            classified["NEW_POLE_74"].append(it)
        elif folder == "NEW POLE 9-4":
            classified["NEW_POLE_94"].append(it)
        elif "EXISTING POLE EMR" in folder:
            classified["EXISTING_POLE"].append(it)
        elif folder == "CABLE":
            classified["CABLE"].append(it)
        elif folder == "JOINT CLOSURE":
            classified["CLOSURE"].append(it)
        elif folder == "SLACK HANGER":
            classified["COIL"].append(it)
        elif folder == "JALAN":
            classified["JALAN"].append(it)
    return classified

# cari sudut jalan terdekat untuk rotasi teks
def nearest_road_angle(x, y, roads):
    min_dist = float("inf")
    best_angle = 0
    for road in roads:
        path = road.get("xy_path", [])
        for i in range(len(path) - 1):
            x1, y1 = path[i]
            x2, y2 = path[i + 1]
            # jarak titik ke segmen garis
            px = x2 - x1
            py = y2 - y1
            norm = px*px + py*py
            if norm == 0:
                continue
            u = max(0, min(1, ((x - x1) * px + (y - y1) * py) / norm))
            dx = x1 + u * px - x
            dy = y1 + u * py - y
            dist = dx*dx + dy*dy
            if dist < min_dist:
                min_dist = dist
                best_angle = math.degrees(math.atan2(py, px))
    return best_angle

# geser posisi teks supaya gak nabrak
def offset_point(x, y, angle, distance=3):
    rad = math.radians(angle + 90)  # geser tegak lurus
    return (x + distance * math.cos(rad), y + distance * math.sin(rad))

def draw_to_template(classified, template_path):
    doc = ezdxf.readfile(template_path)
    msp = doc.modelspace()

    # kumpulkan koordinat
    all_xy = []
    for cat_items in classified.values():
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
    for cat_items in classified.values():
        for obj in cat_items:
            if obj['type'] == 'point':
                obj['xy'] = shifted_all[idx]
                idx += 1
            elif obj['type'] == 'path':
                obj['xy_path'] = shifted_all[idx: idx + len(obj['coords'])]
                idx += len(obj['coords'])

    roads = classified.get("JALAN", [])

    # mapping folder ke block/layer
    for layer_name, cat_items in classified.items():
        for obj in cat_items:
            target_layer = layer_mapping.get(layer_name, layer_name)

            if obj['type'] == 'path':
                msp.add_lwpolyline(obj['xy_path'], dxfattribs={"layer": target_layer})
                continue

            x, y = obj['xy']
            block_name = None
            scale_x = scale_y = scale_z = 1.0

            if layer_name == "FDT":
                block_name = "FDT"
                scale_x = scale_y = scale_z = 0.0025
            elif layer_name == "NEW_POLE_74":
                block_name = "A$C14dd5346"
            elif layer_name == "NEW_POLE_94":
                block_name = "np9"
            elif layer_name == "EXISTING_POLE":
                block_name = "A$Cdb6fd7d1" if obj['folder'] in [
                    "EXISTING POLE EMR 7-4", "EXISTING POLE EMR 7-3"
                ] else "A$C14dd5346"
            elif layer_name == "CLOSURE":
                block_name = "CLOSURE"
                scale_x = scale_y = scale_z = 0.0030
            elif layer_name == "COIL":
                block_name = "COIL"
                scale_x = scale_y = scale_z = 0.0030

            inserted_block = False
            if block_name:
                try:
                    msp.add_blockref(
                        name=block_name,
                        insert=(x, y),
                        dxfattribs={
                            "layer": target_layer,
                            "xscale": scale_x,
                            "yscale": scale_y,
                            "zscale": scale_z,
                        }
                    )
                    inserted_block = True
                except Exception as e:
                    print(f"Gagal insert block {block_name}: {e}")

                if not inserted_block:
                    msp.add_circle(center=(x, y), radius=2, dxfattribs={"layer": target_layer})

                # hitung rotasi teks dari jalan terdekat
                angle = nearest_road_angle(x, y, roads)
            
                # teks diposisikan sama persis dengan block
                tx, ty = x, y
            
                # tambahkan teks label untuk semua object
                if obj["name"]:
                    msp.add_text(obj["name"], dxfattribs={
                        "height": 5.0,
                        "layer": target_layer,
                        "insert": (tx, ty),
                        "rotation": angle
                    })

    return doc


def run_sf():
    st.title("🏗️ KMZ ➝ DWG (Subset Folder)")
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
