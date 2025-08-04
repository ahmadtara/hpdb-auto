import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
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

def get_nearest_angle(x, y, all_lines):
    min_dist = float('inf')
    best_angle = 0.0
    for path in all_lines:
        coords = path['xy_path']
        for i in range(len(coords) - 1):
            x1, y1 = coords[i]
            x2, y2 = coords[i + 1]
            px = x2 - x1
            py = y2 - y1
            norm = math.hypot(px, py)
            if norm == 0:
                continue
            t = max(0, min(1, ((x - x1) * px + (y - y1) * py) / (norm * norm)))
            proj_x = x1 + t * px
            proj_y = y1 + t * py
            dist = math.hypot(x - proj_x, y - proj_y)
            if dist < min_dist:
                min_dist = dist
                best_angle = math.degrees(math.atan2(py, px))
    return best_angle

def draw_to_template(doc, classified):
    msp = doc.modelspace()
    text_height = 2.5
    text_spacing = 4.5
    used_positions = set()

    # Konversi path koordinat menjadi xy_path terlebih dahulu
    for key in classified:
        for item in classified[key]:
            if item['type'] == 'path':
                item['xy_path'] = [latlon_to_xy(lat, lon) for lat, lon in item['coords']]

    for label in ["FDT", "FAT", "NEW_POLE", "EXISTING_POLE", "POLE"]:
        for item in classified[label]:
            x, y = latlon_to_xy(item['latitude'], item['longitude'])

            angle = get_nearest_angle(x, y, classified['DISTRIBUTION_CABLE'])

            # offset otomatis bila tabrakan
            offset_y = 0
            while (round(x, 1), round(y + offset_y, 1)) in used_positions:
                offset_y += text_spacing
            used_positions.add((round(x, 1), round(y + offset_y, 1)))

            msp.add_text(
                item['name'],
                dxfattribs={
                    'height': text_height,
                    'rotation': angle
                }
            ).set_pos((x, y + offset_y))

    return doc

def run_kmz_to_dwg(kmz_file):
    with open("temp.kmz", "wb") as f:
        f.write(kmz_file.read())
    kml_path = extract_kmz("temp.kmz", "temp")
    items = parse_kml(kml_path)
    classified = classify_items(items)
    doc = ezdxf.new()
    updated_doc = draw_to_template(doc, classified)
    output_path = "output.dxf"
    updated_doc.saveas(output_path)
    return output_path

st.title("KMZ to DXF Converter with Auto Text Rotation")
uploaded = st.file_uploader("Upload KMZ file", type="kmz")
if uploaded:
    output = run_kmz_to_dwg(uploaded)
    with open(output, "rb") as f:
        st.download_button("Download DXF", f, file_name="converted.dxf")
