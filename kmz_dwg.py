import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString
from shapely.ops import unary_union, polygonize, linemerge, snap
import osmnx as ox

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)
TARGET_EPSG = "EPSG:32760"

target_folders = {
    'FDT', 'FAT', 'HP COVER', 'NEW POLE 7-3', 'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY', 'DISTRIBUTION CABLE', 'SLING WIRE', 'KOTAK', 'BOUNDARY CLUSTER'
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

            poly_coord = pm.find('.//kml:Polygon//kml:coordinates', ns)
            if poly_coord is not None:
                coords = []
                for c in poly_coord.text.strip().split():
                    lon, lat, *_ = c.split(',')
                    coords.append((float(lon), float(lat)))
                items.append({
                    'type': 'polygon',
                    'name': name_text,
                    'coords': coords,
                    'folder': folder_name
                })
                continue

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
    return items

def extract_boundary_from_items(items):
    polys = []
    for obj in items:
        if obj['type'] == 'polygon' and obj['folder'] == 'BOUNDARY CLUSTER':
            poly = Polygon(obj['coords'])
            polys.append(poly)
    if not polys:
        raise Exception("❌ Tidak ada Polygon valid dari folder BOUNDARY CLUSTER.")
    return unary_union(polys)

def get_osm_roads(polygon):
    roads = ox.features_from_polygon(polygon, tags={"highway": True})
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    roads = roads.explode(index_parts=False)
    roads = roads.clip(polygon)
    roads["geometry"] = roads["geometry"].apply(lambda g: snap(g, g, tolerance=0.0001))
    return roads

def latlon_to_xy(lat, lon):
    return transformer.transform(lon, lat)

def apply_offset(points_xy):
    xs = [x for x, y in points_xy]
    ys = [y for x, y in points_xy]
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    return [(x - cx, y - cy) for x, y in points_xy], (cx, cy)

def run_kmz_to_dwg():
    st.title("📦 KMZ Cluster Tool (Gabung DWG + Roads)")
    st.markdown("""
    ### 🔧 Panduan:
    - Upload file **KMZ** berisi semua folder (FDT, FAT, HP COVER, BOUNDARY CLUSTER, dll).
    - Upload file **Template DXF** untuk memasukkan simbol FDT/FAT/pole.
    - Folder **BOUNDARY CLUSTER** akan dipakai sebagai batas polygon jalan.
    """)

    uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
    uploaded_template = st.file_uploader("📀 Upload Template DXF", type=["dxf"])

    if uploaded_kmz and uploaded_template:
        extract_dir = "temp_kmz"
        os.makedirs(extract_dir, exist_ok=True)

        try:
            with st.spinner("🔍 Memproses data..."):
                with open("template_ref.dxf", "wb") as f:
                    f.write(uploaded_template.read())

                kml_path = extract_kmz(uploaded_kmz, extract_dir)
                items = parse_kml(kml_path)
                boundary_polygon = extract_boundary_from_items(items)
                roads = get_osm_roads(boundary_polygon)

                doc = ezdxf.readfile("template_ref.dxf")
                msp = doc.modelspace()

                # Tambahkan jalan dari OSM ke DXF utama
                if not roads.empty:
                    roads_utm = roads.to_crs(TARGET_EPSG)
                    minx, miny, _, _ = roads_utm.total_bounds
                    for _, row in roads_utm.iterrows():
                        geom = row.geometry
                        geom = linemerge(geom) if isinstance(geom, MultiLineString) else geom
                        buffered = geom.buffer(5)
                        outlines = list(polygonize(buffered.boundary))
                        for outline in outlines:
                            coords = [(pt[0] - minx, pt[1] - miny) for pt in outline.exterior.coords]
                            msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})

                # Tambahkan semua koordinat point dan path dari KMZ seperti biasa (opsional)
                all_xy = []
                for obj in items:
                    if obj['type'] == 'point':
                        all_xy.append(latlon_to_xy(obj['latitude'], obj['longitude']))
                    elif obj['type'] == 'path':
                        all_xy.extend([latlon_to_xy(lat, lon) for lat, lon in obj['coords']])

                if all_xy:
                    shifted_all, (cx, cy) = apply_offset(all_xy)
                    idx = 0
                    for obj in items:
                        if obj['type'] == 'point':
                            obj['xy'] = shifted_all[idx]
                            idx += 1
                        elif obj['type'] == 'path':
                            obj['xy_path'] = shifted_all[idx: idx + len(obj['coords'])]
                            idx += len(obj['coords'])
                            msp.add_lwpolyline(obj['xy_path'], dxfattribs={"layer": obj['folder']})

                out_path = "combined_output.dxf"
                doc.saveas(out_path)

                st.success("✅ DXF gabungan berhasil dibuat!")
                with open(out_path, "rb") as f:
                    st.download_button("⬇️ Download DXF Gabungan", f, file_name="combined_output.dxf")

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan: {e}")

run_kmz_to_dwg()
