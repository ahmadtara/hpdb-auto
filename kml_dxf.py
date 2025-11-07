import os
import zipfile
import requests
import geopandas as gpd
import streamlit as st
import osmnx as ox
import json
from shapely.geometry import shape
from shapely.ops import unary_union
from fastkml import kml

HERE_API_KEY = "jGCMpa59MeURAH39Vzk94kutVqC3vl714_ZvcHodX14"

# ---------------------- KLASIFIKASI JALAN ----------------------
def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']:
        return 'HIGHWAYS'
    elif hwy in ['secondary', 'tertiary']:
        return 'MAJOR_ROADS'
    elif hwy in ['residential', 'unclassified', 'service']:
        return 'MINOR_ROADS'
    elif hwy in ['footway', 'path', 'cycleway']:
        return 'PATHS'
    return 'OTHER'

# ---------------------- EKSTRAK POLIGON ----------------------
def extract_polygon_from_kml_or_kmz(path):
    if path.endswith(".kmz"):
        with zipfile.ZipFile(path, "r") as z:
            for name in z.namelist():
                if name.endswith(".kml"):
                    z.extract(name, "/tmp")
                    path = os.path.join("/tmp", name)
                    break
    gdf = gpd.read_file(path)
    polygons = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polygons.empty:
        raise Exception("❌ Tidak ada Polygon di KML/KMZ.")
    return unary_union(polygons.geometry), polygons.crs

# ---------------------- AMBIL DATA JALAN DARI OSM ----------------------
def get_osm_roads(polygon):
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(polygon, tags=tags)
    except Exception:
        return gpd.GeoDataFrame()
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    roads = roads.explode(index_parts=False)
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    roads = roads.clip(polygon)
    return roads.reset_index(drop=True)

# ---------------------- OPSI FALLBACK: HERE API ----------------------
def get_here_roads(polygon):
    minx, miny, maxx, maxy = polygon.bounds
    url = (
        f"https://vector.hereapi.com/v2/vectortiles/base/mc"
        f"?apikey={HERE_API_KEY}&bbox={miny},{minx},{maxy},{maxx}&layers=roads"
    )
    resp = requests.get(url)
    if resp.status_code != 200:
        raise Exception(f"HERE API error: {resp.text}")
    
    data = resp.json()
    features = []
    for f in data.get("features", []):
        geom = shape(f["geometry"])
        if geom.intersects(polygon):
            features.append({
                "geometry": geom.intersection(polygon),
                "properties": f.get("properties", {})
            })
    if not features:
        return gpd.GeoDataFrame()
    return gpd.GeoDataFrame(features, crs="EPSG:4326")

# ---------------------- EKSPOR KE KML ----------------------
def export_to_kml(gdf, polygon, output_path):
    k = kml.KML()
    doc = kml.Document(ns=None, name="Roadmap", description="Jalan hasil konversi")
    k.append(doc)

    # Tambahkan batas poligon
    boundary = kml.Placemark(ns=None, name="Boundary")
    boundary.geometry = polygon
    doc.append(boundary)

    # Tambahkan setiap jalan ke dalam KML
    for _, row in gdf.iterrows():
        geom = row.geometry
        if geom.is_empty or not geom.is_valid:
            continue
        hwy = str(row.get("highway", ""))
        layer = classify_layer(hwy)
        pm = kml.Placemark(ns=None, name=layer)
        pm.geometry = geom
        doc.append(pm)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(k.to_string(prettyprint=True))

# ---------------------- PROSES UTAMA ----------------------
def process_kml_to_kml(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, _ = extract_polygon_from_kml_or_kmz(kml_path)
    roads = get_osm_roads(polygon)
    if roads.empty:
        roads = get_here_roads(polygon)
    if roads.empty:
        raise Exception("❌ Tidak ada jalan ditemukan (OSM & HERE kosong).")

    kml_output = os.path.join(output_dir, "roadmap.kml")
    export_to_kml(roads, polygon, kml_output)
    return kml_output, True

# ---------------------- STREAMLIT APP ----------------------
def run_kml_dxf():
    st.title("🌍 KML/KMZ → KML Road Converter (OSM + HERE Fallback)")
    kml_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())
                output_dir = "/tmp/output"
                kml_path, ok = process_kml_to_kml(temp_input, output_dir)
                if ok:
                    st.success("✅ Berhasil diekspor ke KML!")
                    with open(kml_path, "rb") as f:
                        st.download_button("⬇️ Download KML", data=f, file_name="roadmap.kml")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

