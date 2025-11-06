import os
import zipfile
import requests
from fastkml import kml
import geopandas as gpd
import streamlit as st
from shapely.geometry import shape, Polygon, MultiPolygon, LineString, MultiLineString
from shapely.ops import unary_union
import osmnx as ox
import json

TARGET_EPSG = "EPSG:32760"
HERE_API_KEY = "k1mDEfR1Q3A_MtLqkxLrhbDcS-1oC4r7WzlgPrcv4Rk"

def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']:
        return 'HIGHWAYS', 14
    elif hwy in ['secondary', 'tertiary']:
        return 'MAJOR_ROADS', 10
    elif hwy in ['residential', 'unclassified', 'service']:
        return 'MINOR_ROADS', 8
    elif hwy in ['footway', 'path', 'cycleway']:
        return 'PATHS', 4
    return 'OTHER', 8

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
        raise Exception("❌ No Polygon found in KML/KMZ")
    return unary_union(polygons.geometry), polygons.crs

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
            features.append({"geometry": geom.intersection(polygon), "properties": f["properties"]})
    if not features:
        return gpd.GeoDataFrame()
    return gpd.GeoDataFrame(features, crs="EPSG:4326")

def export_to_kml(gdf, kml_path):
    """Export GeoDataFrame garis ke file KML"""
    k = kml.KML()
    doc = kml.Document(ns="", name="Road Network", description="Extracted road lines")
    k.append(doc)

    for _, row in gdf.iterrows():
        geom = row.geometry
        if geom.is_empty or not geom.is_valid:
            continue
        hwy = str(row.get("highway", ""))
        layer, width = classify_layer(hwy)
        placemark = kml.Placemark(ns="", name=layer, description=hwy)
        placemark.geometry = geom
        doc.append(placemark)

    with open(kml_path, "w", encoding="utf-8") as f:
        f.write(k.to_string(prettyprint=True))

def process_kml_to_kml(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml_or_kmz(kml_path)

    roads = get_osm_roads(polygon)
    if roads.empty:
        roads = get_here_roads(polygon)
    if roads.empty:
        raise Exception("❌ Tidak ada jalan ditemukan (OSM & HERE kosong).")

    # Pastikan CRS ke WGS84 agar kompatibel dengan KML
    roads = roads.to_crs("EPSG:4326")
    kml_path_out = os.path.join(output_dir, "roadmap.kml")
    export_to_kml(roads, kml_path_out)
    return kml_path_out, True

def run_kml_to_kml():
    st.title("🌍 KML/KMZ ➜ Road KML Converter (OSM + HERE fallback)")
    kml_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())
                output_dir = "/tmp/output"
                kml_out, ok = process_kml_to_kml(temp_input, output_dir)
                if ok:
                    st.success("✅ Berhasil diekspor ke KML!")
                    with open(kml_out, "rb") as f:
                        st.download_button("⬇️ Download Road KML", data=f, file_name="roadmap.kml")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

# Jalankan di Streamlit
# run_kml_to_kml()
