import os
import geopandas as gpd
from shapely.geometry import shape, Polygon, MultiPolygon
from fastkml import kml
import ezdxf
import pandas as pd
import streamlit as st
import requests
import numpy as np
from ultralytics import YOLO
from PIL import Image
import tempfile
import matplotlib.pyplot as plt
from shapely.ops import unary_union
import cv2

TARGET_EPSG = "EPSG:32760"  # UTM Zone 60S
MODEL_PATH = "yolov8-building.pt"  # Path ke model segmentasi bangunan YOLOv8 (custom)
GOOGLE_MAPS_API_KEY = "AIzaSyAOVYRIgupAurZup5y1PRh8Ismb1A3lLao"

# --- Ekstrak Polygon Area dari KML ---
def extract_polygon_from_kml(kml_path):
    gdf = gpd.read_file(kml_path)
    polygons = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polygons.empty:
        raise Exception("No Polygon found in KML")
    return unary_union(polygons.geometry), polygons.crs

# --- Ambil Citra Google Maps dari Polygon ---
def download_static_map(polygon):
    from shapely.geometry import box

    bounds = polygon.bounds
    west, south, east, north = bounds
    center_lat = (south + north) / 2
    center_lon = (west + east) / 2

    url = (
        f"https://maps.googleapis.com/maps/api/staticmap"
        f"?center={center_lat},{center_lon}"
        f"&zoom=18&size=640x640&maptype=satellite"
        f"&key={GOOGLE_MAPS_API_KEY}"
    )

    response = requests.get(url)
    if not response.ok:
        raise Exception("Gagal mengunduh citra dari Google Maps")

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    temp_file.write(response.content)
    temp_file.close()
    return Image.open(temp_file.name).convert("RGB")

# --- Deteksi Bangunan dari Citra ---
def detect_buildings_from_image(image_path):
    model = YOLO(MODEL_PATH)
    results = model(image_path)
    masks = results[0].masks

    if masks is None:
        return []

    polygons = []
    for mask in masks.data:
        mask_np = mask.cpu().numpy().astype(np.uint8) * 255
        contours, _ = cv2.findContours(mask_np, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for cnt in contours:
            if len(cnt) >= 3:
                coords = [(int(p[0][0]), int(p[0][1])) for p in cnt]
                poly = Polygon(coords)
                if poly.is_valid:
                    polygons.append(poly)

    return polygons

# --- Ekspor ke DXF ---
def export_to_dxf_buildings(gdf, dxf_path, polygon=None, polygon_crs=None):
    doc = ezdxf.new()
    msp = doc.modelspace()

    bounds = [(pt[0], pt[1]) for geom in gdf.geometry for pt in geom.exterior.coords]
    min_x = min(x for x, y in bounds)
    min_y = min(y for x, y in bounds)

    for geom in gdf.geometry:
        if geom.geom_type == "Polygon":
            coords = [(pt[0] - min_x, pt[1] - min_y) for pt in geom.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "BUILDINGS"})
        elif geom.geom_type == "MultiPolygon":
            for poly in geom.geoms:
                coords = [(pt[0] - min_x, pt[1] - min_y) for pt in poly.exterior.coords]
                msp.add_lwpolyline(coords, dxfattribs={"layer": "BUILDINGS"})

    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)

# --- Proses Utama ---
def process_kml_to_dxf(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml(kml_path)

    image = download_static_map(polygon)
    image_path = os.path.join(output_dir, "map.png")
    image.save(image_path)

    building_polygons = detect_buildings_from_image(image_path)
    if not building_polygons:
        raise Exception("Tidak ada bangunan terdeteksi dari citra.")

    gdf = gpd.GeoDataFrame(geometry=building_polygons, crs="EPSG:4326")
    gdf = gdf.clip(polygon)

    dxf_path = os.path.join(output_dir, "buildings_detected.dxf")
    geojson_path = os.path.join(output_dir, "buildings_detected.geojson")
    gdf.to_file(geojson_path, driver="GeoJSON")
    export_to_dxf_buildings(gdf.to_crs(TARGET_EPSG), dxf_path)
    return dxf_path, geojson_path, True

# --- Streamlit UI ---
st.set_page_config(page_title="KML → DXF Auto Building Detector", layout="wide")
st.title("🌍 KML → DXF Auto Building Detector")
st.caption("Upload file .KML (batas area perumahan)")

kml_file = st.file_uploader("Upload file .KML", type=["kml"])

if kml_file:
    with st.spinner("💫 Memproses file dan mendeteksi rumah..."):
        try:
            temp_input = f"/tmp/{kml_file.name}"
            with open(temp_input, "wb") as f:
                f.write(kml_file.read())

            output_dir = "/tmp/output"
            dxf_path, geojson_path, ok = process_kml_to_dxf(temp_input, output_dir)

            if ok:
                st.success("✅ Berhasil diekspor ke DXF!")
                with open(dxf_path, "rb") as f:
                    st.download_button("⬇️ Download DXF", data=f, file_name="buildings_detected.dxf")
                with open(geojson_path, "rb") as f:
                    st.download_button("⬇️ Download GeoJSON", data=f, file_name="buildings_detected.geojson")
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan: {e}")
