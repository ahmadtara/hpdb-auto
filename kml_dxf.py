import streamlit as st
import geopandas as gpd
import zipfile
import os
from io import BytesIO
from shapely.geometry import LineString, MultiLineString, Polygon
from shapely.ops import unary_union
from shapely.wkt import loads as load_wkt
import pandas as pd
from fastkml import kml
import tempfile
from pyproj import CRS

# ---------------------------
# Utility Functions
# ---------------------------

def strip_z(geom):
    """Hapus koordinat Z (tinggi) dari geometri shapely."""
    if geom is None or geom.is_empty:
        return geom
    try:
        if geom.has_z:
            coords = [(x, y) for x, y, *_ in geom.coords]
            return type(geom)(coords)
        elif geom.geom_type == "MultiLineString":
            return MultiLineString([strip_z(g) for g in geom.geoms])
        elif geom.geom_type == "Polygon":
            exterior = [(x, y) for x, y, *_ in geom.exterior.coords]
            interiors = [[(x, y) for x, y, *_ in ring.coords] for ring in geom.interiors]
            return Polygon(exterior, interiors)
        else:
            return geom
    except Exception:
        return geom

def classify_layer(hwy):
    """Klasifikasi layer berdasarkan tipe jalan"""
    if "primary" in hwy or "arterial" in hwy:
        return "Arterial", "#FF0000"
    elif "secondary" in hwy or "collector" in hwy:
        return "Collector", "#FFA500"
    elif "residential" in hwy or "local" in hwy:
        return "Local", "#00FF00"
    else:
        return "Other", "#999999"

# ---------------------------
# Ekspor ke KML
# ---------------------------

def export_to_kml(gdf, kml_path, polygon=None, polygon_crs=None):
    ns = '{http://www.opengis.net/kml/2.2}'
    k = kml.KML()
    doc = kml.Document(ns, 'docid', 'RoadMap', 'Generated road map')
    k.append(doc)

    # Pastikan CRS ke EPSG:4326 (lat/lon)
    if gdf.crs is not None and gdf.crs.to_string() != "EPSG:4326":
        gdf = gdf.to_crs("EPSG:4326")

    # Tambahkan setiap feature
    for _, row in gdf.iterrows():
        geom = strip_z(row.geometry)
        if geom.is_empty or not geom.is_valid:
            continue

        hwy = str(row.get("highway", row.get("type", "")))
        layer, color = classify_layer(hwy)

        # Handle MultiLineString
        if isinstance(geom, MultiLineString):
            for part in geom.geoms:
                placemark = kml.Placemark(ns, None, f"{layer}", f"Type: {hwy}", geometry=part)
                doc.append(placemark)
        else:
            placemark = kml.Placemark(ns, None, f"{layer}", f"Type: {hwy}", geometry=geom)
            doc.append(placemark)

    # Tambah boundary polygon jika ada
    if polygon is not None and polygon_crs is not None:
        try:
            poly = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs("EPSG:4326").iloc[0]
            if poly.geom_type == 'Polygon':
                placemark = kml.Placemark(ns, None, "BOUNDARY", "Boundary polygon", geometry=poly)
                doc.append(placemark)
            elif poly.geom_type == 'MultiPolygon':
                for idx, p in enumerate(poly.geoms):
                    placemark = kml.Placemark(ns, None, f"BOUNDARY_{idx+1}", "Boundary part", geometry=p)
                    doc.append(placemark)
        except Exception as e:
            print(f"⚠️ Boundary polygon skip: {e}")

    with open(kml_path, "w", encoding="utf-8") as f:
        f.write(k.to_string(prettyprint=True))

# ---------------------------
# Fungsi utama Streamlit
# ---------------------------

def run_kmz_to_kml():
    st.title("📍 KMZ ➜ KML Converter")

    kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
    if not kmz_file:
        st.stop()

    # Extract KMZ
    with zipfile.ZipFile(kmz_file, 'r') as z:
        kml_files = [f for f in z.namelist() if f.endswith('.kml')]
        if not kml_files:
            st.error("❌ File KMZ tidak berisi file .kml!")
            st.stop()

        kml_data = z.read(kml_files[0])

    with tempfile.NamedTemporaryFile(delete=False, suffix=".kml") as temp_kml:
        temp_kml.write(kml_data)
        temp_kml_path = temp_kml.name

    # Baca KML jadi GeoDataFrame
    try:
        gdf = gpd.read_file(temp_kml_path)
    except Exception as e:
        st.error(f"❌ Gagal membaca KML: {e}")
        st.stop()

    st.success(f"✅ Berhasil memuat {len(gdf)} fitur dari KMZ")

    # Tampilkan preview
    st.dataframe(gdf.head())

    # Simpan ke KML baru
    output_path = os.path.join(tempfile.gettempdir(), "converted_output.kml")
    export_to_kml(gdf, output_path)

    with open(output_path, "rb") as f:
        st.download_button(
            label="⬇️ Download Hasil KML",
            data=f,
            file_name="converted_output.kml",
            mime="application/vnd.google-earth.kml+xml"
        )

    st.success("🎉 Konversi KMZ ➜ KML berhasil!")

# ---------------------------
# Jalankan Aplikasi
# ---------------------------

if __name__ == "__main__":
    run_kmz_to_kml()
