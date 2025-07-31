import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import requests
import threading
import os
import zipfile
import geopandas as gpd
from shapely.geometry import Polygon, MultiPolygon, GeometryCollection, LineString, MultiLineString
from fastkml import kml
import osmnx as ox
import ezdxf
from shapely.ops import unary_union, linemerge, snap, split, polygonize


# ------------------ LOGIN ------------------ #
USERS = {"admin": "admin123", "tara": "123", "rizky": "123"}

def login():
    st.title("🔐 Login Aplikasi")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username in USERS and USERS[username] == password:
            st.session_state["logged_in"] = True
            st.session_state["user"] = username
            st.success("Login berhasil")
            st.experimental_rerun()
        else:
            st.error("Username atau password salah")

# ------------------ KML ➜ DXF ------------------ #
def process_kml_to_dxf(kml_path, output_dir):
    tree = ET.parse(kml_path)
    root = tree.getroot()

    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    placemarks = root.findall('.//kml:Placemark', ns)

    items = []

    for placemark in placemarks:
        name = placemark.find('kml:name', ns)
        coords = placemark.find('.//kml:coordinates', ns)
        if coords is not None:
            points = []
            for coord in coords.text.strip().split():
                lon, lat, *_ = map(float, coord.split(','))
                points.append((lon, lat))
            items.append({
                "name": name.text if name is not None else "",
                "coordinates": points
            })

    if not items:
        return None, None, False

    geojson = {
        "type": "FeatureCollection",
        "features": []
    }

    dxf_path = os.path.join(output_dir, "roadmap_osm.dxf")
    geojson_path = os.path.join(output_dir, "roadmap_osm.geojson")

    doc = ezdxf.new()
    msp = doc.modelspace()

    for item in items:
        line = LineString(item["coordinates"])
        geojson["features"].append({
            "type": "Feature",
            "geometry": mapping(line),
            "properties": {"name": item["name"]}
        })
        msp.add_lwpolyline(item["coordinates"])

    doc.saveas(dxf_path)

    with open(geojson_path, "w") as f:
        json.dump(geojson, f)

    return dxf_path, geojson_path, True

def kml_to_dxf_page():
    st.title("🌍 KML ➜ DXF Road Converter")
    st.caption("Upload file .KML (area batas cluster)")

    kml_file = st.file_uploader("Upload file .KML", type=["kml"])

    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())

                output_dir = "/tmp/output"
                os.makedirs(output_dir, exist_ok=True)
                dxf_path, geojson_path, ok = process_kml_to_dxf(temp_input, output_dir)

                if ok:
                    st.success("✅ Berhasil diekspor ke DXF!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF", data=f, file_name="roadmap_osm.dxf")
                    with open(geojson_path, "rb") as f:
                        st.download_button("⬇️ Download GeoJSON", data=f, file_name="roadmap_osm.geojson")
                else:
                    st.warning("🚫 Tidak ada data garis yang ditemukan dalam file KML.")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

# ------------------ KMZ ➜ HPDB ------------------ #
def hpdb_page():
    st.title("📍 KMZ ➜ HPDB (Auto-Pilot ⚡By.A.Tara-P.)")
    st.write(f"Hai, **{st.session_state['user']}** 👋")
    st.markdown("⚙️ Fitur ini masih dalam tahap pengembangan penuh.")

    uploaded_file = st.file_uploader("📦 Upload file .kmz", type=["kmz"])
    if uploaded_file is not None:
        st.info("📄 KMZ berhasil diupload. (Parsing logic belum dimasukkan)")

# ------------------ MENU UTAMA ------------------ #
def main_page():
    st.sidebar.title("📂 Menu Utama")
    menu = st.sidebar.radio("Pilih halaman", ["KMZ ➜ HPDB", "KML ➜ DXF Road Converter"])

    if st.sidebar.button("🔒 Logout"):
        st.session_state["logged_in"] = False
        st.session_state["user"] = None
        st.rerun()

    if menu == "KMZ ➜ HPDB":
        hpdb_page()
    elif menu == "KML ➜ DXF Road Converter":
        kml_to_dxf_page()

# ------------------ MAIN ------------------ #
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["user"] = None

if st.session_state["logged_in"]:
    main_page()
else:
    login()
