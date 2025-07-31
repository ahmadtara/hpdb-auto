# Combined Streamlit App for KMZ to HPDB and KML to DXF Processing

import os
import zipfile
import geopandas as gpd
import pandas as pd
import xml.etree.ElementTree as ET
import requests
import threading
from shapely.geometry import LineString, MultiLineString
from shapely.ops import unary_union, linemerge, snap, polygonize
from io import BytesIO
from fastkml import kml
import ezdxf
import streamlit as st
import osmnx as ox

# Constants
TELEGRAM_TOKEN = "7885701086:AAEgXt9fN7qufBbsf0NGBDvhtj3IqzohvKw"
TELEGRAM_CHAT_ID = "6122753506"
BOT_API_URL = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"
HERE_API_KEY = "iWCrFicKYt9_AOCtg76h76MlqZkVTn94eHbBl_cE8m0"
TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10

# User login
valid_users = {"snd": "snd0220", "obi": "obi", "rizky": "123"}
blocked_users = set()

# Telegram Functions
def send_telegram(msg):
    try:
        requests.post(f"{BOT_API_URL}/sendMessage", data={"chat_id": TELEGRAM_CHAT_ID, "text": msg})
    except:
        pass

def monitor_telegram():
    offset = None
    while True:
        try:
            resp = requests.get(f"{BOT_API_URL}/getUpdates", params={"timeout": 10, "offset": offset})
            data = resp.json()
            for update in data.get("result", []):
                offset = update["update_id"] + 1
                msg = update.get("message", {}).get("text", "")
                if msg.startswith("/add "):
                    _, uname, pw = msg.strip().split(maxsplit=2)
                    valid_users[uname] = pw
                    send_telegram(f"Akun '{uname}' berhasil ditambahkan.")
                elif msg.startswith("/block "):
                    uname = msg.strip().split()[1]
                    blocked_users.add(uname)
                    send_telegram(f"Akun '{uname}' berhasil diblokir.")
        except:
            continue

threading.Thread(target=monitor_telegram, daemon=True).start()

# UI Pages
def login_page():
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png/960px-MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png", width=300)
    st.markdown("## 🔐 Login to MyRepublic Auto HPDB Auto-Pilot⚡By.A.Tara-P.")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u in blocked_users:
            st.error("⛔ Akun ini telah diblokir.")
        elif u in valid_users and p == valid_users[u]:
            st.session_state["logged_in"] = True
            st.session_state["user"] = u
            send_telegram(f"✅ Login berhasil: {u}")
            st.rerun()
        else:
            st.error("❌ Username atau Password salah!")

def kmz_hpdb_page():
    # Implementation same as previous "main_page" function for HPDB, not shown here for brevity
    st.markdown("### 🚧 Halaman Auto HPDB masih dalam proses loading...")

# Helper Functions for KML to DXF

def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']: return 'HIGHWAYS', 10
    if hwy in ['secondary', 'tertiary']: return 'MAJOR_ROADS', 10
    if hwy in ['residential', 'unclassified', 'service']: return 'MINOR_ROADS', 10
    if hwy in ['footway', 'path', 'cycleway']: return 'PATHS', 10
    return 'OTHER', DEFAULT_WIDTH

def extract_polygon_from_kml(path):
    gdf = gpd.read_file(path)
    polys = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polys.empty: raise Exception("No Polygon found in KML")
    return unary_union(polys.geometry), polys.crs

def get_osm_roads(poly):
    tags = {"highway": True}
    roads = ox.features_from_polygon(poly, tags=tags)
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])].explode(index_parts=False)
    roads = roads[roads.geometry.notnull() & ~roads.geometry.is_empty].clip(poly)
    roads["geometry"] = roads.geometry.apply(lambda g: snap(g, g, 0.0001))
    return roads.reset_index(drop=True)

def strip_z(g):
    if g.geom_type == "LineString" and g.has_z: return LineString([(x, y) for x, y, *_ in g.coords])
    if g.geom_type == "MultiLineString": return MultiLineString([LineString([(x, y) for x, y, *_ in l.coords]) if l.has_z else l for l in g.geoms])
    return g

def export_to_dxf(gdf, dxf_path, polygon=None, polygon_crs=None):
    doc = ezdxf.new()
    msp = doc.modelspace()
    buffers = []
    for _, r in gdf.iterrows():
        geom, hwy = strip_z(r.geometry), str(r.get("highway", ""))
        layer, width = classify_layer(hwy)
        if geom.is_empty or not geom.is_valid: continue
        line = linemerge(geom)
        if isinstance(line, (LineString, MultiLineString)):
            buf = line.buffer(width/2, resolution=8, join_style=2)
            buffers.append(buf)
    if not buffers: raise Exception("❌ Tidak ada garis valid untuk diekspor.")
    merged = unary_union(buffers)
    outlines = list(polygonize(merged.boundary))
    for outline in outlines:
        coords = [(x, y) for x, y in outline.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})
    if polygon is not None and polygon_crs is not None:
        p = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs(TARGET_EPSG).iloc[0]
        for part in (p.geoms if p.geom_type == 'MultiPolygon' else [p]):
            coords = [(x, y) for x, y in part.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})
    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)

def kml_to_dxf_page():
    st.title("🌍 KML → DXF Road Converter")
    kml_file = st.file_uploader("Upload file .KML", type=["kml"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_path = f"/tmp/{kml_file.name}"
                with open(temp_path, "wb") as f: f.write(kml_file.read())
                polygon, crs = extract_polygon_from_kml(temp_path)
                roads = get_osm_roads(polygon)
                if roads.empty: raise Exception("Tidak ada jalan ditemukan.")
                roads_utm = roads.to_crs(TARGET_EPSG)
                geojson_path = "/tmp/roadmap_osm.geojson"
                dxf_path = "/tmp/roadmap_osm.dxf"
                roads_utm.to_file(geojson_path, driver="GeoJSON")
                export_to_dxf(roads_utm, dxf_path, polygon, crs)
                st.success("✅ Berhasil diekspor ke DXF!")
                with open(dxf_path, "rb") as f:
                    st.download_button("⬇️ Download DXF", f, file_name="roadmap_osm.dxf")
                with open(geojson_path, "rb") as f:
                    st.download_button("⬇️ Download GeoJSON", f, file_name="roadmap_osm.geojson")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

# Router
st.set_page_config(page_title="MyRepublic GeoTools", layout="wide")
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["user"] = None

if not st.session_state["logged_in"]:
    login_page()
else:
    tab = st.sidebar.radio("Pilih Halaman", ["Auto HPDB", "KML to DXF"])
    if tab == "Auto HPDB":
        kmz_hpdb_page()
    elif tab == "KML to DXF":
        kml_to_dxf_page()
