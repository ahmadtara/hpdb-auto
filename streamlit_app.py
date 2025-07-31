import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import requests
import threading
import os
import geopandas as gpd
from shapely.geometry import Polygon, MultiPolygon, GeometryCollection, LineString, MultiLineString
from fastkml import kml
import osmnx as ox
import ezdxf
from shapely.ops import unary_union, linemerge, snap, split, polygonize

# -------------------------------- CONFIG -------------------------------- #
TELEGRAM_TOKEN = "7885701086:AAEgXt9fN7qufBbsf0NGBDvhtj3IqzohvKw"
TELEGRAM_CHAT_ID = "6122753506"
HERE_API_KEY = "iWCrFicKYt9_AOCtg76h76MlqZkVTn94eHbBl_cE8m0"
BOT_API_URL = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"
TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10

valid_users = {"snd": "snd0220", "obi": "obi", "rizky": "123"}
blocked_users = set()

# ------------------------- TELEGRAM THREAD ------------------------ #
def send_telegram(message):
    try:
        requests.post(f"{BOT_API_URL}/sendMessage", data={"chat_id": TELEGRAM_CHAT_ID, "text": message})
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

# ------------------------- LOGIN PAGE ------------------------ #
def login_page():
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png/960px-MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png", width=300)
    st.markdown("## 🔐 Login to MyRepublic Auto HPDB Auto-Pilot⚡By.A.Tara-P.")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username in blocked_users:
            st.error("⛔ Akun ini telah diblokir.")
        elif username in valid_users and password == valid_users[username]:
            st.session_state["logged_in"] = True
            st.session_state["user"] = username
            st.success(f"Login berhasil! 🎉 Selamat datang, {username}.")
            send_telegram(f"✅ Login berhasil: {username}")
            st.rerun()
        else:
            st.error("❌ Username atau Password salah!")

# ------------------------- MENU ROUTER ------------------------ #
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

# ------------------------- KMZ TO HPDB ------------------------ #
def hpdb_page():
    st.title("📍 KMZ ➜ HPDB (Auto-Pilot ⚡By.A.Tara-P.)")
    st.write(f"Hai, **{st.session_state['user']}** 👋")

    kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
    template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

    def extract_placemarks(kmz_bytes):
        def recurse_folder(folder, ns, path=""):
            items = []
            name_el = folder.find("kml:name", ns)
            folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
            new_path = f"{path}/{folder_name}" if path else folder_name
            for sub in folder.findall("kml:Folder", ns):
                items += recurse_folder(sub, ns, new_path)
            for pm in folder.findall("kml:Placemark", ns):
                nm = pm.find("kml:name", ns)
                coord = pm.find(".//kml:coordinates", ns)
                if nm is not None and coord is not None:
                    lon, lat = coord.text.strip().split(",")[:2]
                    items.append({
                        "name": nm.text.strip(),
                        "lat": float(lat),
                        "lon": float(lon),
                        "path": new_path
