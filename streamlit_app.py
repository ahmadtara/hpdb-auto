import os
import zipfile
import geopandas as gpd
from shapely.geometry import Polygon, MultiPolygon, GeometryCollection, LineString, MultiLineString
from shapely.ops import unary_union, linemerge, snap, split, polygonize
import osmnx as ox
import ezdxf
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
import streamlit as st
import requests
import threading

# ---------------- CONFIG ---------------- #
st.set_page_config(page_title="MyRepublic Toolkit", layout="wide")

# Telegram
TELEGRAM_TOKEN = "TOKEN_KAMU"
TELEGRAM_CHAT_ID = "CHAT_ID_KAMU"
BOT_API_URL = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"

# HERE API
HERE_API_KEY = "KEY_HERE_KAMU"

# Login users
valid_users = {"snd": "snd0220", "obi": "obi", "rizky": "123"}
blocked_users = set()

# Fungsi Telegram
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

# ---------------- LOGIN PAGE ---------------- #
def login_page():
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png/960px-MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png", width=300)
    st.markdown("## 🔐 Login to MyRepublic Toolkit ⚡ By.A.Tara-P.")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username in blocked_users:
            st.error("⛔ Akun ini telah diblokir.")
        elif username in valid_users and password == valid_users[username]:
            st.session_state["logged_in"] = True
            st.session_state["user"] = username
            send_telegram(f"✅ Login berhasil: {username}")
            st.rerun()
        else:
            st.error("❌ Username atau Password salah!")

# ---------------- AUTO HPDB ---------------- #
def page_hpdb():
    from_hpdb import run_hpdb
    run_hpdb(HERE_API_KEY)

# ---------------- KML to DXF ---------------- #
def page_kml_dxf():
    from kml_dxf import run_kml_dxf
    run_kml_dxf()

# ---------------- MAIN ---------------- #
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["user"] = None

if not st.session_state["logged_in"]:
    login_page()
else:
    menu = st.sidebar.radio("📌 Menu", ["Auto HPDB", "KML → DXF Converter", "Logout"])
    st.sidebar.markdown(f"👤 Logged in as: **{st.session_state['user']}**")
    if menu == "Auto HPDB":
        page_hpdb()
    elif menu == "KML → DXF Converter":
        page_kml_dxf()
    elif menu == "Logout":
        st.session_state["logged_in"] = False
        st.session_state["user"] = None
        st.rerun()
