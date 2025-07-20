import streamlit as st
import pandas as pd
from datetime import datetime
import os
from user_agents import parse as parse_ua

# --- Konfigurasi ---
st.set_page_config(page_title="MyRepublic HPDB", page_icon="🛜", layout="centered")

LOGO_URL = "https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png/960px-MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png"
LOG_FILE = "login_logs.csv"

# --- Data Login ---
USER_CREDENTIALS = {
    "snd": "snd0220",
    "obi": "obi",
    "tara": "123",
    "admin": "admin123"
}

# --- Fungsi Simpan Log ---
def log_login(username, user_agent):
    device_type = detect_device(user_agent)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row = pd.DataFrame([{
        "username": username,
        "timestamp": now,
        "device": device_type
    }])

    if os.path.exists(LOG_FILE):
        df = pd.read_csv(LOG_FILE)
        df = pd.concat([df, new_row], ignore_index=True)
    else:
        df = new_row

    df.to_csv(LOG_FILE, index=False)

# --- Fungsi Deteksi Device ---
def detect_device(user_agent_string):
    ua = parse_ua(user_agent_string)
    if ua.is_mobile:
        return "Mobile"
    elif ua.is_tablet:
        return "Tablet"
    elif ua.is_pc:
        return "Desktop"
    else:
        return "Unknown"

# --- Halaman Login ---
def login_page():
    st.image(LOGO_URL, use_column_width=True)
    st.markdown("### 🔐 Login to MyRepublic Auto HPDB")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username in USER_CREDENTIALS and password == USER_CREDENTIALS[username]:
            st.success(f"Login berhasil! 🎉 {'Selamat datang, admin' if username == 'admin' else ''}")
            st.session_state.logged_in = True
            st.session_state.username = username

            user_agent = st.request_headers.get("User-Agent", "unknown")
            log_login(username, user_agent)

            st.rerun()
        else:
            st.error("Username atau password salah!")

# --- Halaman Admin ---
def admin_dashboard():
    st.title("📊 Admin Dashboard - Login Tracker")

    if os.path.exists(LOG_FILE):
        df = pd.read_csv(LOG_FILE)
        st.dataframe(df)

        st.markdown("### 📈 Ringkasan")
        st.write(df.groupby(["username", "device"]).size().unstack(fill_value=0))
    else:
        st.info("Belum ada data login yang tercatat.")

# --- Halaman Utama Setelah Login ---
def main_page():
    username = st.session_state.username
    if username == "admin":
        admin_dashboard()
    else:
        st.title("📍 HPDB Converter")
        st.info(f"Hai {username}, fitur utama belum ditambahkan di sini.")
        st.write("🔧 Anda dapat menambahkan fitur KMZ ➜ HPDB di sini.")

# --- Routing ---
if "logged_in" not in st.session_state or not st.session_state.logged_in:
    login_page()
else:
    main_page()
