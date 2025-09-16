import streamlit as st
import threading
import requests

# -------------- ✅ KONFIGURASI ---------------- #
TELEGRAM_TOKEN = "7656007924:AAGi1it2M7jE0Sen28myiPhEmYPd1-jsI_Q"
TELEGRAM_CHAT_ID = "6122753506"
HERE_API_KEY = "iWCrFicKYt9_AOCtg76h76MlqZkVTn94eHbBl_cE8m0"

BOT_API_URL = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"

st.set_page_config(page_title="MyRepublic Toolkit", layout="wide")

# -------------- ✅ USER LOGIN ---------------- #
valid_users = {
    "zikni": "zikni",
    "vendor": "vendor",
    "toko": "toko"
}
blocked_users = set()

# -------------- ✅ TELEGRAM ---------------- #
def send_telegram(message):
    try:
        requests.post(f"{BOT_API_URL}/sendMessage", data={"chat_id": TELEGRAM_CHAT_ID, "text": message})
    except:
        pass

# -------------- ✅ PANTAU PESAN BOT ---------------- #
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
                    send_telegram(f"✅ Akun '{uname}' berhasil ditambahkan.")
                elif msg.startswith("/block "):
                    uname = msg.strip().split()[1]
                    blocked_users.add(uname)
                    send_telegram(f"⛔ Akun '{uname}' telah diblokir.")
        except:
            continue

# -------------- ✅ THREAD BACKGROUND ---------------- #
threading.Thread(target=monitor_telegram, daemon=True).start()

# -------------- ✅ LOGIN FORM ---------------- #
def login_page():
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png/960px-MyRepublic_NEW_LOGO_%28September_2023%29_Logo_MyRepublic_Horizontal_-_Black_%281%29.png", width=300)
    st.markdown("## 🔐 Login to Teknologia ⚡")

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

# -------------- ✅ PANGGIL MODUL FUNGSIONALITAS ---------------- #
from from_hpdb import run_hpdb
from kml_dxf import run_kml_dxf
from kmz_dwg import run_kmz_to_dwg  # ✅ Tambahkan ini
from kmz_vs import run_boq  # ✅ Tambahkan ini
from sf import run_sf  # ✅ Tambahkan ini

# -------------- ✅ APLIKASI UTAMA ---------------- #
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["user"] = None

if not st.session_state["logged_in"]:
    login_page()
else:
    menu = st.sidebar.radio("📌 Menu", [
        "KML → Jalan",
        "KMZ 🔄 HPDB",
        "KMZ 🔄 DWG CL",
        "KMZ 🔄 BOQ",
        "KMZ 🔄 DWG SF",
        "Urutkan Pole, HP & Clean",  # ✅ Tambahan menu baru
        "Logout"
    ])
    st.sidebar.markdown(f"👤 Logged in as: **{st.session_state['user']}**")

    if menu == "KML 🔄 Jalan":
        run_kml_dxf()
    elif menu == "KMZ 🔄 HPDB":
        run_hpdb(HERE_API_KEY)
    elif menu == "KMZ 🔄 DWG CL":
        run_kmz_to_dwg()
    elif menu == "KMZ 🔄 BOQ":
        run_boq()
    elif menu == "KMZ 🔄 DWG SF":
        run_sf()
    elif menu == "Urutkan Pole, HP & Clean":
            st.markdown(
                """
                <style>
                .btn {
                    display: inline-flex;
                    align-items: center;
                    gap: 8px;
                    padding: 10px 20px;
                    font-size: 16px;
                    border-radius: 8px;
                    background: #4CAF50;
                    color: white;
                    text-decoration: none;
                    font-weight: bold;
                    margin: 5px;
                }
                .btn:hover {
                    background: #45a049;
                }
                .icon {
                    width: 18px;
                    height: 18px;
                    stroke: white;
                    stroke-width: 2;
                    fill: none;
                }
                </style>
                """,
                unsafe_allow_html=True
            )
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(
                    """
                    <a class="btn" href="https://urutkanpole-kingdion.streamlit.app/" target="_blank">
                        <svg xmlns="http://www.w3.org/2000/svg" class="icon" viewBox="0 0 24 24">
                            <path d="M5 13l4 4L19 7"/>
                        </svg>
                        Buka Urutkan Pole & HP
                    </a>
                    """,
                    unsafe_allow_html=True
                )
            
            with col2:
                st.markdown(
                    """
                    <a class="btn" href="https://kmzrapikan-kingdion.streamlit.app/" target="_blank">
                        <svg xmlns="http://www.w3.org/2000/svg" class="icon" viewBox="0 0 24 24">
                            <circle cx="12" cy="12" r="10"/>
                            <path d="M9 12l2 2 4-4"/>
                        </svg>
                        Bersihkan
                    </a>
                    """,
                    unsafe_allow_html=True
                )
                

                                                                    
    elif menu == "Logout":
        st.session_state["logged_in"] = False
        st.session_state["user"] = None
        st.rerun()




















