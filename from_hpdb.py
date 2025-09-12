import streamlit as st
import pandas as pd
import zipfile
from lxml import etree   # pakai lxml bukan xml.etree
from io import BytesIO
import requests

def run_hpdb(HERE_API_KEY):

    st.title("📍 KMZ ➜ HPDB (Auto-Pilot⚡)")
    st.markdown("""
<h2>👋 Hai, <span style='color:#0A84FF'>bro</span></h2>
✅ <span style='font-weight:bold;'>CATATAN PENTING :</span><br><br>
1️⃣ <span style='color:#FF6B6B;'>TEMPLATE XLSX</span> harus disesuaikan jumlahnya dengan total homepass dari KMZ.<br>
2️⃣ Block agar terpisah otomatis harus pakai titik, contoh <code>B.1</code> dan <code>A.1</code>.<br>
3️⃣ Fitur otomatis: <span style='color:#34C759;'>FAT ID, Pole ID, Pole Latitude, Pole Longitude, Clustername, street, homenumber, oltcode, fdtcode, fatcode, Latitude_homepass, Longitude_homepass</span>.<br>
4️⃣ OLT CODE agar otomatis, di dalam Description FDT wajib diisi kode OLT.<br>
5️⃣ Street tidak semua bisa terisi otomatis karena ada beberapa jalan di maps bertanda unnamed road.
""", unsafe_allow_html=True)

    if st.button("🔒 Logout"):
        st.session_state["logged_in"] = False
        st.session_state["user"] = None
        st.rerun()

    kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
    template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

    # ------------------------------
    # Extract placemarks pakai lxml
    # ------------------------------
    def extract_placemarks(kmz_bytes):
        def first_lonlat_from_pm(pm, ns):
            for xpath in [
                "./kml:Point/kml:coordinates",
                "./kml:LineString/kml:coordinates",
                "./kml:Polygon//kml:coordinates",
                ".//kml:coordinates"
            ]:
                el = pm.find(xpath, ns)
                if el is None or el.text is None:
                    continue
                txt = " ".join(el.text.split())
                tokens = [t for t in txt.split(" ") if "," in t]
                if not tokens:
                    continue
                parts = tokens[0].split(",")
                if len(parts) >= 2:
                    try:
                        lon = float(parts[0])
                        lat = float(parts[1])
                        return lon, lat
                    except ValueError:
                        continue
            return None

        def recurse_folder(folder, ns, path=""):
            items = []
            name_el = folder.find("kml:name", ns)
            folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
            new_path = f"{path}/{folder_name}" if path else folder_name
            for sub in folder.findall("kml:Folder", ns):
                items += recurse_folder(sub, ns, new_path)
            for pm in folder.findall("kml:Placemark", ns):
                nm = pm.find("kml:name", ns)
                if nm is None:
                    continue
                first = first_lonlat_from_pm(pm, ns)
                if first is None:
                    continue
                lon, lat = first
                items.append({
                    "name": nm.text.strip(),
                    "lat": lat,
                    "lon": lon,
                    "path": new_path
                })
            return items

        with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
            f = [f for f in z.namelist() if f.lower().endswith(".kml")][0]
            parser = etree.XMLParser(recover=True)
            root = etree.parse(z.open(f), parser=parser).getroot()
            ns = {"kml": "http://www.opengis.net/kml/2.2"}
            all_pm = []
            for folder in root.findall(".//kml:Folder", ns):
                all_pm += recurse_folder(folder, ns)
            data = {k: [] for k in [
                "FAT", "NEW POLE 7-3", "EXISTING POLE EMR 7-3",
                "EXISTING POLE EMR 7-4", "FDT", "HP COVER"
            ]}
            for p in all_pm:
                for k in data:
                    if k in p["path"]:
                        data[k].append(p)
                        break
            return data

    def extract_fatcode(path):
        # cari token seperti A01, B12, C03, D20 pada path folder
        for part in path.split("/"):
            p = part.strip().upper()
            if len(p) >= 3 and p[0] in "ABCD" and p[1:].isdigit():
                # ambil hanya 3 karakter A01 style (tapi jika ada A001, ambil A01)
                code = p[:3]
                return code
        return "UNKNOWN"

    def reverse_here(lat, lon):
        url = f"https://revgeocode.search.hereapi.com/v1/revgeocode?at={lat},{lon}&apikey={HERE_API_KEY}&lang=en-US"
        r = requests.get(url)
        if r.status_code == 200:
            comp = r.json().get("items", [{}])[0].get("address", {})
            return {
                "district": comp.get("district", "").upper(),
                "subdistrict": comp.get("subdistrict", "").upper().replace("KEL.", "").strip(),
                "postalcode": comp.get("postalCode", "").upper(),
                "street": comp.get("street", "").upper()
            }
        return {"district": "", "subdistrict": "", "postalcode": "", "street": ""}

    if kmz_file and template_file:
        kmz_bytes = kmz_file.read()
        placemarks = extract_placemarks(kmz_bytes)
        df = pd.read_excel(template_file)
        fat = placemarks["FAT"]
        hp = placemarks["HP COVER"]
        fdt = placemarks["FDT"]
        all_poles = (
            placemarks["NEW POLE 7-3"]
            + placemarks["EXISTING POLE EMR 7-3"]
            + placemarks["EXISTING POLE EMR 7-4"]
        )

        rc = reverse_here(fdt[0]["lat"], fdt[0]["lon"]) if fdt else {"district": "", "subdistrict": "", "postalcode": "", "street": ""}
        fdtcode = fdt[0]["name"].strip().upper() if fdt else "UNKNOWN"
        oltcode = "UNKNOWN"

        # Cari OLT dari description FDT
        if fdt:
            with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
                f = [f for f in z.namelist() if f.lower().endswith(".kml")][0]
                parser = etree.XMLParser(recover=True)
                root = etree.parse(z.open(f), parser=parser).getroot()
                ns = {"kml": "http://www.opengis.net/kml/2.2"}
                for pm in root.findall(".//kml:Placemark", ns):
                    name_el = pm.find("kml:name", ns)
                    desc_el = pm.find("kml:description", ns)
                    if name_el is not None and name_el.text.strip().upper() == fdtcode:
                        if desc_el is not None:
                            oltcode = desc_el.text.strip().upper()
                        break

        # ----------------------------
        # Tambahkan kolom baru bila belum ada
        # ----------------------------
        new_cols = [
            "block", "homenumber", "fdtcode", "oltcode",
            "FAT Port", "FDT Tray (Front)", "FDT Port",
            "Line", "Capacity", "Tube Colour", "Core Number",
            "Latitude_homepass", "Longitude_homepass",
            "district", "subdistrict", "postalcode",
            "FAT ID", "Pole Latitude", "Pole Longitude", "Pole ID", "FAT Address"
        ]
        for col in new_cols:
            if col not in df.columns:
                df[col] = ""

        # definisi warna tube (siklus)
        tube_colours = ["BLUE", "ORANGE", "GREEN", "BROWN", "SLATE",
                        "WHITE", "RED", "BLACK", "YELLOW", "VIOLET",
                        "PINK", "AQUA"]

        progress = st.progress(0)
        total = len(hp)

        for i, h in enumerate(hp):
            if i >= len(df):
                break

            fc = extract_fatcode(h["path"])
            df.at[i, "fatcode"] = fc

            # ----- EXTRACT/block & homenumber seperti kode lama -----
            name_parts = h["name"].split(".")
            if len(name_parts) == 2 and name_parts[0].isalnum() and name_parts[1].isdigit():
                df.at[i, "block"] = name_parts[0].strip().upper()
                df.at[i, "homenumber"] = name_parts[1].strip()
            else:
                df.at[i, "block"] = ""
                df.at[i, "homenumber"] = h["name"]

            df.at[i, "Latitude_homepass"] = h["lat"]
            df.at[i, "Longitude_homepass"] = h["lon"]
            df.at[i, "district"] = rc["district"]
            df.at[i, "subdistrict"] = rc["subdistrict"]
            df.at[i, "postalcode"] = rc["postalcode"]
            df.at[i, "fdtcode"] = fdtcode
            df.at[i, "oltcode"] = oltcode

            hh = reverse_here(h["lat"], h["lon"])
            df.at[i, "street"] = hh["street"].replace("JALAN ", "").strip()

            # ----- Cari FAT terdekat berdasarkan fatcode (nama FAT mengandung kode) -----
            mf = next((x for x in fat if fc in x["name"]), None)
            if mf:
                df.at[i, "FAT ID"] = mf["name"]
                df.at[i, "Pole Latitude"] = mf["lat"]
                df.at[i, "Pole Longitude"] = mf["lon"]
                pol = next(
                    (p["name"] for p in all_poles if abs(p["lat"] - mf["lat"]) < 1e-4 and abs(p["lon"] - mf["lon"]) < 1e-4),
                    "POLE_NOT_FOUND"
                )
                df.at[i, "Pole ID"] = pol
                fataddr = reverse_here(mf["lat"], mf["lon"])["street"]
                df.at[i, "FAT Address"] = fataddr
            else:
                df.at[i, "FAT ID"] = "FAT_NOT_FOUND"
                df.at[i, "Pole ID"] = "POLE_NOT_FOUND"
                df.at[i, "FAT Address"] = ""

            # ----------------------------
            # LOGIKA BARU: FAT Port, FDT Tray, FDT Port, Line, Capacity, Tube Colour, Core Number
            # ----------------------------
            # Ambil nomor port dari fatcode jika bisa (A01 -> 1, B12 -> 12)
            fat_port_num = None
            if fc and len(fc) >= 2 and fc[1:].isdigit():
                try:
                    fat_port_num = int(fc[1:])
                except:
                    fat_port_num = None

            # FAT Port: gunakan angka dari fatcode jika valid, else fallback ke urutan i+1
            if fat_port_num is not None:
                df.at[i, "FAT Port"] = fat_port_num
            else:
                df.at[i, "FAT Port"] = (i % 20) + 1  # fallback 1..20

            # FDT Tray (Front) dan FDT Port = siklus 1..10 berulang
            # Prefer menggunakan FAT Port jika ada: map fat_port_num ke 1..10 siklus
            if fat_port_num is not None:
                tray_port_val = ((fat_port_num - 1) % 10) + 1
            else:
                tray_port_val = (i % 10) + 1

            df.at[i, "FDT Tray (Front)"] = tray_port_val
            df.at[i, "FDT Port"] = tray_port_val

            # Line berdasarkan huruf depan fatcode (A/B/C/D)
            if fc and fc[0] in ["A", "B", "C", "D"]:
                df.at[i, "Line"] = f"Line {fc[0]}"
            else:
                df.at[i, "Line"] = "UNKNOWN"

            # Capacity berdasarkan range nomor di fatcode
            cap = "UNKNOWN"
            if fat_port_num is not None:
                if 1 <= fat_port_num <= 10:
                    cap = "24C/2T"
                elif 11 <= fat_port_num <= 15:
                    cap = "36C/3T"
                elif 16 <= fat_port_num <= 20:
                    cap = "48C/4T"
            df.at[i, "Capacity"] = cap

            # Tube Colour: siklus berdasarkan FDT Tray (Front)
            # (Tray 1 -> first colour, Tray 2 -> second, dll)
            tray_idx = int(df.at[i, "FDT Tray (Front)"]) - 1
            df.at[i, "Tube Colour"] = tube_colours[tray_idx % len(tube_colours)]

            # Core Number: berpasangan berdasarkan FAT Port nomor
            # Aturan: port 1 -> cores "1,2"; port 2 -> "3,4"; port 3 -> "5,6" dst.
            # Kita buat pasangan berdasarkan urutan natural FAT Port:
            if fat_port_num is not None:
                pair_start = (fat_port_num - 1) * 2 - ((fat_port_num - 1) // 10) * 20
                # penyesuaian jika fat_port_num 1 -> 1,2; 2 -> 3,4; dst.
                # tetapi untuk port > 10, akan menghasilkan lebih besar dari 20; 
                # jika ingin siklus 1..20, gunakan mapping alternatif:
                # simpler: buat pasangan linear: port_n -> ((n-1)*2+1, (n-1)*2+2)
                start = (fat_port_num - 1) * 2 + 1
                core_pair = f"{start},{start+1}"
            else:
                # fallback: pairing berdasarkan tray_port_val
                start = (tray_port_val - 1) * 2 + 1
                core_pair = f"{start},{start+1}"

            df.at[i, "Core Number"] = core_pair

            # update progress
            progress.progress(int((i + 1) * 100 / max(1, total)))

        progress.empty()
        st.success("✅ Selesai!")
        st.dataframe(df.head(20))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")


# Jika mau jalanin langsung (contoh):
if __name__ == "__main__":
    # Masukkan API KEY HERE-mu di sini jika mau test lokal
    HERE_API_KEY = st.secrets.get("HERE_API_KEY", "") if "st" in globals() else ""
    run_hpdb(HERE_API_KEY)
