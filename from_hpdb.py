import streamlit as st
import pandas as pd
import zipfile
from lxml import etree   # pakai lxml bukan xml.etree
from io import BytesIO
import requests
from collections import defaultdict

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
        for part in path.split("/"):
            if len(part) == 3 and part[0] in "ABCD" and part[1:].isdigit():
                return part
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

    # Warna urutan hanya 10 core
    CORE_COLORS = ["BLUE", "ORANGE", "GREEN", "BROWN", "SLATE",
                   "WHITE", "RED", "BLACK", "YELLOW", "VIOLET"]

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

        progress = st.progress(0)

        # --- Group HP by FAT ID ---
        grouped = defaultdict(list)
        for h in hp:
            fc = extract_fatcode(h["path"])
            mf = next((x for x in fat if fc in x["name"]), None)
            fat_id = mf["name"] if mf else "FAT_NOT_FOUND"
            grouped[fat_id].append(h)

        new_rows = []
        for fat_idx, (fat_id, hps) in enumerate(grouped.items()):
            line = "LINE A" if ".A" in fat_id[-3:].upper() else ("LINE B" if ".B" in fat_id[-3:].upper() else "")
            capacity = len(hps)
            for port_idx, h in enumerate(hps[:10]):  # hanya 10 port
                row = {}
                row["FAT ID"] = fat_id
                row["FAT Port"] = port_idx + 1
                row["Line"] = line
                row["Capacity"] = f"{capacity}C"
                row["Tube Colour"] = CORE_COLORS[port_idx]
                row["Core Number"] = port_idx + 1
                row["Latitude_homepass"] = h["lat"]
                row["Longitude_homepass"] = h["lon"]
                row["fdtcode"] = fdtcode
                row["oltcode"] = oltcode
                row["fatcode"] = extract_fatcode(h["path"])
                hh = reverse_here(h["lat"], h["lon"])
                row["street"] = hh["street"].replace("JALAN ", "").strip()
                new_rows.append(row)

            progress.progress(int((fat_idx + 1) * 100 / len(grouped)))

        progress.empty()
        out_df = pd.DataFrame(new_rows)
        st.success("✅ Selesai! Data sudah sesuai pola 10 core per FAT")
        st.dataframe(out_df.head(20))
        buf = BytesIO()
        out_df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
