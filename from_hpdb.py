import streamlit as st
import pandas as pd
import zipfile
from lxml import etree
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
        for part in path.split("/"):
            if len(part) == 3 and part[0] in "ABCD" and part[1:].isdigit():
                return part
        return "UNKNOWN"

    def reverse_here(lat, lon):
        url = f"https://revgeocode.search.hereapi.com/v1/revgeocode?at={lat},{lon}&apikey={HERE_API_KEY}&lang=en-US"
        try:
            r = requests.get(url, timeout=10)
        except Exception:
            return {"district": "", "subdistrict": "", "postalcode": "", "street": ""}
        if r.status_code == 200:
            comp = r.json().get("items", [{}])[0].get("address", {})
            return {
                "district": comp.get("district", "").upper(),
                "subdistrict": comp.get("subdistrict", "").upper().replace("KEL.", "").strip(),
                "postalcode": comp.get("postalCode", "").upper(),
                "street": comp.get("street", "").upper()
            }
        return {"district": "", "subdistrict": "", "postalcode": "", "street": ""}

    # fungsi pola baru
    def generate_pattern(line, capacity, n):
    """
    line: 'A' atau 'B'
    capacity: contoh '24C/2T', '36C/3T', '48C/4T'
    n: jumlah baris yang dibutuhkan
    """
    patterns = []
    tray = 1
    port = 1
    tube = 1
    core = 1

    if capacity in ["24C/2T", "36C/3T", "48C/4T"]:
        max_tray = int(capacity.split("/")[1][0])  # ambil angka sebelum 'T'
        while len(patterns) < n:
            # 2 core per port
            patterns.append({"FDT Tray (Front)": tray, "FDT Port": port, "Tube Colour": tube, "Core Number": core})
            core += 1
            patterns.append({"FDT Tray (Front)": tray, "FDT Port": port, "Tube Colour": tube, "Core Number": core})
            core += 1
            port += 1
            if port > 10:  # tiap tray max 10 port
                port = 1
                tray += 1
                tube += 1
                if tray > max_tray:  # reset kalau tray melebihi kapasitas
                    tray = 1
                    tube = 1
        return patterns[:n]
    else:
        return [{} for _ in range(n)]  # kalau kapasitas tak dikenal


    # ----- Proses utama -----
    if kmz_file and template_file:
        try:
            kmz_bytes = kmz_file.read()
            placemarks = extract_placemarks(kmz_bytes)
        except Exception as e:
            st.error(f"❌ Gagal membaca KMZ: {e}")
            return

        try:
            template_bytes = template_file.read()
            template_df = pd.read_excel(BytesIO(template_bytes))
        except Exception as e:
            st.error(f"❌ Gagal membaca TEMPLATE XLSX: {e}")
            return

        df = template_df.copy()

        fat = placemarks.get("FAT", [])
        hp = placemarks.get("HP COVER", [])
        fdt = placemarks.get("FDT", [])
        all_poles = (
            placemarks.get("NEW POLE 7-3", [])
            + placemarks.get("EXISTING POLE EMR 7-3", [])
            + placemarks.get("EXISTING POLE EMR 7-4", [])
        )

        rc = reverse_here(fdt[0]["lat"], fdt[0]["lon"]) if fdt else {"district": "", "subdistrict": "", "postalcode": "", "street": ""}
        fdtcode = fdt[0]["name"].strip().upper() if fdt else "UNKNOWN"
        oltcode = "UNKNOWN"

        # Cari OLT dari description FDT
        if fdt:
            try:
                with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
                    f = [f for f in z.namelist() if f.lower().endswith(".kml")][0]
                    parser = etree.XMLParser(recover=True)
                    root = etree.parse(z.open(f), parser=parser).getroot()
                    ns = {"kml": "http://www.opengis.net/kml/2.2"}
                    for pm in root.findall(".//kml:Placemark", ns):
                        name_el = pm.find("kml:name", ns)
                        desc_el = pm.find("kml:description", ns)
                        if name_el is not None and name_el.text and name_el.text.strip().upper() == fdtcode:
                            if desc_el is not None and desc_el.text:
                                oltcode = desc_el.text.strip().upper()
                            break
            except Exception:
                oltcode = "UNKNOWN"

        must_cols = ["block", "homenumber", "fdtcode", "oltcode", "fatcode",
                     "Latitude_homepass", "Longitude_homepass", "district", "subdistrict", "postalcode",
                     "FAT ID", "Pole ID", "Pole Latitude", "Pole Longitude", "FAT Address",
                     "Line", "Capacity", "FAT Port",
                     "FDT Tray (Front)", "FDT Port", "Tube Colour", "Core Number"]
        for col in must_cols:
            if col not in df.columns:
                df[col] = ""

        progress = st.progress(0)
        total = len(hp) if hp else 0
        for i, h in enumerate(hp):
            if i >= len(df):
                break
            fc = extract_fatcode(h["path"])
            df.at[i, "fatcode"] = fc

            name_parts = h["name"].split(".")
            if len(name_parts) == 2 and name_parts[0].isalnum() and name_parts[1].isdigit():
                df.at[i, "block"] = name_parts[0].strip().upper()
                df.at[i, "homenumber"] = name_parts[1].strip()
            else:
                df.at[i, "block"] = ""
                df.at[i, "homenumber"] = h["name"]

            df.at[i, "Latitude_homepass"] = h["lat"]
            df.at[i, "Longitude_homepass"] = h["lon"]
            df.at[i, "district"] = rc.get("district", "")
            df.at[i, "subdistrict"] = rc.get("subdistrict", "")
            df.at[i, "postalcode"] = rc.get("postalcode", "")
            df.at[i, "fdtcode"] = fdtcode
            df.at[i, "oltcode"] = oltcode

            hh = reverse_here(h["lat"], h["lon"])
            df.at[i, "street"] = hh.get("street", "").replace("JALAN ", "").strip()

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

            if total > 0:
                progress.progress(int((i + 1) * 100 / total))

        # AUTO FILL FAT PORT
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            for j, idx in enumerate(group.index, start=1):
                df.at[idx, "FAT Port"] = j

        # AUTO FILL LINE & CAPACITY
        fat_capacity_map = {}
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            fatcodes = group["fatcode"].dropna().unique()
            if len(fatcodes) == 0:
                continue
            fc = fatcodes[0]
            letter = fc[0] if fc and fc[0] in "ABCD" else ""
            try:
                nums = [int(x[1:]) for x in fatcodes if len(x) >= 2 and x[1:].isdigit()]
                max_num = max(nums) if nums else 0
            except Exception:
                max_num = 0

            if 1 <= max_num <= 20:
                cap_val = "48C/4T"
            else:
                cap_val = ""

            first_idx = group.index[0]
            df.at[first_idx, "Line"] = letter
            df.at[first_idx, "Capacity"] = cap_val

            fat_capacity_map[fat_id] = cap_val

        # ISI pola baru
        for fat_id, group in df.groupby("FAT ID", sort=False):
            cap_val = fat_capacity_map.get(fat_id, "")
            if not cap_val:
                continue
            line_val = df.at[group.index[0], "Line"]
            pats = generate_pattern(line_val, cap_val, len(group))
            for idx, pat in zip(group.index, pats):
                for k, v in pat.items():
                    df.at[idx, k] = v

        progress.empty()
        st.success("✅ Selesai!")
        st.dataframe(df.head(10))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
