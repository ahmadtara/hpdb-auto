import streamlit as st
import pandas as pd
import zipfile
from lxml import etree
from io import BytesIO
import requests
import re

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

    # ------------------------------
    # helper: parse jumlah tube dari Capacity string (ex: "24C/2T" -> 2)
    # ------------------------------
    def parse_tubes_from_capacity(capacity_str):
        if not isinstance(capacity_str, str):
            return None
        m = re.search(r'(\d+)\s*T', capacity_str.upper())
        if m:
            try:
                return int(m.group(1))
            except:
                return None
        # fallback: try after slash like /2T
        m2 = re.search(r'/\s*(\d+)\s*T', capacity_str.upper())
        if m2:
            try:
                return int(m2.group(1))
            except:
                return None
        return None

    # ------------------------------
    # build global mapping (tray 1..4, per tray ports 1..10, per port two core entries)
    # pattern matches mapping yang kamu kasih sebelumnya
    # ------------------------------
    def build_global_mapping():
        mapping = []
        for tray in range(1, 5):  # 1..4
            # tray 1 & 3 use tube colours 1 & 2; tray 2 & 4 use tube colours 3 & 4
            if tray in (1, 3):
                tube_pair = [1, 2]
            else:
                tube_pair = [3, 4]
            # for each tube in the pair: first tube -> ports 1..5; second -> ports 6..10
            for t_idx, tube in enumerate(tube_pair):
                for j in range(1, 6):  # j = 1..5
                    port = t_idx * 5 + j  # gives 1..5 for first tube, 6..10 for second tube
                    core1 = (j - 1) * 2 + 1
                    core2 = core1 + 1
                    mapping.append((tray, port, tube, core1))
                    mapping.append((tray, port, tube, core2))
        return mapping  # total 4 trays * (2 tubes * 5 ports * 2 cores) = 80 entries

    global_mapping = build_global_mapping()

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
        total = len(hp)

        # pastikan kolom wajib ada
        must_cols = ["block", "homenumber", "fdtcode", "oltcode", "fatcode",
                     "Latitude_homepass", "Longitude_homepass", "district", "subdistrict", "postalcode",
                     "FAT ID", "Pole ID", "Pole Latitude", "Pole Longitude", "FAT Address",
                     "Line", "Capacity", "FAT Port",
                     "FDT Tray (Front)", "FDT Port", "Tube Colour", "Core Number"]
        for col in must_cols:
            if col not in df.columns:
                df[col] = ""

        for i, h in enumerate(hp):
            if i >= len(df):
                break
            fc = extract_fatcode(h["path"])
            df.at[i, "fatcode"] = fc

            # parsing block & homenumber
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
                df.at[i, "FAT Address"] = reverse_here(mf["lat"], mf["lon"])["street"]
            else:
                df.at[i, "FAT ID"] = "FAT_NOT_FOUND"
                df.at[i, "Pole ID"] = "POLE_NOT_FOUND"
                df.at[i, "FAT Address"] = ""

            progress.progress(int((i + 1) * 100 / max(1, len(hp))))

        # ====== AUTO FILL FAT PORT ======
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            for i, idx in enumerate(group.index, start=1):
                df.at[idx, "FAT Port"] = i

        # ====== AUTO FILL LINE & CAPACITY PER FAT ID ======
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            fatcodes = group["fatcode"].dropna().unique()
            if len(fatcodes) == 0:
                continue
            fc = fatcodes[0]
            letter = fc[0] if fc and fc[0] in "ABCD" else ""
            try:
                nums = [int(fc[1:]) for fc in fatcodes if len(fc) >= 2]
                max_num = max(nums) if nums else 0
            except:
                max_num = 0

            if 1 <= max_num <= 10:
                cap_val = "24C/2T"
            elif 11 <= max_num <= 15:
                cap_val = "36C/3T"
            elif 16 <= max_num <= 20:
                cap_val = "48C/4T"
            else:
                cap_val = ""

            first_idx = group.index[0]
            df.at[first_idx, "Line"] = letter
            df.at[first_idx, "Capacity"] = cap_val

        # ====== FILL FDT TRAY/PORT/TUBE/CORE BERDASARKAN LINE & CAPACITY ======
        # Precompute a per-fat mapping based on global_mapping, line, and capacity
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue

            # read line and capacity for this FAT
            line_val = group["Line"].iloc[0] if "Line" in group.columns else ""
            capacity_val = group["Capacity"].iloc[0] if "Capacity" in group.columns else ""

            # normalize line letter
            line_letter = (str(line_val).strip().upper() or "A")
            if line_letter not in list("ABCD"):
                line_letter = "A"

            # parse tube count from capacity (2T, 3T, 4T)
            tubes_needed = parse_tubes_from_capacity(capacity_val) or 2
            if tubes_needed < 1:
                tubes_needed = 2
            # clamp to max 4 (we only have 4 tube colours in mapping)
            tubes_needed = min(tubes_needed, 4)

            # starting tube based on Line letter: A->1, B->2, C->3, D->4
            start_tube = ord(line_letter) - ord("A") + 1

            # allowed tube colours for this FAT (wrap-around)
            allowed_tubes = [((start_tube - 1 + i) % 4) + 1 for i in range(tubes_needed)]

            # filter global_mapping to only entries that use allowed_tubes
            filtered = [m for m in global_mapping if m[2] in allowed_tubes]

            # rotate filtered so first entry starts at first occurrence of start_tube (if present)
            rot_index = next((i for i, mm in enumerate(filtered) if mm[2] == start_tube), 0)
            if rot_index:
                filtered = filtered[rot_index:] + filtered[:rot_index]

            # safety fallback: if filtered empty (shouldn't happen), use global mapping
            if not filtered:
                filtered = global_mapping.copy()

            # assign to rows in group in their order (group.index preserves order)
            for j, idx in enumerate(group.index):
                # pick mapping entry cycling through filtered if needed
                map_entry = filtered[j % len(filtered)]
                tray_val, fdt_port_val, tube_val, core_val = map_entry
                df.at[idx, "FDT Tray (Front)"] = tray_val
                df.at[idx, "FDT Port"] = fdt_port_val
                df.at[idx, "Tube Colour"] = tube_val
                df.at[idx, "Core Number"] = core_val

        progress.empty()
        st.success("✅ Selesai!")
        st.dataframe(df.head(10))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
