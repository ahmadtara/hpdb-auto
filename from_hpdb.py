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

        progress = st.progress(0)
        total = len(hp)

        # pastikan kolom wajib ada
        must_cols = ["block", "homenumber", "fdtcode", "oltcode", "fatcode",
                     "Latitude_homepass", "Longitude_homepass", "district", "subdistrict", "postalcode",
                     "FAT ID", "Pole ID", "Pole Latitude", "Pole Longitude", "FAT Address",
                     "Line", "Capacity", "FAT Port", "FDT Tray (Front)", "FDT Port", "Tube Colour", "Core Number"]
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

        # ====== AUTO FILL LINE, CAPACITY (tetap seperti semula) ======
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            fatcodes = group["fatcode"].dropna().unique()
            if len(fatcodes) == 0:
                continue
            fc = fatcodes[0]
            letter = fc[0] if fc and fc[0] in "ABCD" else ""
            try:
                nums = [int(x[1:]) for x in fatcodes if len(x) >= 2]
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

        # ====== AUTO FILL FDT Tray, FDT Port, Tube Colour (angka), Core Number ====
        # Pola persis berdasarkan database (foto ketiga):
        # - 2 FAT Port -> 1 FDT Port (FDT Port = ceil(FAT Port/2))
        # - 1 TRAY = 10 FDT Port (port 1..10)
        # - Dalam tray: ports 1-5 => tube slot 1 ; ports 6-10 => tube slot 2
        # - Tube numbering pattern (sesuai foto):
        #     -> tray odd  : tube numbers 1 & 2
        #     -> tray even : tube numbers 3 & 4
        #   (pattern ini berulang)
        # - Core number di dalam tube: tiap FDT Port punya 2 core: port1->cores1&2, port2->3&4, ... sampai 10
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue

            # process each row (row = satu FAT Port)
            for idx in group.index:
                # Ambil FAT Port yang sudah diisi sebelumnya
                fat_port_raw = df.at[idx, "FAT Port"]
                try:
                    fat_port_val = int(fat_port_raw)
                except Exception:
                    # fallback: gunakan posisi dalam group (1-based)
                    local_pos = list(group.index).index(idx) + 1
                    fat_port_val = local_pos

                # FDT Port = ceil(FAT Port / 2)
                fdt_port_num = (fat_port_val + 1) // 2

                # Tray number (1-based), setiap 10 FDT Port = 1 tray
                tray_num = (fdt_port_num - 1) // 10 + 1

                # Port number IN TRAY (1..10)
                port_in_tray = ((fdt_port_num - 1) % 10) + 1

                # Tube slot in tray (1 or 2)
                tube_slot_in_tray = ((port_in_tray - 1) // 5) + 1  # 1 or 2

                # Tube numbering PATTERN per foto:
                # tray odd -> tube numbers 1 & 2
                # tray even -> tube numbers 3 & 4
                if tray_num % 2 == 1:
                    tube_number = tube_slot_in_tray  # 1 or 2
                else:
                    tube_number = 2 + tube_slot_in_tray  # 3 or 4

                # Port position inside tube (1..5)
                port_pos_in_tube = ((port_in_tray - 1) % 5) + 1

                # offset in port: odd FAT Port -> first core of the port, even FAT Port -> second core
                offset_in_port = 1 if (fat_port_val % 2 == 1) else 2

                # Core number within tube (1..10)
                core_number = (port_pos_in_tube - 1) * 2 + offset_in_port

                # Write Tube Colour (angka) and Core Number for every FAT Port row
                df.at[idx, "Tube Colour"] = int(tube_number)
                df.at[idx, "Core Number"] = int(core_number)

                # Only write FDT Tray & FDT Port on the FIRST row of the pair (i.e., when FAT Port is odd)
                if fat_port_val % 2 == 1:
                    df.at[idx, "FDT Tray (Front)"] = int(tray_num)
                    df.at[idx, "FDT Port"] = int(port_in_tray)
                else:
                    # keep blank for second row of the pair
                    df.at[idx, "FDT Tray (Front)"] = ""
                    df.at[idx, "FDT Port"] = ""

        progress.empty()
        st.success("✅ Selesai!")
        st.dataframe(df.head(40))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
