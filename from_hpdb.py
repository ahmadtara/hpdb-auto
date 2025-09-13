import streamlit as st
import pandas as pd
import zipfile
from lxml import etree
from io import BytesIO
import requests
import math

def run_hpdb(HERE_API_KEY):

    st.set_page_config(page_title="HPDB Auto-Pilot", layout="wide")
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

    # Option: enable tray auto-step if template lacks tray pattern or user wants override
    col1, col2 = st.columns([1, 1])
    with col1:
        enable_tray_auto = st.checkbox("🟩 Enable Tray Auto-Step (override template FDT Tray)", value=False,
                                       help="Jika aktif, nomor FDT Tray akan bertambah setiap N baris sesuai 'Tray Step'. Gunakan bila template tidak berisi pola tray atau ingin override.")
    with col2:
        tray_step = st.number_input("Tray Step (N)", min_value=1, value=10, help="Jumlah baris per tray sebelum naik ke nomor tray berikutnya (default 10).")

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
            kml_files = [f for f in z.namelist() if f.lower().endswith(".kml")]
            if not kml_files:
                raise ValueError("KMZ does not contain any .kml file.")
            f = kml_files[0]
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

    # ----- Proses utama -----
    if kmz_file and template_file:
        # read KMZ
        try:
            kmz_bytes = kmz_file.read()
            placemarks = extract_placemarks(kmz_bytes)
        except Exception as e:
            st.error(f"❌ Gagal membaca KMZ: {e}")
            return

        # read template (uploaded file as source of patterns)
        try:
            template_bytes = template_file.read()
            template_df = pd.read_excel(BytesIO(template_bytes), sheet_name=0)
        except Exception as e:
            st.error(f"❌ Gagal membaca TEMPLATE XLSX: {e}")
            return

        # base df: copy of template_df (matching flow sebelumnya)
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
                    kml_files = [f for f in z.namelist() if f.lower().endswith(".kml")]
                    if kml_files:
                        parser = etree.XMLParser(recover=True)
                        root = etree.parse(z.open(kml_files[0]), parser=parser).getroot()
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

        # ensure must cols exist
        must_cols = ["block", "homenumber", "fdtcode", "oltcode", "fatcode",
                     "Latitude_homepass", "Longitude_homepass", "district", "subdistrict", "postalcode",
                     "FAT ID", "Pole ID", "Pole Latitude", "Pole Longitude", "FAT Address",
                     "Line", "Capacity", "FAT Port",
                     "FDT Tray (Front)", "FDT Port", "Tube Colour", "Core Number"]
        for col in must_cols:
            if col not in df.columns:
                df[col] = ""

        # ---------------------------
        # build patterns grouped by Capacity from the template sheet
        # approach: find rows where Capacity is present -> that capacity's pattern rows are subsequent rows
        # Simpler robust approach: group by Capacity column directly: take rows where Capacity==cap and collect pattern cols.
        # That works if template has rows with Capacity repeated (or filled on each pattern row).
        # We'll support both: if Capacity is set on a row, that'll be group key; if not, fallback to carry-over logic.
        # ---------------------------
        pattern_cols = ["FDT Tray (Front)", "FDT Port", "Tube Colour", "Core Number"]
        patterns_by_capacity = {}

        # First try: group rows where Capacity value equals cap (non-null)
        if "Capacity" in template_df.columns:
            # If template has Capacity repeated on pattern rows, this captures it easily
            grouped = template_df[template_df["Capacity"].notna()].groupby("Capacity")
            for cap, _ in grouped:
                # select all template rows where Capacity == cap (including the ones where capacity only appears once)
                mask = (template_df["Capacity"] == cap)
                # But sometimes capacity appears only at top of block; collect following rows until next capacity. Use carry approach too.
                rows = template_df[mask]
                if not rows.empty:
                    # get indices for block start
                    start_idx = rows.index[0]
                    # find next capacity index
                    all_cap_idx = template_df[template_df["Capacity"].notna()].index.tolist()
                    pos = all_cap_idx.index(start_idx)
                    end_idx = template_df.shape[0] - 1
                    if pos + 1 < len(all_cap_idx):
                        end_idx = all_cap_idx[pos + 1] - 1
                    block = template_df.loc[start_idx:end_idx]
                    # collect pattern rows (drop rows with all pattern cols NaN)
                    block_patterns = block[pattern_cols].dropna(how="all")
                    if not block_patterns.empty:
                        patterns_by_capacity[str(cap).strip()] = block_patterns.reset_index(drop=True)
            # If patterns_by_capacity empty, fallback to carry-over scan:
        if not patterns_by_capacity:
            # carry-over scan (previous logic)
            current_cap = None
            for _, row in template_df.iterrows():
                cap_cell = None
                if "Capacity" in template_df.columns:
                    cap_cell = row.get("Capacity")
                if pd.notna(cap_cell) and str(cap_cell).strip() != "":
                    current_cap = str(cap_cell).strip()
                    patterns_by_capacity.setdefault(current_cap, [])
                if current_cap is None:
                    continue
                # check pattern cols
                any_val = False
                values = []
                for pc in pattern_cols:
                    if pc in template_df.columns:
                        v = row.get(pc)
                        if pd.notna(v):
                            any_val = True
                            values.append(v if not isinstance(v, str) else v.strip())
                        else:
                            values.append("")
                    else:
                        values.append("")
                if any_val:
                    patterns_by_capacity[current_cap].append(tuple(values))
            # convert tuple-lists to DataFrame for uniformity
            for k, v in list(patterns_by_capacity.items()):
                if isinstance(v, list):
                    patterns_by_capacity[k] = pd.DataFrame(v, columns=pattern_cols)

        # If still empty -> warn
        if not patterns_by_capacity:
            st.info("ℹ️ Tidak ditemukan pola FDT (Tray/Port/Tube/Core) dalam template berdasarkan kolom Capacity. Pastikan template berisi contoh urutan di kolom Capacity dan kolom pola terisi di baris yang relevan.")

        # ====== isi baris berdasarkan HP list (hanya isi atribut dasar dulu) ======
        progress = st.progress(0)
        total = len(hp) if hp else 0
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

        # ====== AUTO FILL FAT PORT (urut per FAT ID) ======
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            for j, idx in enumerate(group.index, start=1):
                df.at[idx, "FAT Port"] = j

        # ====== AUTO FILL LINE & CAPACITY PER FAT ID ======
        # also keep capacity value per FAT (so we can use it for pattern selection)
        fat_capacity_map = {}  # fat_id -> capacity string
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

            # record capacity (cap_val may be empty string if not matched)
            fat_capacity_map[fat_id] = cap_val

        # ====== Sekarang: isi pola FDT (Tray/Port/Tube/Core) berdasarkan Capacity per FAT ======
        missing_caps = set()
        for fat_id, group in df.groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            cap_val = fat_capacity_map.get(fat_id, "")
            # fallback: take any Capacity inside group rows if cap_val empty
            if not cap_val:
                caps = group["Capacity"].dropna().unique()
                if len(caps) > 0:
                    cap_val = str(caps[0]).strip()
            if not cap_val:
                missing_caps.add(fat_id)
                continue

            # get pattern DataFrame (if exists)
            pattern_df = patterns_by_capacity.get(cap_val)
            # if not found, try matching trimmed representation (e.g. numeric/str differences)
            if pattern_df is None:
                # try fallback key matches (strip/upper)
                for k in list(patterns_by_capacity.keys()):
                    if str(k).strip().upper() == str(cap_val).strip().upper():
                        pattern_df = patterns_by_capacity[k]
                        break

            if pattern_df is None or pattern_df.empty:
                # no pattern available for this capacity
                # if user enabled tray auto-step, we'll still fill tray/port with generated values
                if cap_val not in missing_caps:
                    missing_caps.add(cap_val)
                # proceed to generate defaults below
                pattern_df = None

            # assign pattern entries in order to group rows (circular) OR generate defaults
            n = len(group)
            if pattern_df is not None:
                # repeat pattern to match n rows
                rep_count = math.ceil(n / len(pattern_df))
                pola_repeat = pd.concat([pattern_df] * rep_count, ignore_index=True).iloc[:n]
                for j, idx in enumerate(group.index):
                    df.at[idx, "FDT Tray (Front)"] = pola_repeat.iat[j, pattern_df.columns.get_loc("FDT Tray (Front)")]
                    df.at[idx, "FDT Port"] = pola_repeat.iat[j, pattern_df.columns.get_loc("FDT Port")]
                    df.at[idx, "Tube Colour"] = pola_repeat.iat[j, pattern_df.columns.get_loc("Tube Colour")]
                    df.at[idx, "Core Number"] = pola_repeat.iat[j, pattern_df.columns.get_loc("Core Number")]
            else:
                # generate defaults:
                # - FDT Port: sequential 1..n
                # - Tube Colour / Core Number: left blank
                # - FDT Tray: if enable_tray_auto -> 1 + floor((j)/tray_step) else = 1
                for j, idx in enumerate(group.index):
                    # port sequential
                    df.at[idx, "FDT Port"] = str(j + 1)
                    df.at[idx, "Tube Colour"] = ""
                    df.at[idx, "Core Number"] = ""
                    if enable_tray_auto:
                        tray_num = 1 + (j // int(max(1, tray_step)))
                        df.at[idx, "FDT Tray (Front)"] = str(tray_num)
                    else:
                        df.at[idx, "FDT Tray (Front)"] = "1"

        if missing_caps:
            st.warning(f"⚠️ Tidak ditemukan pola di template untuk Capacity berikut (digenerate default / tidak diisi pola): {', '.join(map(str, sorted(missing_caps)))}. Kamu bisa tambahkan pola di template atau aktifkan 'Tray Auto-Step' untuk generate tray otomatis.")

        progress.empty()
        st.success("✅ Selesai!")
        st.dataframe(df.head(12))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
