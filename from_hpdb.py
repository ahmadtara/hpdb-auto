import streamlit as st
import pandas as pd
import zipfile
from lxml import etree
from io import BytesIO
import requests
import math
import numpy as np

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

    # tambahan UI: toggle snap ke jalan
    snap_to_road = st.checkbox("📍 Aktifkan Snap ke Jalan Terdekat (memperbaiki kolom street / AE)", value=False)
    use_smart_fill = st.checkbox("🧠 Isi otomatis jalan kosong dari tetangga (Smart Fill)", value=True)

    # ------------------------------ #
    #  Extract placemarks from KMZ   #
    # ------------------------------ #
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
                desc_el = pm.find("kml:description", ns)
                desc_text = desc_el.text.strip() if (desc_el is not None and desc_el.text) else ""
                items.append({
                    "name": nm.text.strip(),
                    "lat": lat,
                    "lon": lon,
                    "path": new_path,
                    "description": desc_text
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
                "FAT",
                "NEW POLE 7-3", "NEW POLE 7-4", "NEW POLE 9-4",
                "EXISTING POLE EMR 7-3", "EXISTING POLE EMR 7-4", "EXISTING POLE EMR 9-4",
                "FDT", "HP COVER"
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


    def find_nearest_pole(fat, poles):
        fx, fy = fat["lat"], fat["lon"]
        nearest = None
        min_dist = float("inf")
        for p in poles:
            dist = math.hypot(p["lat"] - fx, p["lon"] - fy)
            if dist < min_dist:
                min_dist = dist
                nearest = p
        return nearest

    # Cached reverse geocode using rounded coords to reduce calls
    reverse_cache = {}
    session = requests.Session()

    def reverse_here(lat, lon):
        # round to 6 decimals to unify very-close coordinates
        key = (round(float(lat), 6), round(float(lon), 6))
        if key in reverse_cache and isinstance(reverse_cache[key], dict) and "rev" in reverse_cache[key]:
            return reverse_cache[key]["rev"]
        url = f"https://revgeocode.search.hereapi.com/v1/revgeocode?at={key[0]},{key[1]}&apikey={HERE_API_KEY}&lang=en-US"
        try:
            r = session.get(url, timeout=8)
            if r.status_code == 200:
                comp = r.json().get("items", [{}])[0].get("address", {})
                value = {
                    "district": comp.get("district", "").upper(),
                    "subdistrict": comp.get("subdistrict", "").upper().replace("KEL.", "").strip(),
                    "postalcode": comp.get("postalCode", "").upper(),
                    "street": comp.get("street", "").upper()
                }
                # store into cache under key, preserving any 'snap' entry too
                prev = reverse_cache.get(key, {})
                prev["rev"] = value
                reverse_cache[key] = prev
                return value
        except Exception:
            pass
        empty = {"district": "", "subdistrict": "", "postalcode": "", "street": ""}
        prev = reverse_cache.get(key, {})
        prev["rev"] = empty
        reverse_cache[key] = prev
        return empty

    # ------------------------------ #
    #   Snap-to-road function (HERE)
    # ------------------------------ #
    def snap_to_nearest_road(lat, lon):
        """
        Gunakan HERE Snap-to-Road (nearest) endpoint.
        Jika berhasil, return dict {"lat":..., "lon":..., "street": "..."}
        Jika gagal, return {"lat": original, "lon": original, "street": ""}
        Hasil disimpan ke reverse_cache agar tidak ulang panggil.
        """
        key = (round(float(lat), 6), round(float(lon), 6))
        # cek cache dulu (jika ada snapped hasil sebelumnya)
        if key in reverse_cache and isinstance(reverse_cache[key], dict) and "snap" in reverse_cache[key]:
            return reverse_cache[key]["snap"]

        # endpoint snap-to-road (HERE)
        # NOTE: endpoint ini tersedia di beberapa paket HERE; jika tidak tersedia, fungsi akan fallback
        url = f"https://snap-to-road.hereapi.com/v1/nearest?apikey={HERE_API_KEY}&prox={lat},{lon}"
        try:
            r = session.get(url, timeout=8)
            if r.status_code == 200:
                data = r.json()
                # struktur respons: items -> first item -> position, address (street)
                if isinstance(data, dict) and data.get("items"):
                    itm = data["items"][0]
                    pos = itm.get("position", {})
                    addr = itm.get("address", {})
                    snapped_lat = pos.get("lat", lat)
                    snapped_lon = pos.get("lng", lon)
                    street_name = addr.get("street", "") or ""
                    res = {"lat": snapped_lat, "lon": snapped_lon, "street": street_name.upper()}
                    prev = reverse_cache.get(key, {})
                    prev["snap"] = res
                    reverse_cache[key] = prev
                    return res
        except Exception:
            pass

        # jika gagal, simpan hasil kosong agar tidak retry berulang
        res = {"lat": lat, "lon": lon, "street": ""}
        prev = reverse_cache.get(key, {})
        prev["snap"] = res
        reverse_cache[key] = prev
        return res

        # =====================================
    # SMART FILL JALAN KOSONG (v1.0)
    # =====================================
    def smart_fill_streets(street_list):
        filled = street_list.copy()
        for i in range(len(filled)):
            if not filled[i] or str(filled[i]).strip() == "":
                prev_street = filled[i - 1] if i > 0 else ""
                next_street = filled[i + 1] if i < len(filled) - 1 else ""
                if prev_street:
                    filled[i] = prev_street
                elif next_street:
                    filled[i] = next_street
                else:
                    filled[i] = "TANPA NAMA"
        return filled


    if kmz_file and template_file:
        kmz_bytes = kmz_file.read()
        placemarks = extract_placemarks(kmz_bytes)

        df = pd.read_excel(template_file, sheet_name="Homepass Database")
        expected_cols = ["FDT Tray (Front)", "FDT Port", "Tube Colour", "Core Number"]
        mapping_df = df[expected_cols].dropna(how="all").reset_index(drop=True)

        fat = placemarks["FAT"]
        hp = placemarks["HP COVER"]
        fdt = placemarks["FDT"]

        all_poles = (
            placemarks["NEW POLE 7-3"]
            + placemarks["NEW POLE 7-4"]
            + placemarks["NEW POLE 9-4"]
            + placemarks["EXISTING POLE EMR 7-3"]
            + placemarks["EXISTING POLE EMR 7-4"]
            + placemarks["EXISTING POLE EMR 9-4"]
        )

        # fdt reverse geocode base values
        if fdt:
            fdt_item = fdt[0]
            rc = reverse_here(fdt_item["lat"], fdt_item["lon"])
            fdtcode = fdt_item["name"].strip().upper()
            # use description captured earlier for oltcode if present, else UNKNOWN
            oltcode = fdt_item.get("description", "").strip().upper() or "UNKNOWN"
        else:
            rc = {"district": "", "subdistrict": "", "postalcode": "", "street": ""}
            fdtcode = "UNKNOWN"
            oltcode = "UNKNOWN"

        progress = st.progress(0)
        total = max(1, len(hp))

        must_cols = [
            "block", "homenumber", "fdtcode", "oltcode", "fatcode",
            "Latitude_homepass", "Longitude_homepass", "district", "subdistrict", "postalcode",
            "FAT ID", "Pole ID", "Pole Latitude", "Pole Longitude", "FAT Address",
            "Line", "Capacity", "FAT Port",
            "FDT Tray (Front)", "FDT Port", "Tube Colour", "Core Number", "street"
        ]
        for col in must_cols:
            if col not in df.columns:
                df[col] = ""

        # Prepare lists to fill dataframe in batch
        n_rows = min(len(hp), len(df))
        fatcode_list = []
        block_list = []
        homenumber_list = []
        lat_list = []
        lon_list = []
        district_list = []
        subdistrict_list = []
        postalcode_list = []
        fdtcode_list = []
        oltcode_list = []
        street_list = []
        fat_id_list = []
        fat_address_list = []
        pole_id_list = []
        pole_lat_list = []
        pole_lon_list = []

        # caches for FAT address & nearest pole by FAT name to avoid recomputation
        fatname_to_address = {}
        fatname_to_nearpole = {}

        for i in range(n_rows):
            h = hp[i]
            fc = extract_fatcode(h["path"])
            fatcode_list.append(fc)

            name_parts = h["name"].split(".")
            if len(name_parts) == 2 and name_parts[0].isalnum() and name_parts[1].isdigit():
                block_list.append(name_parts[0].strip().upper())
                homenumber_list.append(name_parts[1].strip())
            else:
                block_list.append("")
                homenumber_list.append(h["name"])

            # default lat/lon from KMZ
            orig_lat = h["lat"]
            orig_lon = h["lon"]

            # jika snap diaktifkan: coba snap ke jalan terdekat
            if snap_to_road:
                snap_res = snap_to_nearest_road(orig_lat, orig_lon)
                use_lat = snap_res.get("lat", orig_lat)
                use_lon = snap_res.get("lon", orig_lon)
                snap_street = snap_res.get("street", "")
            else:
                use_lat = orig_lat
                use_lon = orig_lon
                snap_street = ""

            lat_list.append(use_lat)
            lon_list.append(use_lon)

            # Ambil street: prioritas dari snapping jika ada, jika kosong fallback ke reverse geocode biasa
            if snap_street:
                # hapus awalan JALAN jika ada, dan strip
                street_name = snap_street.replace("JALAN ", "").strip()
            else:
                hh = reverse_here(use_lat, use_lon)
                street_name = hh["street"].replace("JALAN ", "").strip() if hh["street"] else ""

            street_list.append(street_name)

            # district/subdistrict/postalcode dari FDT base rc (sama seperti sebelumnya)
            district_list.append(rc["district"])
            subdistrict_list.append(rc["subdistrict"])
            postalcode_list.append(rc["postalcode"])
            fdtcode_list.append(fdtcode)
            oltcode_list.append(oltcode)

            # find FAT item match
            mf = next((x for x in fat if fc in x["name"]), None)
            if mf:
                fat_name = mf["name"]
                fat_id_list.append(fat_name)
                # get fat address once per fat_name
                if fat_name in fatname_to_address:
                    fat_address_list.append(fatname_to_address[fat_name])
                else:
                    fa = reverse_here(mf["lat"], mf["lon"])["street"]
                    fatname_to_address[fat_name] = fa
                    fat_address_list.append(fa)

                # nearest pole once per fat_name
                if fat_name in fatname_to_nearpole:
                    npole = fatname_to_nearpole[fat_name]
                else:
                    npole = find_nearest_pole(mf, all_poles)
                    fatname_to_nearpole[fat_name] = npole
                if npole:
                    pole_id_list.append(npole["name"])
                    pole_lat_list.append(npole["lat"])
                    pole_lon_list.append(npole["lon"])
                else:
                    pole_id_list.append("POLE_NOT_FOUND")
                    pole_lat_list.append("")
                    pole_lon_list.append("")
            else:
                fat_id_list.append("FAT_NOT_FOUND")
                fat_address_list.append("")
                pole_id_list.append("POLE_NOT_FOUND")
                pole_lat_list.append("")
                pole_lon_list.append("")

            # Update progress only periodically to reduce rerenders
            if (i % 25 == 0) or (i == n_rows - 1):
                progress.progress(int((i + 1) * 100 / total))

        
         # ✅ Tambahkan di sini, di luar loop utama
        if use_smart_fill:
            street_list = smart_fill_streets(street_list)
        
        # Assign lists to dataframe in batch (only for the rows we processed)
        idx_slice = df.index[:n_rows]
        df.loc[idx_slice, "fatcode"] = fatcode_list
        df.loc[idx_slice, "block"] = block_list
        df.loc[idx_slice, "homenumber"] = homenumber_list
        df.loc[idx_slice, "Latitude_homepass"] = lat_list
        df.loc[idx_slice, "Longitude_homepass"] = lon_list
        df.loc[idx_slice, "district"] = district_list
        df.loc[idx_slice, "subdistrict"] = subdistrict_list
        df.loc[idx_slice, "postalcode"] = postalcode_list
        df.loc[idx_slice, "fdtcode"] = fdtcode_list
        df.loc[idx_slice, "oltcode"] = oltcode_list
        df.loc[idx_slice, "street"] = street_list
        df.loc[idx_slice, "FAT ID"] = fat_id_list
        df.loc[idx_slice, "FAT Address"] = fat_address_list
        df.loc[idx_slice, "Pole ID"] = pole_id_list
        df.loc[idx_slice, "Pole Latitude"] = pole_lat_list
        df.loc[idx_slice, "Pole Longitude"] = pole_lon_list

        # ====== AUTO FILL FAT PORT ======
        # ensure FAT Port column exists and numeric
        if "FAT Port" not in df.columns:
            df["FAT Port"] = 0
        # do grouping and assign sequential port numbers per FAT ID (only rows we filled)
        for fat_id, group in df.loc[idx_slice].groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            for i_seq, idx in enumerate(group.index, start=1):
                df.at[idx, "FAT Port"] = i_seq

        # ====== AUTO FILL LINE & CAPACITY PER FAT ID ======
        for fat_id, group in df.loc[idx_slice].groupby("FAT ID", sort=False):
            if fat_id == "" or fat_id == "FAT_NOT_FOUND":
                continue
            fatcodes = group["fatcode"].dropna().unique()
            if len(fatcodes) == 0:
                continue
            fc0 = fatcodes[0]
            letter = fc0[0] if fc0 and fc0[0] in "ABCD" else ""
            try:
                nums = [int(x[1:]) for x in fatcodes if len(x) >= 2 and x[1:].isdigit()]
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
            df.at[first_idx, "Line"] = ("LINE  " + letter) if letter else ""
            df.at[first_idx, "Capacity"] = cap_val

        # ====== FINAL: ASSIGN FDT Tray, FDT Port, Tube Colour, Core Number ======
        # same logic as contoh.xlsx: global tray/port; tube/core reset when Line changes
        df['FAT Port'] = pd.to_numeric(df['FAT Port'], errors='coerce').fillna(0).astype(int)

        for col in ["FDT Tray (Front)", "FDT Port", "Tube Colour", "Core Number"]:
            if col not in df.columns:
                df[col] = np.nan
            else:
                df[col] = np.nan

        # Build FAT ID -> Line mapping (inherit previous if missing)
        fat_line = {}
        prev_line = None
        for fat_id, group in df.groupby('FAT ID', sort=False):
            nonnull_lines = group['Line'].dropna().unique()
            if len(nonnull_lines) > 0:
                line_val = nonnull_lines[0]
                prev_line = line_val
            else:
                line_val = prev_line
            fat_line[fat_id] = line_val

        current_tray = 1
        current_port = 1
        current_tube = 1
        current_core = 1
        prev_line = None

        for fat in df['FAT ID'].drop_duplicates():
            if not fat or fat == "FAT_NOT_FOUND":
                continue

            line = fat_line.get(fat, None)
            if line != prev_line:
                current_tube = 1
                current_core = 1
                prev_line = line

            grp = df[df['FAT ID'] == fat]
            if grp.empty:
                continue
            first_candidates = grp.index[grp['FAT Port'] == 1]
            if len(first_candidates) > 0:
                idx_first = first_candidates[0]
            else:
                idx_first = grp.index[0]
            second_candidates = grp.index[grp['FAT Port'] == 2]
            if len(second_candidates) > 0:
                idx_second = second_candidates[0]
            else:
                idx_second = grp.index[1] if len(grp) > 1 else None

            df.at[idx_first, "FDT Tray (Front)"] = current_tray
            df.at[idx_first, "FDT Port"] = current_port
            df.at[idx_first, "Tube Colour"] = current_tube
            df.at[idx_first, "Core Number"] = current_core

            if idx_second is not None:
                df.at[idx_second, "Tube Colour"] = current_tube
                df.at[idx_second, "Core Number"] = current_core + 1

            current_core += 2
            if current_core > 10:
                current_core -= 10
                current_tube += 1

            current_port += 1
            if current_port > 10:
                current_port = 1
                current_tray += 1

        progress.empty()
        st.success("✅ Selesai! (logika FDT Tray/Port/Tube/Core sudah disesuaikan seperti contoh.xlsx)")
        st.dataframe(df.head(60))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")

