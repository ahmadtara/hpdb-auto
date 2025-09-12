import streamlit as st
import pandas as pd
import zipfile
from lxml import etree   # pakai lxml bukan xml.etree
from io import BytesIO
import requests
import re
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
    # Helper untuk ekstraksi koordinat dan rekursi folder
    # ------------------------------
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

    def extract_placemarks(kmz_bytes):
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

    # ------------------------------
    # Normalisasi dan ekstrak kode FAT
    # ------------------------------
    def normalize_code(token):
        # token like A1, a01, A001 -> A01 (two digit)
        m = re.match(r'([ABCD])0*([0-9]{1,3})$', token.upper().strip())
        if not m:
            return None
        letter = m.group(1)
        num = int(m.group(2))
        return f"{letter}{num:02d}"

    def extract_fatcode_from_text(txt):
        if txt is None:
            return None
        txt = txt.upper()
        m = re.search(r'([ABCD])0*([0-9]{1,3})', txt)
        if m:
            return f"{m.group(1)}{int(m.group(2)):02d}"
        return None

    def extract_fatcode(path):
        # cek setiap segmen path dulu, lalu fallback ke full path
        for part in path.split("/"):
            c = extract_fatcode_from_text(part)
            if c:
                return c
        return extract_fatcode_from_text(path) or "UNKNOWN"

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

        # Cari OLT dari description FDT (sama seperti sebelumnya)
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

        # ----------------------------
        # Build mapping FAT -> daftar kode port (mis. A01, A02, ...)
        # ----------------------------
        code_to_fat_name = {}
        fat_name_to_codes = defaultdict(set)

        for f_item in fat:
            # cek name dan path teks untuk kode
            s = (f_item.get("name","") + " " + f_item.get("path","")).upper()
            codes_found = re.findall(r'([ABCD]\d{1,3})', s)
            for raw in codes_found:
                norm = normalize_code(raw)
                if norm:
                    code_to_fat_name[norm] = f_item["name"]
                    fat_name_to_codes[f_item["name"]].add(norm)

        # ----------------------------
        # PASS 1: isi kolom dasar & kumpulkan metadata per homepass
        # ----------------------------
        hp_meta = []  # list of dicts: {idx, fc, fat_name}
        total = len(hp)
        progress = st.progress(0)

        for i, h in enumerate(hp):
            if i >= len(df):
                break

            # fatcode dari path
            fc = extract_fatcode(h["path"])
            df.at[i, "fatcode"] = fc

            # block & homenumber
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

            # Tentukan FAT (mf) utk kolom FAT ID, Pole, dll (cara fallback sebelumnya)
            mf = None
            if fc and fc != "UNKNOWN":
                mapped = code_to_fat_name.get(fc)
                if mapped:
                    mf = next((x for x in fat if x["name"] == mapped), None)
            if mf is None:
                mf = next((x for x in fat if fc in x.get("name","").upper()), None)

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

            hp_meta.append({"idx": i, "fc": fc, "fat_name": df.at[i, "FAT ID"]})
            progress.progress(int((i + 1) * 100 / max(1, total)))

        # ----------------------------
        # PASS 2: proses blok per FAT ID (kontigu) dan isi FAT Port dkk berdasarkan jumlah port FAT itu
        # ----------------------------
        def capacity_from_num(n):
            if 1 <= n <= 10:
                return "24C/2T"
            elif 11 <= n <= 15:
                return "36C/3T"
            elif 16 <= n <= 20:
                return "48C/4T"
            return "UNKNOWN"

        i = 0
        while i < len(hp_meta):
            curr_fat = hp_meta[i]["fat_name"]
            block_indices = []
            j = i
            # kumpulkan blok kontigu dengan FAT ID yang sama
            while j < len(hp_meta) and hp_meta[j]["fat_name"] == curr_fat:
                block_indices.append(hp_meta[j]["idx"])
                j += 1

            if curr_fat == "FAT_NOT_FOUND":
                # lewati pengisian FAT Port (biarkan kosong)
                i = j
                continue

            # ambil daftar kode port untuk FAT ini (urutkan asc)
            port_codes = sorted(list(fat_name_to_codes.get(curr_fat, [])),
                                key=lambda c: int(c[1:]) if len(c) > 1 else 0)
            number_of_ports = len(port_codes)

            # fallback: jika tidak ada mapping, gunakan fc unik di blok (urutan)
            if number_of_ports == 0:
                block_fcs = [hp_meta[k]["fc"] for k in range(i, j)]
                block_codes = [c for c in (list(dict.fromkeys(block_fcs))) if c != "UNKNOWN"]
                block_codes = sorted(block_codes, key=lambda c: int(c[1:]) if len(c) > 1 else 0)
                port_codes = block_codes
                number_of_ports = len(port_codes)

            # Isi FAT Port hanya sampai jumlah port (batas). Untuk setiap port ke-k dalam FAT:
            # - tray_val = floor(k/10)+1  (tray tetap 1 sampai 10 cores, lalu 2 dst)
            # - core_seq = (k % 10) + 1  (1..10)
            # - core pair = (core_seq-1)*2+1 , (core_seq-1)*2+2  -> "1,2" ... "19,20"
            for k, row_idx in enumerate(block_indices):
                if k < number_of_ports:
                    assigned_code = port_codes[k]  # ex 'A01'
                    try:
                        assigned_port_num = int(assigned_code[1:])
                    except:
                        assigned_port_num = k + 1
                    # isi FAT Port numeric
                    df.at[row_idx, "FAT Port"] = assigned_port_num

                    # --- NEW LOGIC sesuai permintaan: tray/tube berubah setiap 10 core ---
                    tray_val = (k // 10) + 1  # tray 1 untuk k=0..9 ; tray 2 untuk k=10..19
                    core_seq = (k % 10) + 1   # 1..10 then reset when tray increments

                    # FDT Tray (Front) dan FDT Port = tray_val (tray block of 10)
                    df.at[row_idx, "FDT Tray (Front)"] = tray_val
                    df.at[row_idx, "FDT Port"] = tray_val

                    # Line berdasarkan huruf assigned_code[0]
                    letter = assigned_code[0] if len(assigned_code) > 0 else ""
                    df.at[row_idx, "Line"] = f"Line {letter}" if letter else "UNKNOWN"

                    # Capacity berdasarkan nomor port
                    df.at[row_idx, "Capacity"] = capacity_from_num(assigned_port_num)

                    # Tube Colour: berdasarkan tray_val -> warna siklus per tray
                    df.at[row_idx, "Tube Colour"] = tube_colours[(tray_val - 1) % len(tube_colours)]

                    # Core Number: pasangan berdasarkan core_seq (1..10)
                    core_start = (core_seq - 1) * 2 + 1
                    df.at[row_idx, "Core Number"] = f"{core_start},{core_start+1}"
                else:
                    # melebihi jumlah port FAT -> biarkan kosong (batas isi sesuai permintaan)
                    df.at[row_idx, "FAT Port"] = ""
                    df.at[row_idx, "FDT Tray (Front)"] = ""
                    df.at[row_idx, "FDT Port"] = ""
                    df.at[row_idx, "Line"] = ""
                    df.at[row_idx, "Capacity"] = ""
                    df.at[row_idx, "Tube Colour"] = ""
                    df.at[row_idx, "Core Number"] = ""
            # lanjut ke blok berikutnya
            i = j

        progress.empty()
        st.success("✅ Selesai! (logika tray/tube/core: 10-core per tray, reset core tiap tray telah diterapkan)")
        st.dataframe(df.head(50))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")


# Jika mau jalanin langsung (contoh):
if __name__ == "__main__":
    HERE_API_KEY = st.secrets.get("HERE_API_KEY", "") if "st" in globals() else ""
    run_hpdb(HERE_API_KEY)
