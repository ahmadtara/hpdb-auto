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
    # helper
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
            part = part.strip().upper()
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

    def find_fdt_port_col(df):
        # cari kolom yang mengandung kata FDT dan PORT (case-insensitive)
        for c in df.columns:
            u = c.upper().replace(" ", "")
            if "FDT" in u and "PORT" in u:
                return c
        # fallback nama umum
        for alt in ["FDT Port", "FDT_PORT", "FDTPORT", "FDT\nPort", "FDT Port "]:
            if alt in df.columns:
                return alt
        return None

    # ------------------------------
    # main
    # ------------------------------
    if kmz_file and template_file:
        kmz_bytes = kmz_file.read()
        placemarks = extract_placemarks(kmz_bytes)
        df = pd.read_excel(template_file)

        fat = placemarks.get("FAT", [])
        hp = placemarks.get("HP COVER", [])
        fdt = placemarks.get("FDT", [])
        all_poles = (
            placemarks.get("NEW POLE 7-3", []) +
            placemarks.get("EXISTING POLE EMR 7-3", []) +
            placemarks.get("EXISTING POLE EMR 7-4", [])
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
                        if desc_el is not None and desc_el.text:
                            oltcode = desc_el.text.strip().upper()
                        break

        progress = st.progress(0)
        total = max(1, len(hp))

        # pastikan kolom wajib ada (tetapi jangan overwrite kolom template FDT/Tube/Core)
        must_cols = ["block", "homenumber", "fdtcode", "oltcode", "fatcode",
                     "Latitude_homepass", "Longitude_homepass", "district", "subdistrict", "postalcode",
                     "FAT ID", "Pole ID", "Pole Latitude", "Pole Longitude", "FAT Address",
                     "Line", "Capacity"]
        for col in must_cols:
            if col not in df.columns:
                df[col] = ""

        # isi data per HP (tetap isi fatcode, koordinat, alamat, FAT ID, Pole ID, dll.)
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

            progress.progress(int((i + 1) * 100 / total))

        # ---------------------------
        # SEKARANG: hitung Line & Capacity per LETTER GROUP (A/B/C/D)
        # - Capacity ditentukan dari nilai terbesar suffix dalam grup
        # - Hanya tulis ke 1 baris per grup:
        #     pilih baris pertama di grup yang punya FDT Port terisi,
        #     kalau tidak ada, pilih baris pertama munculnya grup
        # ---------------------------
        # kosongkan dulu kolom Line & Capacity agar pasti hanya 1 cell per group terisi
        df["Line"] = ""
        df["Capacity"] = ""

        fdt_port_col = find_fdt_port_col(df)  # bisa None kalau tidak ada di template

        # kumpulkan per letter
        letter_groups = {}
        for idx, fc in df["fatcode"].items():
            if not isinstance(fc, str):
                continue
            fc = fc.strip().upper()
            if len(fc) >= 2 and fc[0] in "ABCD" and fc[1:].isdigit():
                letter = fc[0]
                try:
                    num = int(fc[1:])
                except:
                    continue
                letter_groups.setdefault(letter, []).append((idx, num))

        # untuk tiap letter, tentukan max num -> capacity, dan baris target -> isi Line & Capacity
        for letter, pairs in letter_groups.items():
            if not pairs:
                continue
            # urutkan berdasarkan index (posisi di dataframe) agar deterministik
            pairs_sorted = sorted(pairs, key=lambda x: x[0])
            idxs = [p[0] for p in pairs_sorted]
            nums = [p[1] for p in pairs_sorted]
            max_num = max(nums)

            # capacity rules berdasarkan max_num
            if 1 <= max_num <= 10:
                cap_val = "24C/2T"
            elif 11 <= max_num <= 15:
                cap_val = "36C/3T"
            elif 16 <= max_num <= 20:
                cap_val = "48C/4T"
            else:
                cap_val = ""

            # pilih baris target:
            target_idx = None
            if fdt_port_col and fdt_port_col in df.columns:
                # cari baris di idxs yang memiliki nilai non-empty pada kolom FDT Port
                for ix in idxs:
                    try:
                        val = df.at[ix, fdt_port_col]
                    except KeyError:
                        val = ""
                    if pd.notna(val) and str(val).strip() != "":
                        target_idx = ix
                        break
            # fallback: pilih baris pertama munculnya group
            if target_idx is None:
                target_idx = idxs[0]

            # isi hanya pada target_idx
            df.at[target_idx, "Line"] = letter
            df.at[target_idx, "Capacity"] = cap_val
            # sisanya tetap kosong (biar mirip merged cell / satu entry per group)

        progress.empty()
        st.success("✅ Selesai!")
        st.dataframe(df.head(30))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
