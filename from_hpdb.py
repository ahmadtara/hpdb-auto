import streamlit as st
import pandas as pd
import zipfile
from lxml import etree   # pakai lxml bukan xml.etree
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
    # Helper untuk FDT mapping
    # ------------------------------
    def get_line_from_fat_id(fat_id: str) -> str:
        m = re.search(r'([A-Z])0*(\d+)$', fat_id)
        return f"LINE {m.group(1)}" if m else "LINE ?"

    def get_capacity_from_fat_id(fat_id: str) -> str:
        m = re.search(r'[A-Z]0*(\d+)$', fat_id)
        if not m:
            return "UNKNOWN"
        num = int(m.group(1))
        if 1 <= num <= 10:
            return "24C/2T"
        elif 11 <= num <= 15:
            return "36C/3T"
        else:
            return "48C/4T"

    def expand_fat_id_to_rows(fat_id: str):
        """Return list of dict rows for one FAT ID following rules:
           trays = based on capacity (2/3/4), each tray 10 ports, each port 2 cores.
        """
        capacity = get_capacity_from_fat_id(fat_id)
        # tray_count: parse like '2T' -> 2, '3T' -> 3, '4T' -> 4
        try:
            tray_count = int(capacity.split("/")[1][0])
        except Exception:
            tray_count = 4  # fallback
        line = get_line_from_fat_id(fat_id)
        tray_to_colour = {1: "Blue", 2: "Yellow", 3: "Green", 4: "Brown"}
        port_per_tray = 10
        rows = []
        for tray in range(1, tray_count + 1):
            for port in range(1, port_per_tray + 1):
                for core in [1, 2]:
                    rows.append({
                        "FDT Tray (Front)": tray,
                        "FDT Port": port,
                        "Line": line,
                        "Capacity": capacity,
                        "Tube Colour": tray_to_colour.get(tray, str(tray)),
                        "Core Number": core,
                        "FAT ID": fat_id,
                        "FAT PORT": tray_count * port_per_tray  # total ports per FAT
                    })
        return rows

    # ------------------------------
    # Main flow
    # ------------------------------
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
        total = len(hp) if hp else 0

        # ensure columns exist
        for col in ["block", "homenumber", "fdtcode", "oltcode"]:
            if col not in df.columns:
                df[col] = ""

        # main loop for HP cover -> fill df template rows
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
                fataddr = reverse_here(mf["lat"], mf["lon"])["street"]
                df.at[i, "FAT Address"] = fataddr
            else:
                df.at[i, "FAT ID"] = "FAT_NOT_FOUND"
                df.at[i, "Pole ID"] = "POLE_NOT_FOUND"
                df.at[i, "FAT Address"] = ""

            # update progress
            if total:
                progress.progress(int((i + 1) * 100 / total))

        progress.empty()
        st.success("✅ Selesai pengisian HPDB!")

        # === Generate FDT_Port_Mapping sheet ===
        # Build ordered list of FAT IDs from placemarks['FAT'] (preserve order in KMZ)
        fat_ids_ordered = [x["name"] for x in fat]
        # fallback: if none in placemarks but some in df, use unique df
        if not fat_ids_ordered:
            fat_ids_ordered = [v for v in df["FAT ID"].unique() if pd.notna(v)]

        fdt_rows = []
        for fat_id in fat_ids_ordered:
            if not fat_id or fat_id == "FAT_NOT_FOUND":
                continue
            fdt_rows.extend(expand_fat_id_to_rows(fat_id))

        if fdt_rows:
            df_fdt_map = pd.DataFrame(fdt_rows)
            # sort rapi
            df_fdt_map = df_fdt_map.sort_values(by=["FAT ID", "FDT Tray (Front)", "FDT Port", "Core Number"])
            st.markdown("**Preview: FDT_Port_Mapping**")
            st.dataframe(df_fdt_map.head(200))
        else:
            df_fdt_map = pd.DataFrame(columns=["FDT Tray (Front)", "FDT Port", "Line", "Capacity", "Tube Colour", "Core Number", "FAT ID", "FAT PORT"])
            st.info("Tidak ada FAT ID yang ditemukan untuk generasi FDT mapping.")

        # === Create Excel with two sheets and provide download button ===
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            # HPDB sheet (original template with filled columns)
            df.to_excel(writer, sheet_name="HPDB", index=False)
            # FDT mapping sheet
            df_fdt_map.to_excel(writer, sheet_name="FDT_Port_Mapping", index=False)
            writer.save()
        buf.seek(0)

        st.download_button("📥 Download Hasil (HPDB + FDT_Port_Mapping)", buf.getvalue(), file_name="hasil_hpdb_with_fdt_mapping.xlsx")
        st.info("File berisi 2 sheet: 'HPDB' dan 'FDT_Port_Mapping' — buka di Excel untuk melihat warna/tray jika perlu formatting tambahan.")
