import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
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

    # ===== EXTRACT PLACEMARKS AMAN =====
    def extract_placemarks(kmz_bytes):
        def recurse_folder(folder, ns, path=""):
            placemarks = []
            folder_name_el = folder.find("kml:name", ns)
            folder_name = folder_name_el.text.strip() if folder_name_el is not None else "UNKNOWN"
            new_path = f"{path}/{folder_name}" if path else folder_name

            # Ambil semua placemark
            for pm in folder.findall("kml:Placemark", ns):
                name_el = pm.find("kml:name", ns)
                name = name_el.text.strip() if name_el is not None else "Unnamed"

                coords = pm.findall(".//kml:coordinates", ns)
                for coord in coords:
                    if coord is None or coord.text is None:
                        continue
                    text = coord.text.strip()
                    if not text:
                        continue
                    parts = text.split(",")
                    if len(parts) < 2:
                        continue
                    try:
                        lon, lat = float(parts[0]), float(parts[1])
                        placemarks.append({
                            "name": name,
                            "lat": lat,
                            "lon": lon,
                            "path": new_path
                        })
                    except ValueError:
                        continue  # skip kalau parsing gagal

            # Rekursif ke folder dalam folder
            for sub in folder.findall("kml:Folder", ns):
                placemarks += recurse_folder(sub, ns, new_path)

            return placemarks

        with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
            kml_files = [f for f in z.namelist() if f.lower().endswith(".kml")]
            if not kml_files:
                raise FileNotFoundError("Tidak ada file KML di dalam KMZ")
            f = kml_files[0]
            root = ET.parse(z.open(f)).getroot()
            ns = {"kml": "http://www.opengis.net/kml/2.2"}

            all_pm = []
            for folder in root.findall(".//kml:Folder", ns):
                all_pm += recurse_folder(folder, ns)

            # Grouping by category
            data = {k: [] for k in ["FAT", "NEW POLE 7-3", "EXISTING POLE EMR 7-3", "EXISTING POLE EMR 7-4", "FDT", "HP COVER"]}
            for p in all_pm:
                path = p.get("path", "")
                for k in data:
                    if k in path:
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

    # ===== MAIN PROCESS =====
    if kmz_file and template_file:
        kmz_bytes = kmz_file.read()
        placemarks = extract_placemarks(kmz_bytes)
        df = pd.read_excel(template_file)

        fat = placemarks["FAT"]
        hp = placemarks["HP COVER"]
        fdt = placemarks["FDT"]
        all_poles = placemarks["NEW POLE 7-3"] + placemarks["EXISTING POLE EMR 7-3"] + placemarks["EXISTING POLE EMR 7-4"]

        rc = reverse_here(fdt[0]["lat"], fdt[0]["lon"]) if fdt else {"district": "", "subdistrict": "", "postalcode": "", "street": ""}
        fdtcode = fdt[0]["name"].strip().upper() if fdt else "UNKNOWN"
        oltcode = "UNKNOWN"

        # Ambil OLT dari FDT description
        if fdt:
            with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
                kml_file = [f for f in z.namelist() if f.lower().endswith(".kml")][0]
                tree = ET.parse(z.open(kml_file))
                root = tree.getroot()
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

        for col in ["block", "homenumber", "fdtcode", "oltcode"]:
            if col not in df.columns:
                df[col] = ""

        # ===== Populate dataframe =====
        for i, h in enumerate(hp):
            if i >= len(df): break
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
                pol = next((p["name"] for p in all_poles if abs(p["lat"] - mf["lat"]) < 1e-4 and abs(p["lon"] - mf["lon"]) < 1e-4), "POLE_NOT_FOUND")
                df.at[i, "Pole ID"] = pol
                fataddr = reverse_here(mf["lat"], mf["lon"])["street"]
                df.at[i, "FAT Address"] = fataddr
            else:
                df.at[i, "FAT ID"] = "FAT_NOT_FOUND"
                df.at[i, "Pole ID"] = "POLE_NOT_FOUND"
                df.at[i, "FAT Address"] = ""

            progress.progress(int((i + 1) * 100 / total))

        progress.empty()
        st.success("✅ Selesai!")
        st.dataframe(df.head(10))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
