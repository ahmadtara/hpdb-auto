import streamlit as st
import pandas as pd
import zipfile
from lxml import etree   # pakai lxml bukan xml.etree
from io import BytesIO
import requests

def run_hpdb(HERE_API_KEY):

    st.title("📍 KMZ ➜ HPDB (Auto-Pilot⚡)")
    st.markdown("""...""", unsafe_allow_html=True)  # tetap pakai markdown kamu

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

        # pastikan kolom ada (tambahkan Line dan Capacity)
        for col in ["block", "homenumber", "fdtcode", "oltcode", "Line", "Capacity"]:
            if col not in df.columns:
                df[col] = ""

        for i, h in enumerate(hp):
            if i >= len(df): 
                break
            fc = extract_fatcode(h["path"])
            df.at[i, "fatcode"] = fc

            # -- isi block / homenumber seperti sebelumnya --
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

            # ------- LOGIKA BARU: isi Line & Capacity berdasarkan fatcode -------
            # eks: fc = "A01" atau "B12"
            line_val = ""
            capacity_val = ""
            try:
                if isinstance(fc, str) and len(fc) >= 2 and fc[0] in "ABCD" and fc[1:].isdigit():
                    letter = fc[0]
                    num = int(fc[1:])  # '01' -> 1
                    if 1 <= num <= 20:
                        line_val = letter  # Line = A/B/C/D
                        # Capacity mapping
                        if 1 <= num <= 10:
                            capacity_val = "24C/2T"
                        elif 11 <= num <= 15:
                            capacity_val = "36C/3T"
                        elif 16 <= num <= 20:
                            capacity_val = "48C/4T"
                        else:
                            capacity_val = ""
                    else:
                        # nomor di luar 1-20 => kosong/UNKNOWN sesuai kebutuhan
                        line_val = ""
                        capacity_val = ""
                else:
                    line_val = ""
                    capacity_val = ""
            except Exception:
                line_val = ""
                capacity_val = ""

            df.at[i, "Line"] = line_val
            df.at[i, "Capacity"] = capacity_val
            # --------------------------------------------------------------------

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

            progress.progress(int((i + 1) * 100 / total))

        progress.empty()
        st.success("✅ Selesai!")
        st.dataframe(df.head(10))
        buf = BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
