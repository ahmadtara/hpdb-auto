import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import requests
import time

st.title("📍 KMZ ➜ HPDB (Auto-Fill dengan OpenCage)")

API_KEY = "91b8be587a2e4eb095f24802fd462089"

kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

# Fungsi ambil alamat dari OpenCage API
def reverse_geocode(lat, lon):
    url = f"https://api.opencagedata.com/geocode/v1/json?q={lat}+{lon}&key={API_KEY}&language=id"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            if data["results"]:
                comp = data["results"][0]["components"]
                return {
                    "street": comp.get("road", "").upper(),
                    "district": comp.get("village", comp.get("suburb", "")).upper(),
                    "subdistrict": comp.get("county", comp.get("city_district", "")).upper(),
                    "postalcode": comp.get("postcode", "")
                }
    except:
        pass
    return {"street": "", "district": "", "subdistrict": "", "postalcode": ""}

# Fungsi ekstraksi dari KMZ
def extract_placemarks(kmz_bytes):
    def recurse_folder(folder, ns, path=""):
        placemarks = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name

        for sub in folder.findall("kml:Folder", ns):
            placemarks += recurse_folder(sub, ns, new_path)

        for pm in folder.findall("kml:Placemark", ns):
            name_el = pm.find("kml:name", ns)
            coord_el = pm.find(".//kml:coordinates", ns)
            if name_el is not None and coord_el is not None:
                coords = coord_el.text.strip().split(",")
                if len(coords) >= 2:
                    placemarks.append({
                        "name": name_el.text.strip(),
                        "lat": float(coords[1].strip()),
                        "lon": float(coords[0].strip()),
                        "path": new_path
                    })
        return placemarks

    with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
        kml_file = [f for f in z.namelist() if f.endswith(".kml")][0]
        with z.open(kml_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}

            all_placemarks = []
            for folder in root.findall(".//kml:Folder", ns):
                all_placemarks += recurse_folder(folder, ns)

            data = {"FAT": [], "FDT": [], "HP COVER": [], "NEW POLE 7-3": [], "EXISTING POLE EMR 7-3": [], "EXISTING POLE EMR 7-4": []}

            for p in all_placemarks:
                for key in data:
                    if key in p["path"]:
                        data[key].append(p)
                        break
            return data

# Fungsi bantu FAT & POLE
def extract_fatcode_from_path(path):
    parts = path.split("/")
    for part in parts:
        if len(part) == 3 and part[0] in "ABCD" and part[1:].isdigit():
            return part
    return "UNKNOWN"

def find_fat_by_fatcode(fatcode, fat_list):
    for fat in fat_list:
        if fatcode in fat["name"]:
            return fat
    return None

def find_matching_pole(fat, all_poles, tol=0.0001):
    for pole in all_poles:
        if abs(fat["lat"] - pole["lat"]) < tol and abs(fat["lon"] - pole["lon"]) < tol:
            return pole["name"]
    return "POLE_NOT_FOUND"

# Proses utama
if kmz_file and template_file:
    kmz_name = kmz_file.name.replace(".kmz", "")
    placemarks = extract_placemarks(kmz_file.read())
    df_template = pd.read_excel(template_file)

    fat_list = placemarks["FAT"]
    hp_list = placemarks["HP COVER"]
    fdt_list = placemarks["FDT"]
    fdtcode = fdt_list[0]["name"] if fdt_list else "FDT_UNKNOWN"

    # Ambil alamat FDT untuk district, subdistrict, postalcode
    if fdt_list:
        fdt_coord = fdt_list[0]
        fdt_address = reverse_geocode(fdt_coord["lat"], fdt_coord["lon"])
        district_all = fdt_address["district"]
        subdistrict_all = fdt_address["subdistrict"]
        postalcode_all = fdt_address["postalcode"]
    else:
        district_all = ""
        subdistrict_all = ""
        postalcode_all = ""

    all_poles = placemarks["NEW POLE 7-3"] + placemarks["EXISTING POLE EMR 7-3"] + placemarks["EXISTING POLE EMR 7-4"]

    total_hp = len(hp_list)
    st.info(f"Jumlah HP COVER: {total_hp}. Proses ini memerlukan waktu sekitar {total_hp * 1.2:.1f} detik...")

    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, hp in enumerate(hp_list):
        if idx >= len(df_template):
            break

        fatcode = extract_fatcode_from_path(hp["path"])
        df_template.at[idx, "fatcode"] = fatcode
        df_template.at[idx, "homenumber"] = hp["name"]
        df_template.at[idx, "Latitude_homepass"] = hp["lat"]
        df_template.at[idx, "Longitude_homepass"] = hp["lon"]

        # Ambil nama jalan tiap HP COVER
        hp_address = reverse_geocode(hp["lat"], hp["lon"])
        df_template.at[idx, "street"] = hp_address["street"]

        # Kolom district, subdistrict, postalcode
        df_template.at[idx, "district"] = district_all
        df_template.at[idx, "subdistrict"] = subdistrict_all
        df_template.at[idx, "postalcode"] = postalcode_all

        matched_fat = find_fat_by_fatcode(fatcode, fat_list)
        if matched_fat:
            df_template.at[idx, "FAT ID"] = matched_fat["name"]
            df_template.at[idx, "Pole Latitude"] = matched_fat["lat"]
            df_template.at[idx, "Pole Longitude"] = matched_fat["lon"]
            df_template.at[idx, "Pole ID"] = find_matching_pole(matched_fat, all_poles)
        else:
            df_template.at[idx, "FAT ID"] = "FAT_NOT_FOUND"
            df_template.at[idx, "Pole ID"] = "POLE_NOT_FOUND"

        df_template.at[idx, "fdtcode"] = fdtcode
        df_template.at[idx, "Clustername"] = kmz_name
        df_template.at[idx, "Commercial_name"] = kmz_name

        # Update progress
        progress = int(((idx + 1) / total_hp) * 100)
        progress_bar.progress(progress)
        status_text.text(f"Memproses HP COVER {idx+1} dari {total_hp}...")
        time.sleep(1)  # Simulasi waktu tunggu agar terlihat smooth

    st.success("✅ Proses selesai!")
    st.dataframe(df_template.head(10))

    # Download file hasil
    output = BytesIO()
    df_template.to_excel(output, index=False)
    st.download_button("📥 Download File Hasil", output.getvalue(), file_name="HASIL_HPDB_LENGKAP.xlsx")
