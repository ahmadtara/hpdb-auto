import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from opencage.geocoder import OpenCageGeocode
import time

# API KEY
key = '91b8be587a2e4eb095f24802fd462089'
geocoder = OpenCageGeocode(key)

st.title("📍 KMZ ➜ HPDB (Auto-Pilot ⚡By.A.Tara-P.)")

kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

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
        kml_file = [f for f in z.namelist() if f.endswith(".kml") or f.endswith(".KML")][0]
        with z.open(kml_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}

            all_placemarks = []
            for folder in root.findall(".//kml:Folder", ns):
                all_placemarks += recurse_folder(folder, ns)

            data = {
                "FAT": [], 
                "NEW POLE 7-3": [], 
                "EXISTING POLE EMR 7-3": [],
                "EXISTING POLE EMR 7-4": [],
                "FDT": [], 
                "HP COVER": []
            }

            for p in all_placemarks:
                for key in data:
                    if key in p["path"]:
                        data[key].append(p)
                        break
            return data

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

def reverse_geocode(lat, lon):
    try:
        result = geocoder.reverse_geocode(lat, lon)
        if result and len(result):
            components = result[0]['components']
            return {
                "district": components.get("suburb", "").upper(),
                "subdistrict": components.get("city_district", "").upper(),
                "street": components.get("road", "").upper()
            }
    except:
        pass
    return {"district": "", "subdistrict": "", "street": ""}

if kmz_file and template_file:
    kmz_name = kmz_file.name.replace(".kmz", "")
    placemarks = extract_placemarks(kmz_file.read())
    df_template = pd.read_excel(template_file)

    fat_list = placemarks["FAT"]
    hp_list = placemarks["HP COVER"]
    fdt_list = placemarks["FDT"]
    fdtcode = fdt_list[0]["name"] if fdt_list else "FDT_UNKNOWN"

    # Ambil informasi lokasi dari titik FDT
    if fdt_list:
        fdt_geo = reverse_geocode(fdt_list[0]["lat"], fdt_list[0]["lon"])
        district = fdt_geo["district"]
        subdistrict = fdt_geo["subdistrict"]
    else:
        district = subdistrict = ""

    # Gabungkan semua POLE
    all_poles = placemarks["NEW POLE 7-3"] + placemarks["EXISTING POLE EMR 7-3"] + placemarks["EXISTING POLE EMR 7-4"]

    row = 0
    total = len(hp_list)
    progress = st.progress(0, text="🔄 Memproses koordinat...")

    for idx, hp in enumerate(hp_list):
        if row >= len(df_template):
            break

        fatcode = extract_fatcode_from_path(hp["path"])
        df_template.at[row, "fatcode"] = fatcode
        df_template.at[row, "homenumber"] = hp["name"]
        df_template.at[row, "Latitude_homepass"] = hp["lat"]
        df_template.at[row, "Longitude_homepass"] = hp["lon"]

        matched_fat = find_fat_by_fatcode(fatcode, fat_list)
        if matched_fat:
            df_template.at[row, "FAT ID"] = matched_fat["name"]
            df_template.at[row, "Pole Latitude"] = matched_fat["lat"]
            df_template.at[row, "Pole Longitude"] = matched_fat["lon"]
            df_template.at[row, "Pole ID"] = find_matching_pole(matched_fat, all_poles)
            fat_address = reverse_geocode(matched_fat["lat"], matched_fat["lon"])["street"]
            df_template.at[row, "FAT Address"] = fat_address
        else:
            df_template.at[row, "FAT ID"] = "FAT_NOT_FOUND"
            df_template.at[row, "Pole ID"] = "POLE_NOT_FOUND"
            df_template.at[row, "FAT Address"] = ""

        # Nama jalan (street) dari titik HP COVER
        hp_location = reverse_geocode(hp["lat"], hp["lon"])
        df_template.at[row, "street"] = hp_location["street"]

        # Kolom tetap
        df_template.at[row, "fdtcode"] = fdtcode
        df_template.at[row, "Clustername"] = kmz_name
        df_template.at[row, "Commercial_name"] = kmz_name
        df_template.at[row, "district"] = district
        df_template.at[row, "subdistrict"] = subdistrict

        row += 1
        progress.progress(int((idx + 1) / total * 100), text=f"⏳ Memproses {idx + 1}/{total}...")

    progress.empty()
    st.success("✅ Semua data berhasil diproses!")

    st.dataframe(df_template.head(10))

    output = BytesIO()
    df_template.to_excel(output, index=False)
    st.download_button("📥 Download File Hasil", output.getvalue(), file_name="HASIL_HPDB_LENGKAP.xlsx")
