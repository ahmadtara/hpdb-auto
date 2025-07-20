import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from io import BytesIO
import re
import requests
import time

def extract_fatcode_from_path(path):
    match = re.search(r"FAT\\\\(.+?)$", path)
    return match.group(1) if match else "FAT_UNKNOWN"

def extract_placemarks(kmz_data):
    with zipfile.ZipFile(BytesIO(kmz_data)) as z:
        kml_file = [f for f in z.namelist() if f.endswith('.kml')][0]
        with z.open(kml_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            folders = root.findall('.//kml:Folder', ns)
            placemarks = {
                "FAT": [], "NEW POLE 7-3": [], "HP COVER": [],
                "FDT": [], "EXISTING POLE EMR 7-3": [], "EXISTING POLE EMR 7-4": []
            }
            for folder in folders:
                name = folder.find('kml:name', ns).text if folder.find('kml:name', ns) is not None else ""
                if name in placemarks:
                    for pm in folder.findall('.//kml:Placemark', ns):
                        pname = pm.find('kml:name', ns).text if pm.find('kml:name', ns) is not None else ""
                        coords = pm.find('.//kml:coordinates', ns)
                        if coords is not None:
                            lon, lat, *_ = map(float, coords.text.strip().split(','))
                            placemarks[name].append({"name": pname, "lat": lat, "lon": lon, "path": name + "\\" + pname})
    return placemarks

def find_fat_by_fatcode(fatcode, fat_list):
    for fat in fat_list:
        if fatcode.lower() in fat['name'].lower():
            return fat
    return None

def find_matching_pole(fat, poles):
    for pole in poles:
        if abs(fat['lat'] - pole['lat']) < 0.0001 and abs(fat['lon'] - pole['lon']) < 0.0001:
            return pole['name']
    return "POLE_NOT_FOUND"

def get_location_info(lat, lon, api_key):
    url = f"https://api.opencagedata.com/geocode/v1/json?q={lat},{lon}&key={api_key}&language=id"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data['results']:
            comp = data['results'][0]['components']
            district = comp.get('district') or comp.get('county') or comp.get('state_district') or "UNKNOWN"
            village = comp.get('village') or comp.get('suburb') or comp.get('neighbourhood') or "UNKNOWN"
            return district.upper(), village.upper()
    return "UNKNOWN", "UNKNOWN"

def get_street_name(lat, lon, api_key):
    url = f"https://api.opencagedata.com/geocode/v1/json?q={lat},{lon}&key={api_key}&language=id"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data['results']:
            comp = data['results'][0]['components']
            street = comp.get('road') or comp.get('street') or "UNKNOWN"
            return street.upper()
    return "UNKNOWN"

st.title("📍 KMZ ➜ HPDB (Auto-Pilot ⚡By.A.Tara-P.)")
kmz_file = st.file_uploader("Unggah file KMZ", type=[".kmz"])
template_file = st.file_uploader("Unggah template Excel", type=[".xlsx"])

if kmz_file and template_file:
    kmz_name = kmz_file.name.replace(".kmz", "")
    placemarks = extract_placemarks(kmz_file.read())
    df_template = pd.read_excel(template_file)

    fat_list = placemarks["FAT"]
    hp_list = placemarks["HP COVER"]
    fdtcode = placemarks["FDT"][0]["name"] if placemarks["FDT"] else "FDT_UNKNOWN"

    all_poles = placemarks["NEW POLE 7-3"] + placemarks["EXISTING POLE EMR 7-3"] + placemarks["EXISTING POLE EMR 7-4"]

    fdt_coords = (placemarks["FDT"][0]["lat"], placemarks["FDT"][0]["lon"])
    st.info("🔄 Mengambil data lokasi District dan Subdistrict dari koordinat FDT...")
    district, subdistrict = get_location_info(fdt_coords[0], fdt_coords[1], api_key="91b8be587a2e4eb095f24802fd462089")
    time.sleep(1)

    row = 0
    for hp in hp_list:
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
        else:
            df_template.at[row, "FAT ID"] = "FAT_NOT_FOUND"
            df_template.at[row, "Pole ID"] = "POLE_NOT_FOUND"

        df_template.at[row, "fdtcode"] = fdtcode
        df_template.at[row, "Clustername"] = kmz_name
        df_template.at[row, "Commercial_name"] = kmz_name
        df_template.at[row, "district"] = district
        df_template.at[row, "subdistrict"] = subdistrict

        st.info(f"📍 Mendeteksi nama jalan dari koordinat Homepass {hp['lat']}, {hp['lon']}")
        street = get_street_name(hp["lat"], hp["lon"], api_key="91b8be587a2e4eb095f24802fd462089")
        df_template.at[row, "street"] = street
        time.sleep(1.5)

        row += 1

    st.success("✅ Data berhasil diproses lengkap.")
    st.dataframe(df_template.head(10))

    output = BytesIO()
    df_template.to_excel(output, index=False)
    st.download_button("📥 Download File Hasil", output.getvalue(), file_name="HASIL_HPDB_LENGKAP.xlsx")
