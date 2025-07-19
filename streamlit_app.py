import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

st.title("📍 KMZ ➜ HPDB (FAT, Pole, FDT, HP COVER)")

kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

def extract_placemarks(kmz_bytes):
    with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
        kml_file = [f for f in z.namelist() if f.endswith(".kml")][0]
        with z.open(kml_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            data = {"FAT": [], "NEW POLE 7-3": [], "FDT": [], "HP COVER": []}

            for folder in root.findall(".//kml:Folder", ns):
                folder_name_el = folder.find("kml:name", ns)
                if folder_name_el is None:
                    continue
                folder_name = folder_name_el.text.upper()

                for fkey in data:
                    if fkey in folder_name:
                        for pm in folder.findall("kml:Placemark", ns):
                            name_el = pm.find("kml:name", ns)
                            coord_el = pm.find(".//kml:coordinates", ns)
                            if name_el is not None and coord_el is not None:
                                coords = coord_el.text.strip().split(",")
                                if len(coords) >= 2:
                                    data[fkey].append({
                                        "name": name_el.text.strip(),
                                        "lat": float(coords[1].strip()),
                                        "lon": float(coords[0].strip())
                                    })
            return data

def find_nearest_pole(fat, poles, tol=0.0001):
    for pole in poles:
        if abs(fat["lat"] - pole["lat"]) < tol and abs(fat["lon"] - pole["lon"]) < tol:
            return pole["name"]
    return "POLE_NOT_FOUND"

if kmz_file and template_file:
    kmz_name = kmz_file.name.replace(".kmz", "")
    placemarks = extract_placemarks(kmz_file.read())
    df_template = pd.read_excel(template_file)

    for i in range(min(len(placemarks["FAT"]), len(df_template))):
        fat = placemarks["FAT"][i]
        df_template.at[i, "FAT ID"] = fat["name"]
        df_template.at[i, "Pole Latitude"] = fat["lat"]
        df_template.at[i, "Pole Longitude"] = fat["lon"]
        df_template.at[i, "Pole ID"] = find_nearest_pole(fat, placemarks["NEW POLE 7-3"])
        df_template.at[i, "fdtcode"] = placemarks["FDT"][i]["name"] if i < len(placemarks["FDT"]) else f"FDT_{i+1}"
        df_template.at[i, "Clustername"] = kmz_name
        df_template.at[i, "Commercial_name"] = kmz_name

    for i in range(min(len(placemarks["HP COVER"]), len(df_template))):
        hp = placemarks["HP COVER"][i]
        df_template.at[i, "homenumber"] = hp["name"]
        df_template.at[i, "Latitude_homepass"] = hp["lat"]
        df_template.at[i, "Longitude_homepass"] = hp["lon"]

    st.success("✅ Data berhasil dimasukkan ke dalam template.")
    st.dataframe(df_template.head(10))

    output = BytesIO()
    df_template.to_excel(output, index=False)
    st.download_button("📥 Download File Hasil", output.getvalue(), file_name="HASIL_HPDB_LENGKAP.xlsx")
