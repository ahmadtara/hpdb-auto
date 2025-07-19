import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

st.title("📍 KMZ ➜ HPDB (Auto-Fill)")

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

            # Kelompokkan berdasarkan folder
            data = {"FAT": [], "NEW POLE 7-3": [], "FDT": [], "HP COVER": []}
            for p in all_placemarks:
                for key in data:
                    if key in p["path"]:
                        data[key].append(p)
                        break
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

    fat_list = placemarks["FAT"]
    pole_list = placemarks["NEW POLE 7-3"]
    fdtcode = placemarks["FDT"][0]["name"] if placemarks["FDT"] else "FDT_UNKNOWN"
    hp_list = placemarks["HP COVER"]

    for i in range(len(df_template)):
        if i < len(fat_list):
            fat = fat_list[i]
            df_template.at[i, "FAT ID"] = fat["name"]
            df_template.at[i, "Pole Latitude"] = fat["lat"]
            df_template.at[i, "Pole Longitude"] = fat["lon"]
            df_template.at[i, "Pole ID"] = find_nearest_pole(fat, pole_list)

        if i < len(hp_list):
            df_template.at[i, "homenumber"] = hp_list[i]["name"]
            df_template.at[i, "Latitude_homepass"] = hp_list[i]["lat"]
            df_template.at[i, "Longitude_homepass"] = hp_list[i]["lon"]

        df_template.at[i, "fdtcode"] = fdtcode
        df_template.at[i, "Clustername"] = kmz_name
        df_template.at[i, "Commercial_name"] = kmz_name

    st.success("✅ Data berhasil dimasukkan ke dalam template.")
    st.dataframe(df_template.head(10))

    output = BytesIO()
    df_template.to_excel(output, index=False)
    st.download_button("📥 Download File Hasil", output.getvalue(), file_name="HASIL_HPDB_LENGKAP.xlsx")
