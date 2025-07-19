import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

st.title("📍 Konversi KMZ ➜ TEMPLATE HPDB")

kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

def extract_kml_from_kmz(kmz_bytes):
    with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
        kml_name = [f for f in z.namelist() if f.endswith(".kml")][0]
        with z.open(kml_name) as kml_file:
            tree = ET.parse(kml_file)
            return tree.getroot()

def extract_placemarks(elem, ns, folder=None):
    placemarks = []
    for child in elem:
        tag = child.tag.split("}")[-1]
        if tag == "Folder":
            folder_name_el = child.find("ns0:name", ns)
            folder_name = folder_name_el.text if folder_name_el is not None else folder
            placemarks += extract_placemarks(child, ns, folder_name)
        elif tag == "Placemark":
            name_el = child.find("ns0:name", ns)
            coord_el = child.find(".//ns0:coordinates", ns)
            if name_el is not None and coord_el is not None:
                name = name_el.text.strip()
                lon, lat, *_ = coord_el.text.strip().split(",")
                placemarks.append({
                    "folder": folder,
                    "name": name,
                    "lat": float(lat.strip()),
                    "lon": float(lon.strip())
                })
    return placemarks

def find_matching_pole(fat_point, poles, tolerance=0.0001):
    for p in poles:
        if abs(p["lat"] - fat_point["lat"]) < tolerance and abs(p["lon"] - fat_point["lon"]) < tolerance:
            return p["name"]
    return "POLE_NOT_FOUND"

if kmz_file:
    root = extract_kml_from_kmz(kmz_file.read())
    ns = {'ns0': 'http://www.opengis.net/kml/2.2'}
    placemarks = extract_placemarks(root, ns)

    df_all = pd.DataFrame(placemarks)
    st.subheader("📄 Daftar Semua Titik dari KMZ")
    st.dataframe(df_all)

if kmz_file and template_file:
    project_name = kmz_file.name.replace(".kmz", "")

    df_fat = [p for p in placemarks if p["folder"] and "FAT" in p["folder"].upper()]
    df_fdt = [p for p in placemarks if p["folder"] and "FDT" in p["folder"].upper()]
    df_hp = [p for p in placemarks if "HP COVER" in (p["folder"] or "").upper()]
    df_pole = [p for p in placemarks if "NEW POLE 7-3" in (p["folder"] or "").upper()]

    df_template = pd.read_excel(template_file)

    for i in range(min(len(df_hp), len(df_template))):
        if i < len(df_fat):
            fat = df_fat[i]
            df_template.at[i, "FATID"] = fat["name"]
            df_template.at[i, "Pole Latitude"] = fat["lat"]
            df_template.at[i, "Pole Longitude"] = fat["lon"]
            df_template.at[i, "Pole ID"] = find_matching_pole(fat, df_pole)
            df_template.at[i, "fdtcode"] = df_fdt[i]["name"] if i < len(df_fdt) else f"FDT_{i+1}"
            df_template.at[i, "Clustername"] = project_name
            df_template.at[i, "Commercial_name"] = project_name

        df_template.at[i, "homenumber"] = df_hp[i]["name"]
        df_template.at[i, "Latitude_homepass"] = df_hp[i]["lat"]
        df_template.at[i, "Longitude_homepass"] = df_hp[i]["lon"]

    st.success("✅ Data berhasil dimasukkan ke dalam TEMPLATE.")

    st.dataframe(df_template.head(10))

    output = BytesIO()
    df_template.to_excel(output, index=False)
    st.download_button("📅 Download File Hasil", output.getvalue(), file_name="TEMPLATE_HASIL_HPDB.xlsx")
