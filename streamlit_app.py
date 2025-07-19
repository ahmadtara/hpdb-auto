import streamlit as st
import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import tempfile
from openpyxl import load_workbook

def extract_kml(kmz_path, extract_folder):
    with zipfile.ZipFile(kmz_path, 'r') as kmz:
        kmz.extractall(extract_folder)
    for root, _, files in os.walk(extract_folder):
        for file in files:
            if file.endswith('.kml'):
                return os.path.join(root, file)
    return None

def parse_kml(kml_file):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    tree = ET.parse(kml_file)
    root = tree.getroot()
    document = root.find('kml:Document', ns)

    fat_ids, poles, homes, fdt_codes = [], [], [], []

    for folder in document.findall('.//kml:Folder', ns):
        folder_name = folder.find('kml:name', ns)
        folder_name = folder_name.text if folder_name is not None else ''

        for placemark in folder.findall('kml:Placemark', ns):
            name_elem = placemark.find('kml:name', ns)
            coord_elem = placemark.find('.//kml:coordinates', ns)
            if name_elem is None or coord_elem is None:
                continue

            name = name_elem.text.strip()
            coord_text = coord_elem.text.strip()
            lon, lat, *_ = coord_text.split(',')

            lat, lon = float(lat), float(lon)

            if 'FAT' in folder_name.upper():
                fat_ids.append(name)
            elif 'NEW POLE 7-3' in folder_name.upper():
                poles.append((name, lat, lon))
            elif 'HP COVER' in folder_name.upper():
                homes.append((name, lat, lon))
            elif 'FDT' in folder_name.upper():
                fdt_codes.append(name)

    return fat_ids, poles, homes, fdt_codes

def extract_cluster_and_commercial(kmz_filename):
    filename = os.path.basename(kmz_filename)
    title = os.path.splitext(filename)[0]
    clustername = title.strip()
    commercial_name = clustername  # Disamakan
    return clustername, commercial_name

def fill_template(template_path, output_path, fat_ids, poles, homes, fdt_codes, clustername, commercial_name):
    df_template = pd.read_excel(template_path, sheet_name="Homepass Database")

    max_len = max(len(fat_ids), len(poles), len(homes), len(fdt_codes))
    df_filled = df_template.copy()

    for i in range(max_len):
        if i >= len(df_filled):
            df_filled.loc[i] = None  # Tambah baris jika belum ada

        df_filled.loc[i, 'Fat ID'] = fat_ids[i] if i < len(fat_ids) else ''
        df_filled.loc[i, 'Pole ID'] = poles[i][0] if i < len(poles) else ''
        df_filled.loc[i, 'Pole Latitude'] = poles[i][1] if i < len(poles) else ''
        df_filled.loc[i, 'Pole Longitude'] = poles[i][2] if i < len(poles) else ''
        df_filled.loc[i, 'homenumber'] = homes[i][0] if i < len(homes) else ''
        df_filled.loc[i, 'Latitude_homepass'] = homes[i][1] if i < len(homes) else ''
        df_filled.loc[i, 'Longitude_homepass'] = homes[i][2] if i < len(homes) else ''
        df_filled.loc[i, 'fdtcode'] = fdt_codes[i] if i < len(fdt_codes) else ''
        df_filled.loc[i, 'Clustername'] = clustername
        df_filled.loc[i, 'Commercial_name'] = commercial_name

    df_filled.to_excel(output_path, sheet_name='Homepass Database', index=False)

# === Streamlit UI ===

st.title("📍 KMZ to HPDB (Template Excel) Converter")

kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB.xlsx", type=["xlsx"])

if kmz_file and template_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        kmz_path = os.path.join(tmpdir, kmz_file.name)
        with open(kmz_path, 'wb') as f:
            f.write(kmz_file.read())

        template_path = os.path.join(tmpdir, template_file.name)
        with open(template_path, 'wb') as f:
            f.write(template_file.read())

        extract_folder = os.path.join(tmpdir, 'kmz_extract')
        os.makedirs(extract_folder, exist_ok=True)

        kml_file = extract_kml(kmz_path, extract_folder)
        if not kml_file:
            st.error("❌ Gagal membaca KML dari KMZ")
        else:
            fat_ids, poles, homes, fdt_codes = parse_kml(kml_file)
            clustername, commercial_name = extract_cluster_and_commercial(kmz_path)

            output_path = os.path.join(tmpdir, 'HPDB_Output.xlsx')
            fill_template(template_path, output_path, fat_ids, poles, homes, fdt_codes, clustername, commercial_name)

            st.success("✅ Data berhasil dimasukkan ke Template!")

            with open(output_path, 'rb') as f:
                st.download_button(
                    label="⬇️ Download Excel Hasil",
                    data=f,
                    file_name="HPDB_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
