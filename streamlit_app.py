import streamlit as st
import zipfile
import os
import shutil
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

# Fungsi untuk mengekstrak KML dari KMZ
def extract_kml_from_kmz(kmz_bytes, extract_folder):
    with zipfile.ZipFile(kmz_bytes, 'r') as kmz:
        kmz.extractall(extract_folder)
    for root, _, files in os.walk(extract_folder):
        for file in files:
            if file.endswith('.kml'):
                return os.path.join(root, file)
    return None

# Parsing isi file KML
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

# Ekstrak nama cluster & commercial
def extract_cluster_and_commercial(file_name):
    title = os.path.splitext(file_name)[0]
    parts = title.split(' - ', 1)
    clustername = parts[0].strip() if len(parts) > 0 else ''
    commercial_name = parts[1].strip() if len(parts) > 1 else ''
    return clustername, commercial_name

# Membentuk DataFrame
def build_dataframe(fat_ids, poles, homes, fdt_codes, clustername, commercial_name):
    max_len = max(len(fat_ids), len(poles), len(homes), len(fdt_codes))
    rows = []
    for i in range(max_len):
        row = {
            'Fat ID': fat_ids[i] if i < len(fat_ids) else '',
            'Pole ID': poles[i][0] if i < len(poles) else '',
            'Pole Latitude': poles[i][1] if i < len(poles) else '',
            'Pole Longitude': poles[i][2] if i < len(poles) else '',
            'homenumber': homes[i][0] if i < len(homes) else '',
            'Latitude_homepass': homes[i][1] if i < len(homes) else '',
            'Longitude_homepass': homes[i][2] if i < len(homes) else '',
            'fdtcode': fdt_codes[i] if i < len(fdt_codes) else '',
            'Clustername': clustername,
            'Commercial_name': commercial_name
        }
        rows.append(row)
    return pd.DataFrame(rows)

# Antarmuka Streamlit
st.title("KMZ ➜ Excel | Homepass Database Generator")
uploaded_kmz = st.file_uploader("Upload file .kmz", type=["kmz"])

if uploaded_kmz is not None:
    with st.spinner("⏳ Memproses file..."):
        temp_dir = 'temp_kmz_extract'
        os.makedirs(temp_dir, exist_ok=True)

        # Simpan file sementara
        kmz_file_path = os.path.join(temp_dir, uploaded_kmz.name)
        with open(kmz_file_path, "wb") as f:
            f.write(uploaded_kmz.getbuffer())

        # Ekstrak & parsing
        kml_file = extract_kml_from_kmz(kmz_file_path, temp_dir)
        if not kml_file:
            st.error("❌ Gagal mengekstrak file KML.")
        else:
            fat_ids, poles, homes, fdt_codes = parse_kml(kml_file)
            clustername, commercial_name = extract_cluster_and_commercial(uploaded_kmz.name)
            df = build_dataframe(fat_ids, poles, homes, fdt_codes, clustername, commercial_name)

            st.success("✅ Data berhasil diproses.")
            st.dataframe(df)

            # Export ke Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name="Homepass Database")
            st.download_button("📥 Download HPDB_Output.xlsx", data=output.getvalue(),
                               file_name="HPDB_Output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        shutil.rmtree(temp_dir)
