import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

st.title("📍 Konversi KMZ ➜ TEMPLATE HPDB (khusus FAT)")

kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

def extract_fat_from_kmz(kmz_bytes):
    with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
        kml_file = [f for f in z.namelist() if f.endswith('.kml')][0]
        with z.open(kml_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}

            placemarks = []
            for folder in root.findall(".//kml:Folder", ns):
                folder_name_el = folder.find("kml:name", ns)
                if folder_name_el is not None and "FAT" in folder_name_el.text.upper():
                    for pm in folder.findall("kml:Placemark", ns):
                        name_el = pm.find("kml:name", ns)
                        coord_el = pm.find(".//kml:coordinates", ns)
                        if name_el is not None and coord_el is not None:
                            coords = coord_el.text.strip().split(",")
                            if len(coords) >= 2:
                                placemarks.append({
                                    "name": name_el.text.strip(),
                                    "lat": float(coords[1].strip()),
                                    "lon": float(coords[0].strip())
                                })
            return placemarks

if kmz_file and template_file:
    fats = extract_fat_from_kmz(kmz_file.read())
    df_template = pd.read_excel(template_file)

    st.success(f"✅ Jumlah FAT ditemukan: {len(fats)}")
    st.write(fats)

    # Masukkan data FAT ke dalam template
    for i in range(min(len(fats), len(df_template))):
        df_template.at[i, "FAT ID"] = fats[i]["name"]
        df_template.at[i, "Pole Latitude"] = fats[i]["lat"]
        df_template.at[i, "Pole Longitude"] = fats[i]["lon"]

    st.success("✅ Data FAT berhasil dimasukkan ke dalam template.")
    st.dataframe(df_template.head(10))

    output = BytesIO()
    df_template.to_excel(output, index=False)
    st.download_button("📥 Download File Hasil", output.getvalue(), file_name="HASIL_TEMPLATE_HPDB_FAT.xlsx")
