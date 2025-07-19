import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

st.title("📍 Ambil Data FAT dari KMZ ke Template HPDB")

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
            for pm in root.findall(".//kml:Placemark", ns):
                folder = pm.find("../kml:name", ns)
                folder_name = folder.text if folder is not None else ""
                if "FAT" in folder_name.upper():
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
    fat_data = extract_fat_from_kmz(kmz_file.read())
    df_fat = pd.DataFrame(fat_data)

    st.write("✅ Jumlah FAT ditemukan:", len(df_fat))
    st.dataframe(df_fat)

    df_template = pd.read_excel(template_file)

    for i in range(min(len(df_fat), len(df_template))):
        df_template.at[i, "FATID"] = df_fat[i]["name"]
        df_template.at[i, "Pole Latitude"] = df_fat[i]["lat"]
        df_template.at[i, "Pole Longitude"] = df_fat[i]["lon"]

    st.success("✅ Data FAT berhasil dimasukkan ke dalam template.")
    st.dataframe(df_template.head(10))

    output = BytesIO()
    df_template.to_excel(output, index=False)
    st.download_button("📥 Download File Hasil", output.getvalue(), file_name="HASIL_TEMPLATE_HPDB.xlsx")
