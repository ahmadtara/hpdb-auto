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

def extract_placemarks(elem, ns, folder_path=""):
    placemarks = []
    for child in elem:
        tag = child.tag.split("}")[-1]
        if tag == "Folder":
            folder_name_el = child.find("ns0:name", ns)
            folder_name = folder_name_el.text.strip() if folder_name_el is not None else ""
            full_path = f"{folder_path}/{folder_name}" if folder_path else folder_name
            placemarks += extract_placemarks(child, ns, full_path)
        elif tag == "Placemark":
            name_el = child.find("ns0:name", ns)
            name = name_el.text.strip() if name_el is not None else "Unnamed"

            # Ambil koordinat dari Point, LineString, atau Polygon
            coord_el = None
            for geom_tag in ["Point", "LineString", "Polygon"]:
                geom = child.find(f".//ns0:{geom_tag}", ns)
                if geom is not None:
                    coord_el = geom.find(".//ns0:coordinates", ns)
                    if coord_el is not None:
                        break

            if coord_el is not None:
                coord_text = coord_el.text.strip()
                coord_pairs = coord_text.split()
                if coord_pairs:
                    lon_lat = coord_pairs[0].split(",")
                    if len(lon_lat) >= 2:
                        try:
                            lon = float(lon_lat[0].strip())
                            lat = float(lon_lat[1].strip())
                            placemarks.append({
                                "folder": folder_path,
                                "name": name,
                                "lat": lat,
                                "lon": lon
                            })
                        except ValueError:
                            continue
    return placemarks

def find_matching_pole(fat_point, poles, tolerance=0.0001):
    for p in poles:
        if abs(p["lat"] - fat_point["lat"]) < tolerance and abs(p["lon"] - fat_point["lon"]) < tolerance:
            return p["name"]
    return "POLE_NOT_FOUND"

if kmz_file and template_file:
    root = extract_kml_from_kmz(kmz_file.read())
    ns = {'ns0': 'http://www.opengis.net/kml/2.2'}
    placemarks = extract_placemarks(root, ns)

    st.write("📌 Total Placemarks terbaca:", len(placemarks))
    df_all = pd.DataFrame(placemarks)
    st.dataframe(
