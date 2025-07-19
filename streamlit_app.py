import streamlit as st
import zipfile
import os
import tempfile
import xml.etree.ElementTree as ET
import pandas as pd
from openpyxl import load_workbook

def extract_kml(kmz_file, extract_path):
    with zipfile.ZipFile(kmz_file, 'r') as kmz:
        kmz.extractall(extract_path)
    for root, _, files in os.walk(extract_path):
        for file in files:
            if file.endswith(".kml"):
                return os.path.join(root, file)
    return None

def parse_kml(kml_path):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    tree = ET.parse(kml_path)
    root = tree.getroot()
    document = root.find('kml:Document', ns)

    fat_ids, poles, homes, fdt_codes = [], [], [], []

    for folder in document.findall('.//kml:Folder', ns):
        folder_name_elem = folder.find('kml:name', ns)
        folder_name = folder_name_elem.text.strip() if folder_name_elem is not None else ''

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

def extract_cluster_and_commercial(filename):
    title = os.path.splitext(os.path.basename(filename))[0]
    cluster = title.split(' - ')[0].strip()
    return cluster, cluster  # Commercial_name = Clustername

def write_to_template(fat_ids, poles, homes, fdt_codes, clustername, commercial_name, output_path):
    template_path = "TEMPLATE HPDB.xlsx"
    wb = load_workbook(template_path)
    ws = wb["Homepass Database"]

    start_row = 2  # row 1 = header
    max_len = max(len(fat_ids), len(poles), len(homes), len(fdt_codes))

    for i in range(max_len):
        ws.cell(row=start_row + i, column=1, value=fat_ids[i] if i < len(fat_ids) else '')
        ws.cell(row=start_row + i, column=2, value=poles[i][0] if i < len(poles_
