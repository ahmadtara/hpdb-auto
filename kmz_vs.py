import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import re
from openpyxl import load_workbook

def run_boq():

    st.title("📊 KMZ ➜ BOQ - 1 FDT")

    kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
    template_file = st.file_uploader("Upload TEMPLATE BOQ - 1 FDT.xlsx", type=["xlsx"])

    if kmz_file and template_file:
        kmz_bytes = kmz_file.read()

        # --- Extract KMZ ---
        def recurse_folder(folder, ns, path=""):
            items = []
            name_el = folder.find("kml:name", ns)
            folder_name = name_el.text.upper() if name_el is not None and name_el.text else "UNKNOWN"
            new_path = f"{path}/{folder_name}" if path else folder_name
            # subfolder
            for sub in folder.findall("kml:Folder", ns):
                items += recurse_folder(sub, ns, new_path)
            # placemark
            for pm in folder.findall("kml:Placemark", ns):
                nm = pm.find("kml:name", ns)
                desc = pm.find("kml:description", ns)
                items.append({
                    "name": nm.text.strip() if nm is not None and nm.text else "",
                    "desc": desc.text.strip() if desc is not None and desc.text else "",
                    "path": new_path
                })
            return items

        with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
            kml_files = [f for f in z.namelist() if f.lower().endswith(".kml")]
            if not kml_files:
                st.error("❌ Tidak ditemukan file .kml di dalam .kmz")
                return
            root = ET.parse(z.open(kml_files[0])).getroot()
            ns = {"kml": "http://www.opengis.net/kml/2.2"}
            placemarks = []
            for folder in root.findall(".//kml:Folder", ns):
                placemarks += recurse_folder(folder, ns)

        # --- Helper ---
        def get_items(foldername):
            return [p for p in placemarks if foldername.upper() in p["path"]]

        def extract_number(txt):
            if not txt:
                return None
            m = re.search(r"(\d[\d\.,]*)", txt)
            if not m:
                return None
            s = m.group(1).replace(",", "").replace(".", "")
            try:
                return int(s)
            except:
                return None

        # --- Filter ---
        dist = get_items("DISTRIBUTION CABLE")
        sling = get_items("SLING WIRE")
        fat = get_items("FAT")
        new_pole_74 = get_items("NEW POLE 7-4")
        new_pole_73 = get_items("NEW POLE 7-3")
        exist_pole = get_items("EXISTING POLE EMR 7-4") + get_items("EXISTING POLE EMR 7-3")

        # --- Load Excel ---
        wb = load_workbook(template_file)
        ws = wb["BoM AE"]

        # --- Distribution Cable ---
        line_map = {"A": [2, 6, 10], "B": [3, 7, 11], "C": [4, 8, 12], "D": [5, 9, 13]}
        for line, cells in line_map.items():
            fat_line = [p for p in fat if f"LINE {line}" in p["path"]]
            cnt = len(fat_line)
            dist_vals = [extract_number(p["desc"]) for p in dist if f"LINE {line}" in p["path"]]
            dist_vals = [v for v in dist_vals if v is not None]
            if cnt in range(1, 11) and dist_vals:
                ws[f"C{cells[0]}"] = sum(dist_vals)
            elif cnt in range(11, 16) and dist_vals:
                ws[f"C{cells[1]}"] = sum(dist_vals)
            elif cnt in range(16, 21) and dist_vals:
                ws[f"C{cells[2]}"] = sum(dist_vals)

        # --- Sling Wire ---
        sling_total = sum([extract_number(p["name"]) for p in sling if extract_number(p["name"])])
        ws["C15"] = sling_total

        # --- FAT total ---
        total_fat = len(fat)
        if 1 <= total_fat <= 24:
            ws["C30"] = total_fat
        elif 25 <= total_fat <= 36:
            ws["C31"] = total_fat
        elif 37 <= total_fat <= 48:
            ws["C32"] = total_fat

        # --- FAT per Line ---
        ws["C36"] = len([p for p in fat if "LINE A" in p["path"]])
        ws["C37"] = len([p for p in fat if "LINE B" in p["path"]])
        ws["C38"] = len([p for p in fat if "LINE C" in p["path"]])
        ws["C39"] = len([p for p in fat if "LINE D" in p["path"]])

        # --- Poles ---
        ws["C54"] = len(new_pole_74)
        ws["C55"] = len(new_pole_73)
        ws["C60"] = len(exist_pole)

        # --- Save hasil ---
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        st.success("✅ BOQ berhasil dibuat!")
        st.download_button("📥 Download BOQ", buf.getvalue(), file_name="hasil_BOQ.xlsx")
