import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import re
from openpyxl import load_workbook
import pandas as pd

def run_boq():
    st.title("📊 KMZ ➜ BOQ - 1 FDT")

    kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
    template_file = st.file_uploader("Upload TEMPLATE BOQ - 1 FDT.xlsx", type=["xlsx"])

    if not (kmz_file and template_file):
        return

    kmz_bytes = kmz_file.read()

    # ---------- Parse KML ----------
    def recurse_folder(folder, ns, path=""):
        items = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None and name_el.text else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name
        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns, new_path)
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
        raw = z.read(kml_files[0])
        root = ET.fromstring(raw)
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        doc = root.find("kml:Document", ns) or root
        placemarks = recurse_folder(doc, ns, path="")

    # ---------- Helper ----------
    def extract_number_from_string(s):
        if not s:
            return None
        m = re.search(r'[\d.,]+', s)
        if not m:
            return None
        token = m.group(0)
        parts = re.findall(r'\d+', token)
        if not parts:
            return None
        if len(parts) == 1:
            return int(parts[0])
        if len(parts) >= 2 and len(parts[-1]) == 3:
            return int(''.join(parts))
        return int(parts[0])

    def get_items(foldername):
        fn = foldername.upper()
        return [p for p in placemarks if fn in p["path"].split("/")]

    lines = ["A", "B", "C", "D"]

    # ---------- Kumpulan Item ----------
    dist = get_items("DISTRIBUTION CABLE")
    sling = get_items("SLING WIRE")
    fat = get_items("FAT")
    hp_cover = get_items("HP COVER")

    # ---------- Distribution Cable ----------
    dist_per_line = {L: 0 for L in lines}
    for p in dist:
        for L in lines:
            if f"LINE {L}" in p["path"].split("/"):
                v = extract_number_from_string(p.get("desc", ""))
                if v is not None:
                    dist_per_line[L] += v
                break

    # ---------- Sling Wire ----------
    sling_items_abcd = [p for p in sling if any(f"LINE {L}" in p["path"].split("/") for L in lines)]
    sling_nums = [extract_number_from_string(p.get("name", "")) for p in sling_items_abcd if extract_number_from_string(p.get("name", "")) is not None]
    sling_total = sum(sling_nums)

    # ---------- FAT ----------
    fat_counts = {L: len([p for p in fat if f"LINE {L}" in p["path"].split("/")]) for L in lines}
    total_fat = sum(fat_counts.values())

    # ---------- HP COVER ----------
    hp_cover_counts = {L: len([p for p in hp_cover if f"LINE {L}" in p["path"].split("/")]) for L in lines}
    total_hp_cover = sum(hp_cover_counts.values())

    # ---------- POLE ----------
    def count_pole_per_line(items):
        return {L: len([p for p in items if f"LINE {L}" in p["path"].split("/")]) for L in lines}

    new_74 = count_pole_per_line(get_items("NEW POLE 7-4"))
    new_73 = count_pole_per_line(get_items("NEW POLE 7-3"))
    new_725 = count_pole_per_line(get_items("NEW POLE 7-2.5"))
    new_94 = count_pole_per_line(get_items("NEW POLE 9-4"))

    exist_items = (
        get_items("EXISTING POLE EMR 7-4") +
        get_items("EXISTING POLE EMR 7-3") +
        get_items("EXISTING POLE EMR 7-2.5") +
        get_items("EXISTING POLE EMR 9-4")
    )
    exist_per_line = count_pole_per_line(exist_items)

    # ---------- Write ke Excel ----------
    wb = load_workbook(template_file)
    ws = wb["BoM AE"]

    # Mapping distribusi
    line_map = {"A": [2, 6, 10], "B": [3, 7, 11], "C": [4, 8, 12], "D": [5, 9, 13]}
    for line, cells in line_map.items():
        cnt = fat_counts.get(line, 0)
        val = dist_per_line.get(line, 0)
        if 1 <= cnt <= 10:
            ws[f"C{cells[0]}"] = val
        elif 11 <= cnt <= 15:
            ws[f"C{cells[1]}"] = val
        elif 16 <= cnt <= 20:
            ws[f"C{cells[2]}"] = val

    ws["C15"] = sling_total

    # FAT total
    if 1 <= total_fat <= 24:
        ws["C30"] = 1
    elif 25 <= total_fat <= 36:
        ws["C31"] = 1
    elif 37 <= total_fat <= 48:
        ws["C32"] = 1

    ws["C36"], ws["C37"], ws["C38"], ws["C39"] = fat_counts["A"], fat_counts["B"], fat_counts["C"], fat_counts["D"]

    # Pole totals
    ws["C54"] = sum(new_74.values())
    ws["C55"] = sum(new_73.values())
    ws["C56"] = sum(new_725.values())
    ws["C58"] = sum(new_94.values())
    ws["C61"] = sum(exist_per_line.values())

    # Sheet BoQ NRO Cluster
    ws2 = wb["BoQ NRO Cluster"]
    ws2["O5"] = total_hp_cover
    ws2["O3"] = kmz_file.name.rsplit(".", 1)[0]

    # ---------- OUTPUT ----------
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.success("✅ BOQ berhasil dibuat!")

    # ---------- Tampilkan tabel ringkasan ----------
    df_summary = pd.DataFrame({
        "Line": lines,
        "New Pole 7-4": [new_74[L] for L in lines],
        "New Pole 7-3": [new_73[L] for L in lines],
        "New Pole 7-2.5": [new_725[L] for L in lines],
        "New Pole 9-4": [new_94[L] for L in lines],
        "Existing Pole": [exist_per_line[L] for L in lines],
    })
    df_summary.loc["Total"] = df_summary.sum(numeric_only=True)
    df_summary.loc["Total", "Line"] = "TOTAL"

    st.dataframe(df_summary)

    st.download_button("📥 Download BOQ", buf.getvalue(), file_name="hasil_BOQ.xlsx")
