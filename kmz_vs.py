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

    # ---------- parse KML (single recursion from Document to avoid duplicates) ----------
    def recurse_folder(folder, ns, path=""):
        items = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None and name_el.text else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name
        # subfolders
        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns, new_path)
        # placemarks
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
        # ambil element Document, kalau tidak ada pakai root
        doc = root.find("kml:Document", ns) or root
        placemarks = recurse_folder(doc, ns, path="")

    # ---------- helper: robust number extractor ----------
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
        # jika multiple parts dan last part len == 3 => gabungkan (thousand grouping)
        if len(parts) >= 2 and len(parts[-1]) == 3:
            return int(''.join(parts))
        # lain-lain (kemungkinan ada decimal) -> ambil integer part pertama
        return int(parts[0])

    # helper filter
    def get_items(foldername):
        fn = foldername.upper()
        return [p for p in placemarks if fn in p["path"]]

    # ---------- collect categories ----------
    dist = get_items("DISTRIBUTION CABLE")
    sling = get_items("SLING WIRE")
    fat = get_items("FAT")
    new_pole_74 = get_items("NEW POLE 7-4")
    new_pole_73 = get_items("NEW POLE 7-3")
    exist_pole = get_items("EXISTING POLE EMR 7-4") + get_items("EXISTING POLE EMR 7-3")

    # ---------- calculations ----------
    lines = ["A", "B", "C", "D"]

    # Distribution: sum angka dari description per LINE
    dist_per_line = {L: 0 for L in lines}
    for p in dist:
        for L in lines:
            if f"LINE {L}" in p["path"]:
                v = extract_number_from_string(p.get("desc", ""))
                if v is not None:
                    dist_per_line[L] += v
                break

    # Sling: jumlah angka dari name hanya untuk LINE A/B/C (sesuai permintaan)
    sling_items_abc = [p for p in sling if any(f"LINE {L}" in p["path"] for L in ["A", "B", "C"])]
    sling_nums = []
    for p in sling_items_abc:
        v = extract_number_from_string(p.get("name", ""))
        if v is not None:
            sling_nums.append(v)
    sling_total = sum(sling_nums)

    # FAT per line & total
    fat_counts = {L: len([p for p in fat if f"LINE {L}" in p["path"]]) for L in lines}
    total_fat = sum(fat_counts.values())

    # Poles counts (hitung total untuk A,B,C lines)
    def count_items_in_lines(items, lines_list=["A", "B", "C"]):
        return len([p for p in items if any(f"LINE {L}" in p["path"] for L in lines_list)])

    np74_count = count_items_in_lines(new_pole_74, ["A", "B", "C"])
    np73_count = count_items_in_lines(new_pole_73, ["A", "B", "C"])
    exist_count = count_items_in_lines(exist_pole, ["A", "B", "C"])

    # ---------- show summary (debug / verifikasi) ----------
    st.markdown("### 🔎 Summary from KMZ (raw counts)")
    df_summary = pd.DataFrame([
        {"metric": "Distribution A (desc sum)", "value": dist_per_line["A"]},
        {"metric": "Distribution B (desc sum)", "value": dist_per_line["B"]},
        {"metric": "Distribution C (desc sum)", "value": dist_per_line["C"]},
        {"metric": "Distribution D (desc sum)", "value": dist_per_line["D"]},
        {"metric": "Sling total (A+B+C from name)", "value": sling_total},
        {"metric": "FAT total (all lines)", "value": total_fat},
        {"metric": "FAT A (count)", "value": fat_counts["A"]},
        {"metric": "FAT B (count)", "value": fat_counts["B"]},
        {"metric": "FAT C (count)", "value": fat_counts["C"]},
        {"metric": "FAT D (count)", "value": fat_counts["D"]},
        {"metric": "NEW POLE 7-4 (A/B/C)", "value": np74_count},
        {"metric": "NEW POLE 7-3 (A/B/C)", "value": np73_count},
        {"metric": "EXISTING POLE EMR 7-4+7-3 (A/B/C)", "value": exist_count},
    ])
    st.table(df_summary)

    # ---------- write into Excel (BoM AE) using the exact mapping you requested ----------
    wb = load_workbook(template_file)
    ws = wb["BoM AE"]

    # Distribution mapping (cells per range)
    line_map = {"A": [2, 6, 10], "B": [3, 7, 11], "C": [4, 8, 12], "D": [5, 9, 13]}
    for line, cells in line_map.items():
        # jumlah titik FAT pada line
        cnt = fat_counts.get(line, 0)
        # jumlah distribution (sum angka dari description) pada line
        val = dist_per_line.get(line, 0)
        if cnt >= 1 and cnt <= 10:
            ws[f"C{cells[0]}"] = val
        elif cnt >= 11 and cnt <= 15:
            ws[f"C{cells[1]}"] = val
        elif cnt >= 16 and cnt <= 20:
            ws[f"C{cells[2]}"] = val
        # jika cnt = 0 maka tidak menulis apa-apa

    # Sling wire -> C15
    ws["C15"] = sling_total

    # FAT total -> C30 / C31 / C32
    if 1 <= total_fat <= 24:
        ws["C30"] = total_fat
    elif 25 <= total_fat <= 36:
        ws["C31"] = total_fat
    elif 37 <= total_fat <= 48:
        ws["C32"] = total_fat

    # FAT per line -> C36..C39
    ws["C36"] = fat_counts["A"]
    ws["C37"] = fat_counts["B"]
    ws["C38"] = fat_counts["C"]
    ws["C39"] = fat_counts["D"]

    # Poles
    ws["C54"] = np74_count
    ws["C55"] = np73_count
    ws["C60"] = exist_count

    # ---------- save and provide download ----------
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.success("✅ BOQ berhasil dibuat! (download di bawah)")
    st.download_button("📥 Download BOQ", buf.getvalue(), file_name="hasil_BOQ.xlsx")
