import streamlit as st
import zipfile
import os
import xml.etree.ElementTree as ET
import pandas as pd
import tempfile

def run_kmz_vs_hpdb():
    st.set_page_config(page_title="KMZ vs Template XLSX Checker", layout="centered")
    st.title("📍 Validasi HP COVER: KMZ vs Template XLSX")
    st.markdown("""
    ### ✨ Upload KMZ dan Template Excel
    Akan dicek apakah semua `block.homenumber` di folder **HP COVER** dari file KMZ cocok dengan yang ada di template Excel.
    """)

    kmz_file = st.file_uploader("📁 Upload File KMZ", type=["kmz"])
    template_file = st.file_uploader("📄 Upload Template XLSX", type=["xlsx"])

    target_folder = "HP COVER"

    def extract_placemarks(kmz_bytes):
        placemarks = {}
        with tempfile.TemporaryDirectory() as tmpdir:
            kmz_path = os.path.join(tmpdir, "temp.kmz")
            with open(kmz_path, "wb") as f:
                f.write(kmz_bytes)

            with zipfile.ZipFile(kmz_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)

            kml_path = os.path.join(tmpdir, "doc.kml")
            tree = ET.parse(kml_path)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}

            folders = root.findall('.//kml:Folder', ns)
            for folder in folders:
                name_tag = folder.find('kml:name', ns)
                if name_tag is None:
                    continue
                folder_name = name_tag.text.strip().upper()
                if folder_name != target_folder:
                    continue
                placemarks[folder_name] = []

                for pm in folder.findall('.//kml:Placemark', ns):
                    name = pm.find('kml:name', ns)
                    name_text = name.text.strip() if name is not None else ""
                    placemarks[folder_name].append({"name": name_text})

        return placemarks

    if kmz_file and template_file:
        kmz_bytes = kmz_file.read()
        placemarks = extract_placemarks(kmz_bytes)
        df = pd.read_excel(template_file)

        if target_folder not in placemarks:
            st.error(f"❌ Folder '{target_folder}' tidak ditemukan dalam KMZ!")
        else:
            hp = placemarks[target_folder]

            # Ambil block & homenumber dari KMZ
            kmz_blocks_homenumbers = set()
            for h in hp:
                name_parts = h["name"].split(".")
                if len(name_parts) == 2:
                    block, homenumber = name_parts[0].strip().upper(), name_parts[1].strip()
                    kmz_blocks_homenumbers.add((block, homenumber))
                else:
                    st.warning(f"❗ Format salah di placemark: '{h['name']}' (abaikan)")

            # Ambil block & homenumber dari XLSX
            xlsx_blocks_homenumbers = set()
            if "block" not in df.columns or "homenumber" not in df.columns:
                st.error("❌ Kolom 'block' dan 'homenumber' harus ada di template XLSX!")
            else:
                for _, row in df.iterrows():
                    block = str(row["block"]).strip().upper()
                    homenumber = str(row["homenumber"]).strip()
                    xlsx_blocks_homenumbers.add((block, homenumber))

                # Cek selisih dua set
                kmz_only = kmz_blocks_homenumbers - xlsx_blocks_homenumbers
                xlsx_only = xlsx_blocks_homenumbers - kmz_blocks_homenumbers

                if not kmz_only and not xlsx_only:
                    st.success("✅ Semua block dan homenumber di XLSX sesuai dengan HP COVER di KMZ!")
                else:
                    st.error("❌ Ada perbedaan antara XLSX dan HP COVER!")

                    diff_data = []

                    if kmz_only:
                        st.warning(f"🔹 Di KMZ (HP COVER) tapi TIDAK ADA di XLSX: {len(kmz_only)} item")
                        for b, h in kmz_only:
                            diff_data.append({"block": b, "homenumber": h, "sumber": "KMZ Only"})

                    if xlsx_only:
                        st.warning(f"🔸 Di XLSX tapi TIDAK ADA di KMZ (HP COVER): {len(xlsx_only)} item")
                        for b, h in xlsx_only:
                            diff_data.append({"block": b, "homenumber": h, "sumber": "XLSX Only"})

                    diff_df = pd.DataFrame(diff_data)
                    st.dataframe(diff_df, use_container_width=True)

                    # Download button
                    csv = diff_df.to_csv(index=False).encode("utf-8")
                    st.download_button("📥 Download Selisih CSV", data=csv, file_name="selisih_kmz_vs_xlsx.csv", mime="text/csv")
    else:
        st.info("⬆️ Silakan upload file KMZ dan template XLSX terlebih dahulu.")
