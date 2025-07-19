import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import requests

st.title("📍 KMZ ➜ HPDB (Auto‑Fill + Geocode)")

kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

def extract_placemarks(kmz_bytes):
    def recurse_folder(folder, ns, path=""):
        results = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name

        for sub in folder.findall("kml:Folder", ns):
            results += recurse_folder(sub, ns, new_path)
        for pm in folder.findall("kml:Placemark", ns):
            nm = pm.find("kml:name", ns)
            co = pm.find(".//kml:coordinates", ns)
            if nm is not None and co is not None:
                lon, lat = map(float, co.text.strip().split(",")[:2])
                results.append({"name": nm.text.strip(), "lat": lat, "lon": lon, "path": new_path})
        return results

    with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
        kf = next(f for f in z.namelist() if f.lower().endswith(".kml"))
        root = ET.parse(z.open(kf)).getroot()
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        all_pms = sum([recurse_folder(f, ns) for f in root.findall(".//kml:Folder", ns)], [])

        categories = [
            "FAT", "NEW POLE 7-3", "EXISTING POLE EMR 7-3",
            "EXISTING POLE EMR 7-4", "FDT", "HP COVER"
        ]
        data = {c: [] for c in categories}
        for p in all_pms:
            for c in data:
                if c in p["path"]:
                    data[c].append(p)
                    break
        return data

def extract_fatcode(path):
    for part in path.split("/"):
        if len(part) == 3 and part[0] in "ABCD" and part[1:].isdigit():
            return part
    return "UNKNOWN"

def find_fat(fatcode, fats):
    return next((f for f in fats if fatcode in f["name"]), None)

def find_pole_id(fat, poles, tol=1e-4):
    pole = next((p for p in poles if abs(p["lat"]-fat["lat"])<tol and abs(p["lon"]-fat["lon"])<tol), None)
    return pole["name"] if pole else "POLE_NOT_FOUND"

def reverse_geocode(lat, lon):
    url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lon}&zoom=18&addressdetails=1"
    resp = requests.get(url, headers={"User-Agent": "streamlit-app"})
    if not resp.ok:
        return "", "", ""
    ad = resp.json().get("address", {})
    return ad.get("postcode", ""), ad.get("city_district","") or ad.get("suburb",""), ad.get("village","") or ad.get("suburb","")

if kmz_file and template_file:
    data = extract_placemarks(kmz_file.read())
    df = pd.read_excel(template_file)

    fats, poles = data["FAT"], data["NEW POLE 7-3"] + data["EXISTING POLE EMR 7-3"] + data["EXISTING POLE EMR 7-4"]
    hp_list = data["HP COVER"]
    fdtcode = data["FDT"][0]["name"] if data["FDT"] else "FDT_UNKNOWN"

    # Ambil lokasi pertama HP COVER sebagai referensi geocode
    if not hp_list:
        st.error("⚠️ Folder HP COVER kosong!")
        st.stop()
    sample = hp_list[0]
    postalcode, subdistrict, district = reverse_geocode(sample["lat"], sample["lon"])

    for idx, hp in enumerate(hp_list):
        if idx >= len(df): break
        code = extract_fatcode(hp["path"])
        df.at[idx, "fatcode"] = code
        df.at[idx, ["homenumber","Latitude_homepass","Longitude_homepass"]] = [hp["name"], hp["lat"], hp["lon"]]

        fat = find_fat(code, fats)
        if fat:
            df.at[idx, ["FAT ID","Pole Latitude","Pole Longitude"]] = [fat["name"], fat["lat"], fat["lon"]]
            df.at[idx, "Pole ID"] = find_pole_id(fat, poles)
        else:
            df.at[idx, ["FAT ID","Pole ID"]] = ["FAT_NOT_FOUND","POLE_NOT_FOUND"]

        df.at[idx, "fdtcode"] = fdtcode
        df.at[idx, ["Clustername","Commercial_name"]] = [kmz_file.name.replace(".kmz","")] * 2

        df.at[idx, "postalcode"] = postalcode
        df.at[idx, "subdistrict"] = subdistrict
        df.at[idx, "district"] = district

    st.success("✅ Semua data berhasil diisi.")
    st.dataframe(df.head(10))

    buf = BytesIO()
    df.to_excel(buf, index=False)
    st.download_button("📥 Unduh Hasil", buf.getvalue(), file_name="HASIL_HPDB_FINAL.xlsx")
