import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import requests

HERE_API_KEY = "iWCrFicKYt9_AOCtg76h76MlqZkVTn94eHbBl_cE8m0"

st.title("📍 KMZ ➜ HPDB (Auto-Pilot ⚡By.A.Tara-P.)")
kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

def extract_placemarks(kmz_bytes):
    def recurse_folder(folder, ns, path=""):
        items = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name
        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns, new_path)
        for pm in folder.findall("kml:Placemark", ns):
            nm = pm.find("kml:name", ns)
            coord = pm.find(".//kml:coordinates", ns)
            if nm is not None and coord is not None:
                lon, lat = coord.text.strip().split(",")[:2]
                items.append({
                    "name": nm.text.strip(),
                    "lat": float(lat),
                    "lon": float(lon),
                    "path": new_path
                })
        return items

    with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
        f = [f for f in z.namelist() if f.lower().endswith(".kml")][0]
        root = ET.parse(z.open(f)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        all_pm = []
        for folder in root.findall(".//kml:Folder", ns):
            all_pm += recurse_folder(folder, ns)
        data = {k: [] for k in ["FAT", "NEW POLE 7-3", "EXISTING POLE EMR 7-3", "EXISTING POLE EMR 7-4", "FDT", "HP COVER"]}
        for p in all_pm:
            for k in data:
                if k in p["path"]:
                    data[k].append(p)
                    break
        return data

def extract_fatcode(path):
    for part in path.split("/"):
        if len(part) == 3 and part[0] in "ABCD" and part[1:].isdigit():
            return part
    return "UNKNOWN"

def reverse_here(lat, lon):
    url = (
        f"https://revgeocode.search.hereapi.com/v1/revgeocode"
        f"?at={lat},{lon}&apikey={HERE_API_KEY}&lang=en-US"
    )
    r = requests.get(url)
    if r.status_code == 200:
        comp = r.json().get("items", [{}])[0].get("address", {})
        return {
            "district": comp.get("district", "").upper(),
            "subdistrict": comp.get("subdistrict", "").upper().replace("KEL.", "").strip(),
            "postalcode": comp.get("postalCode", "").upper(),
            "street": comp.get("street", "").upper()
        }
    return {"district": "", "subdistrict": "", "postalcode": "", "street": ""}

if kmz_file and template_file:
    kmz_bytes = kmz_file.read()
    placemarks = extract_placemarks(kmz_bytes)
    df = pd.read_excel(template_file)
    fat = placemarks["FAT"]
    hp = placemarks["HP COVER"]
    fdt = placemarks["FDT"]
    all_poles = placemarks["NEW POLE 7-3"] + placemarks["EXISTING POLE EMR 7-3"] + placemarks["EXISTING POLE EMR 7-4"]

    if fdt:
        rc = reverse_here(fdt[0]["lat"], fdt[0]["lon"])
    else:
        rc = {"district": "", "subdistrict": "", "postalcode": "", "street": ""}

    fdtcode = "UNKNOWN"
    oltcode = "UNKNOWN"

    if fdt:
        fdtcode = fdt[0]["name"].strip().upper()

        with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
            f = [f for f in z.namelist() if f.lower().endswith(".kml")][0]
            tree = ET.parse(z.open(f))
            root = tree.getroot()
            ns = {"kml": "http://www.opengis.net/kml/2.2"}
            for pm in root.findall(".//kml:Placemark", ns):
                name_el = pm.find("kml:name", ns)
                desc_el = pm.find("kml:description", ns)
                if name_el is not None and name_el.text.strip().upper() == fdtcode:
                    if desc_el is not None:
                        oltcode = desc_el.text.strip().upper()
                    break

    progress = st.progress(0)
    total = len(hp)

    for col in ["block", "homenumber", "fdtcode", "oltcode"]:
        if col not in df.columns:
            df[col] = ""

    for i, h in enumerate(hp):
        if i >= len(df):
            break

        fc = extract_fatcode(h["path"])
        df.at[i, "fatcode"] = fc

        name_parts = h["name"].split(".")
        if len(name_parts) == 2 and name_parts[0].isalnum() and name_parts[1].isdigit():
            df.at[i, "block"] = name_parts[0].strip().upper()
            df.at[i, "homenumber"] = name_parts[1].strip()
        else:
            df.at[i, "block"] = ""
            df.at[i, "homenumber"] = h["name"]

        df.at[i, "Latitude_homepass"] = h["lat"]
        df.at[i, "Longitude_homepass"] = h["lon"]
        df.at[i, "district"] = rc["district"]
        df.at[i, "subdistrict"] = rc["subdistrict"]
        df.at[i, "postalcode"] = rc["postalcode"]
        df.at[i, "fdtcode"] = fdtcode
        df.at[i, "oltcode"] = oltcode

        hh = reverse_here(h["lat"], h["lon"])
        df.at[i, "street"] = hh["street"].replace("JALAN ", "").strip()

        mf = next((x for x in fat if fc in x["name"]), None)
        if mf:
            df.at[i, "FAT ID"] = mf["name"]
            df.at[i, "Pole Latitude"] = mf["lat"]
            df.at[i, "Pole Longitude"] = mf["lon"]
            pol = next((p["name"] for p in all_poles if abs(p["lat"] - mf["lat"]) < 1e-4 and abs(p["lon"] - mf["lon"]) < 1e-4), "POLE_NOT_FOUND")
            df.at[i, "Pole ID"] = pol
            fataddr = reverse_here(mf["lat"], mf["lon"])["street"]
            df.at[i, "FAT Address"] = fataddr
        else:
            df.at[i, "FAT ID"] = "FAT_NOT_FOUND"
            df.at[i, "Pole ID"] = "POLE_NOT_FOUND"
            df.at[i, "FAT Address"] = ""

        progress.progress(int((i + 1) * 100 / total))
    df["O4"] = len(hp)
    progress.empty()
    st.success("✅ Selesai!")

    st.dataframe(df.head(10))
    buf = BytesIO()
    df.to_excel(buf, index=False)
    st.download_button("📥 Download Hasil", buf.getvalue(), file_name="hasil_hpdb.xlsx")
