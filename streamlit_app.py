import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import requests

st.title("📍 KMZ ➜ HPDB (Auto-Fill Extended)")

kmz_file = st.file_uploader("Upload file .KMZ", type=["kmz"])
template_file = st.file_uploader("Upload TEMPLATE HPDB (.xlsx)", type=["xlsx"])

def extract_placemarks(kmz_bytes):
    def recurse_folder(folder, ns, path=""):
        placemarks = []
        name_el = folder.find("kml:name", ns)
        new_path = f"{path}/{name_el.text.upper()}" if name_el is not None else path
        for sub in folder.findall("kml:Folder", ns):
            placemarks += recurse_folder(sub, ns, new_path)
        for pm in folder.findall("kml:Placemark", ns):
            nm = pm.find("kml:name", ns)
            co = pm.find(".//kml:coordinates", ns)
            if nm is not None and co is not None:
                lon, lat = map(float, co.text.strip().split(",")[:2])
                placemarks.append({"name": nm.text, "lat":lat, "lon":lon, "path":new_path})
        return placemarks
    with zipfile.ZipFile(BytesIO(kmz_bytes)) as z:
        kf = [f for f in z.namelist() if f.lower().endswith('.kml')][0]
        root = ET.parse(z.open(kf)).getroot()
        ns = {'kml':'http://www.opengis.net/kml/2.2'}
        all_pms = sum((recurse_folder(f,ns) for f in root.findall(".//kml:Folder",ns)), [])
        categories = ["FAT","NEW POLE 7-3","EXISTING POLE EMR 7-3","EXISTING POLE EMR 7-4","FDT","HP COVER"]
        data={k:[] for k in categories}
        for p in all_pms:
            for k in data:
                if k in p["path"]:
                    data[k].append(p); break
        return data

def extract_fatcode(path):
    for piece in path.split("/"):
        if len(piece)==3 and piece[0] in "ABCD" and piece[1:].isdigit(): return piece
    return "UNKNOWN"

def find_fat(fatcode, fats):
    return next((f for f in fats if fatcode in f["name"]), None)

def find_pole(fat, poles, tol=1e-4):
    return next((p["name"] for p in poles if abs(p["lat"]-fat["lat"])<tol and abs(p["lon"]-fat["lon"])<tol), "POLE_NOT_FOUND")

def reverse_geocode(lat, lon):
    url = ("https://nominatim.openstreetmap.org/reverse?format=json"
           f"&lat={lat}&lon={lon}&addressdetails=1")
    resp = requests.get(url, headers={'User-Agent':'streamlit-app'})
    if resp.ok:
        ad = resp.json().get("address",{})
        return ad.get("postcode",""), ad.get("suburb",""), ad.get("city_district","") or ad.get("county","")
    return "", "", ""

if kmz_file and template_file:
    data = extract_placemarks(kmz_file.read())
    df = pd.read_excel(template_file)
    fats, poles = data["FAT"], data["NEW POLE 7-3"]+data["EXISTING POLE EMR 7-3"]+data["EXISTING POLE EMR 7-4"]
    cycle = data["HP COVER"]
    fdtcode = (data["FDT"][0]["name"] if data["FDT"] else "FDT_UNKNOWN")
    zipc, subd, dist = ("","","")
    # get postalcode/district once
    if cycle:
        zipc, subd, dist = reverse_geocode(cycle[0]["lat"], cycle[0]["lon"])

    r=0
    for hp in cycle:
        if r>=len(df): break
        code=extract_fatcode(hp["path"])
        df.at[r,"fatcode"]=code
        df.at[r,["homenumber","Latitude_homepass","Longitude_homepass"]]=[hp["name"],hp["lat"],hp["lon"]]
        fat=find_fat(code,fats)
        if fat:
            df.at[r,["FAT ID","Pole Latitude","Pole Longitude"]]=[fat["name"],fat["lat"],fat["lon"]]
            df.at[r,"Pole ID"]=find_pole(fat,poles)
        else:
            df.at[r,["FAT ID","Pole ID"]]=["FAT_NOT_FOUND","POLE_NOT_FOUND"]
        df.at[r,"fdtcode"]=fdtcode
        df.at[r,["Clustername","Commercial_name"]]=[kmz_file.name.replace(".kmz","")]*2
        df.at[r,["postalcode","subdistrict","district"]]=[zipc,subd,dist]
        r+=1

    st.success("✅ Semua data terisi.")
    st.dataframe(df.head())

    buf=BytesIO(); df.to_excel(buf,index=False)
    st.download_button("📥 Unduh Hasil",buf.getvalue(),file_name="HASIL_HPDB_EXTENDED.xlsx")
