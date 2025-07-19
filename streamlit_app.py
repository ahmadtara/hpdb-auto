import streamlit as st
import pandas as pd
import zipfile, xml.etree.ElementTree as ET
import requests
from io import BytesIO
import time

# Fungsi reverse geocode
def rev_geocode(lat, lon):
    url = "https://nominatim.openstreetmap.org/reverse"
    params = {
        "format":"json",
        "lat": lat,
        "lon": lon,
        "addressdetails": 1
    }
    r = requests.get(url, params=params, headers={"User-Agent":"streamlit-app"})
    if r.status_code != 200:
        return {"postcode":"", "suburb":"", "city_district":"", "road":""}
    data = r.json().get("address", {})
    return {
        "postalcode": data.get("postcode", ""),
        "district": data.get("suburb", "").upper(),             # kelurahan
        "subdistrict": data.get("city_district", "").upper(),   # kecamatan
        "street": data.get("road", "")
    }

# ... fungsi extract_placemarks & helper lainnya tetap sama ...


if kmz_file and template_file:
    kmz_name = kmz_file.name.replace(".kmz", "")
    placemarks = extract_placemarks(kmz_file.read())
    df = pd.read_excel(template_file)
    
    fat_list = placemarks["FAT"]
    hp_list = placemarks["HP COVER"]
    all_poles = placemarks["NEW POLE 7-3"] + placemarks["EXISTING POLE EMR 7-3"] + placemarks["EXISTING POLE EMR 7-4"]
    fdtcode = placemarks["FDT"][0]["name"] if placemarks["FDT"] else "FDT_UNKNOWN"
    
    for i, hp in enumerate(hp_list):
        if i >= len(df): break
        
        lat, lon = hp["lat"], hp["lon"]
        geo = rev_geocode(lat, lon)
        time.sleep(1)  # rate limit
        
        fatcode = extract_fatcode_from_path(hp["path"])
        df.at[i, "fatcode"] = fatcode
        df.at[i, "homenumber"] = hp["name"]
        df.at[i, "Latitude_homepass"] = lat
        df.at[i, "Longitude_homepass"] = lon
        
        geo_keys = ["postalcode", "district", "subdistrict", "street"]
        for k in geo_keys:
            df.at[i, k] = geo.get(k, "")
        
        matched_fat = find_fat_by_fatcode(fatcode, fat_list)
        if matched_fat:
            df.at[i, "FAT ID"] = matched_fat["name"]
            df.at[i, "Pole Latitude"] = matched_fat["lat"]
            df.at[i, "Pole Longitude"] = matched_fat["lon"]
            df.at[i, "Pole ID"] = find_matching_pole(matched_fat, all_poles)
        else:
            df.at[i, "FAT ID"] = "FAT_NOT_FOUND"
            df.at[i, "Pole ID"] = "POLE_NOT_FOUND"
            
        df.at[i, "fdtcode"] = fdtcode
        df.at[i, "Clustername"] = kmz_name
        df.at[i, "Commercial_name"] = kmz_name
    
    st.success("✅ Data berhasil diisi dengan reverse‑geocoding.")
    st.dataframe(df.head(10))
    out = BytesIO()
    df.to_excel(out, index=False)
    st.download_button("📥 Unduh Hasil", out.getvalue(), file_name="HASIL_HPDB.xlsx")
