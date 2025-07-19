import zipfile
import xml.etree.ElementTree as ET

def extract_kml_from_kmz(path):
    with zipfile.ZipFile(path) as z:
        kml_name = [f for f in z.namelist() if f.endswith(".kml")][0]
        with z.open(kml_name) as kml_file:
            return ET.parse(kml_file).getroot()

def extract_placemarks_verbose(elem, ns, folder=None):
    placemarks = []
    for child in elem:
        tag = child.tag.split("}")[-1]
        if tag == "Folder":
            folder_name_el = child.find("ns0:name", ns)
            folder_name = folder_name_el.text.strip() if folder_name_el is not None else folder
            placemarks += extract_placemarks_verbose(child, ns, folder_name)
        elif tag == "Placemark":
            name_el = child.find("ns0:name", ns)
            coord_el = child.find(".//ns0:coordinates", ns)
            if name_el is not None and coord_el is not None:
                name = name_el.text.strip()
                coord_parts = coord_el.text.strip().split(",")
                lon = float(coord_parts[0].strip())
                lat = float(coord_parts[1].strip())
                placemarks.append({
                    "folder": folder,
                    "name": name,
                    "lat": lat,
                    "lon": lon
                })
    return placemarks

# Ganti ini dengan path file KMZ kamu
kmz_path = "SRI MERANTI RW 16 PEKANBARU.kmz"

# Proses dan tampilkan hasil
root = extract_kml_from_kmz(kmz_path)
ns = {'ns0': 'http://www.opengis.net/kml/2.2'}
placemarks_detected = extract_placemarks_verbose(root, ns)

print("Total titik ditemukan:", len(placemarks_detected))
print("Contoh 10 data pertama:")
for p in placemarks_detected[:10]:
    print(f"[{p['folder']}] {p['name']} - ({p['lat']}, {p['lon']})")
