import os
import zipfile
import requests
import geopandas as gpd
import osmnx as ox
from shapely.geometry import shape
from shapely.ops import unary_union
import streamlit as st
from fastkml import kml as fkml

HERE_API_KEY = "jGCMpa59MeURAH39Vzk94kutVqC3vl714_ZvcHodX14"


# ----------------------------
# Baca polygon dari KML / KMZ
# ----------------------------
def extract_polygon(path):
    if path.endswith(".kmz"):
        with zipfile.ZipFile(path, "r") as z:
            for name in z.namelist():
                if name.endswith(".kml"):
                    z.extract(name, "/tmp")
                    path = os.path.join("/tmp", name)
                    break
    gdf = gpd.read_file(path)
    polys = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polys.empty:
        raise Exception("❌ Tidak ada polygon di file ini.")
    return unary_union(polys.geometry)


# ----------------------------
# Ambil jalan dari OSM
# ----------------------------
def get_osm_roads(polygon):
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(polygon, tags=tags)
    except Exception:
        return gpd.GeoDataFrame()

    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    roads = roads.explode(index_parts=False).reset_index(drop=True)
    roads = roads.clip(polygon)
    return roads


# ----------------------------
# Ambil jalan dari HERE (fallback)
# ----------------------------
def get_here_roads(polygon):
    minx, miny, maxx, maxy = polygon.bounds
    url = (
        f"https://vector.hereapi.com/v2/vectortiles/base/mc"
        f"?apikey={HERE_API_KEY}&bbox={miny},{minx},{maxy},{maxx}&layers=roads"
    )
    r = requests.get(url)
    if r.status_code != 200:
        return gpd.GeoDataFrame()
    data = r.json()
    feats = []
    for f in data.get("features", []):
        geom = shape(f["geometry"])
        if geom.intersects(polygon):
            feats.append({"geometry": geom.intersection(polygon)})
    if not feats:
        return gpd.GeoDataFrame()
    return gpd.GeoDataFrame(feats, crs="EPSG:4326")


# ----------------------------
# Simpan hasil ke file KML (garis saja)
# ----------------------------
def export_lines_to_kml(gdf, out_path):
    ns = "{http://www.opengis.net/kml/2.2}"
    k = fkml.KML()
    doc = fkml.Document(ns, "1", "RoadLines", "Road Lines Only")
    k.append(doc)
    folder = fkml.Folder(ns, "f1", "Roads", "Extracted lines only")
    doc.append(folder)

    if gdf.crs and gdf.crs.to_string() != "EPSG:4326":
        gdf = gdf.to_crs("EPSG:4326")

    for i, row in gdf.iterrows():
        geom = row.geometry
        if geom is None:
            continue
        placemark = fkml.Placemark(ns, str(i), "road", "", geometry=geom)
        folder.append(placemark)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(k.to_string(prettyprint=True))


# ----------------------------
# MAIN PROCESS
# ----------------------------
def process_kml(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon = extract_polygon(kml_path)
    roads = get_osm_roads(polygon)
    if roads.empty:
        roads = get_here_roads(polygon)
    if roads.empty:
        raise Exception("❌ Tidak ada jalan ditemukan di area ini.")
    out_kml = os.path.join(output_dir, "roads_line.kml")
    export_lines_to_kml(roads, out_kml)
    return out_kml


# ----------------------------
# STREAMLIT APP
# ----------------------------
def run_kml_line():
    st.title("📍 KML/KMZ ➜ Road Line Extractor (OSM + HERE)")
    st.markdown("Output: **garis jalan (LineString)** saja seperti hasil DXF.")

    file = st.file_uploader("Upload file .KML / .KMZ (berisi area)", type=["kml", "kmz"])
    if file:
        with st.spinner("🔍 Memproses area dan mengambil garis jalan..."):
            try:
                tmp_in = f"/tmp/{file.name}"
                with open(tmp_in, "wb") as f:
                    f.write(file.read())

                output_dir = "/tmp/output"
                kml_out = process_kml(tmp_in, output_dir)
                st.success("✅ Selesai! File garis jalan siap diunduh.")
                with open(kml_out, "rb") as f:
                    st.download_button("⬇️ Download Garis Jalan (KML)", f, file_name="roads_line.kml")

            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")


if __name__ == "__main__":
    run_kml_line()
