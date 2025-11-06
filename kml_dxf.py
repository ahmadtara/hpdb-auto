import os
import zipfile
import requests
from fastkml import kml
import geopandas as gpd
import streamlit as st
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, shape
from shapely.ops import unary_union, linemerge, polygonize
import osmnx as ox

# ---------------------------- #
# KONFIGURASI DASAR
# ---------------------------- #
TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10
HERE_API_KEY = "k1mDEfR1Q3A_MtLqkxLrhbDcS-1oC4r7WzlgPrcv4Rk"  # Ganti API Key HERE Maps kamu


# ---------------------------- #
# KLASIFIKASI LAYER
# ---------------------------- #
def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']:
        return 'HIGHWAYS', 14
    elif hwy in ['secondary', 'tertiary']:
        return 'MAJOR_ROADS', 10
    elif hwy in ['residential', 'unclassified', 'service']:
        return 'MINOR_ROADS', 8
    elif hwy in ['footway', 'path', 'cycleway']:
        return 'PATHS', 4
    return 'OTHER', DEFAULT_WIDTH


# ---------------------------- #
# EKSTRAK POLYGON DARI KML/KMZ
# ---------------------------- #
def extract_polygon_from_kml_or_kmz(path):
    if path.endswith(".kmz"):
        with zipfile.ZipFile(path, "r") as z:
            for name in z.namelist():
                if name.endswith(".kml"):
                    z.extract(name, "/tmp")
                    path = os.path.join("/tmp", name)
                    break
    gdf = gpd.read_file(path)
    polygons = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polygons.empty:
        raise Exception("❌ No Polygon found in KML/KMZ")
    return unary_union(polygons.geometry), polygons.crs


# ---------------------------- #
# AMBIL JALAN DARI OSM
# ---------------------------- #
def get_osm_roads(polygon):
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(polygon, tags=tags)
    except Exception:
        return gpd.GeoDataFrame()
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    roads = roads.explode(index_parts=False)
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    roads = roads.clip(polygon)
    return roads.reset_index(drop=True)


# ---------------------------- #
# AMBIL JALAN DARI HERE API (Fallback)
# ---------------------------- #
def get_here_roads(polygon):
    minx, miny, maxx, maxy = polygon.bounds
    url = (
        f"https://vector.hereapi.com/v2/vectortiles/base/mc"
        f"?apikey={HERE_API_KEY}&bbox={miny},{minx},{maxy},{maxx}&layers=roads"
    )
    resp = requests.get(url)
    if resp.status_code != 200:
        raise Exception(f"HERE API error: {resp.text}")

    data = resp.json()
    features = []
    for f in data.get("features", []):
        geom = shape(f["geometry"])
        if geom.intersects(polygon):
            features.append({"geometry": geom.intersection(polygon), "properties": f.get("properties", {})})
    if not features:
        return gpd.GeoDataFrame()
    return gpd.GeoDataFrame(features, crs="EPSG:4326")


# ---------------------------- #
# HAPUS Z DARI GEOMETRI
# ---------------------------- #
def strip_z(geom):
    if geom.geom_type == "LineString" and geom.has_z:
        return LineString([(x, y) for x, y, *_ in geom.coords])
    elif geom.geom_type == "MultiLineString":
        return MultiLineString([
            LineString([(x, y) for x, y, *_ in line.coords]) if line.has_z else line
            for line in geom.geoms
        ])
    return geom


# ---------------------------- #
# EKSPOR KE DXF
# ---------------------------- #
def export_to_dxf(gdf, dxf_path, polygon=None, polygon_crs=None):
    doc = ezdxf.new()
    msp = doc.modelspace()
    all_buffers = []

    for _, row in gdf.iterrows():
        geom = strip_z(row.geometry)
        hwy = str(row.get("highway", row.get("type", "")))
        layer, width = classify_layer(hwy)
        if geom.is_empty or not geom.is_valid:
            continue

        merged = linemerge(geom) if isinstance(geom, MultiLineString) else geom
        if isinstance(merged, (LineString, MultiLineString)):
            buffered = merged.buffer(width / 2, resolution=8, join_style=2)
            all_buffers.append(buffered)

    if not all_buffers:
        raise Exception("❌ Tidak ada garis valid untuk diekspor.")

    all_union = unary_union(all_buffers)
    outlines = list(polygonize(all_union.boundary))
    min_x = min(pt[0] for geom in outlines for pt in geom.exterior.coords)
    min_y = min(pt[1] for geom in outlines for pt in geom.exterior.coords)

    for outline in outlines:
        coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})

    # Boundary polygon
    if polygon is not None and polygon_crs is not None:
        poly = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs(TARGET_EPSG).iloc[0]
        if poly.geom_type == 'Polygon':
            coords = [(pt[0] - min_x, pt[1] - min_y) for pt in poly.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})
        elif poly.geom_type == 'MultiPolygon':
            for p in poly.geoms:
                coords = [(pt[0] - min_x, pt[1] - min_y) for pt in p.exterior.coords]
                msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})

    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)


# ---------------------------- #
# EKSPOR KE KML
# ---------------------------- #
def export_to_kml(gdf, kml_path, polygon=None, polygon_crs=None):
    from fastkml import kml
    import geopandas as gpd
    from shapely.geometry import MultiLineString

    ns = '{http://www.opengis.net/kml/2.2}'
    k = kml.KML()
    doc = kml.Document(ns, 'docid', 'RoadMap', 'Generated road map')
    k.append(doc)

    # Pastikan CRS ke EPSG:4326
    if gdf.crs is not None and gdf.crs.to_string() != "EPSG:4326":
        gdf = gdf.to_crs("EPSG:4326")

    # Tambah layer jalan
    for _, row in gdf.iterrows():
        geom = strip_z(row.geometry)
        if geom.is_empty or not geom.is_valid:
            continue
        hwy = str(row.get("highway", row.get("type", "")))
        layer, _ = classify_layer(hwy)

        if isinstance(geom, MultiLineString):
            for part in geom.geoms:
                placemark = kml.Placemark(ns, None, f"{layer}", f"Type: {hwy}", geometry=part)
                doc.append(placemark)
        else:
            placemark = kml.Placemark(ns, None, f"{layer}", f"Type: {hwy}", geometry=geom)
            doc.append(placemark)

    # Tambah boundary polygon
    if polygon is not None and polygon_crs is not None:
        try:
            poly = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs("EPSG:4326").iloc[0]
            if poly.geom_type == 'Polygon':
                placemark = kml.Placemark(ns, None, "BOUNDARY", "Boundary polygon", geometry=poly)
                doc.append(placemark)
            elif poly.geom_type == 'MultiPolygon':
                for idx, p in enumerate(poly.geoms):
                    placemark = kml.Placemark(ns, None, f"BOUNDARY_{idx+1}", "Boundary part", geometry=p)
                    doc.append(placemark)
        except Exception as e:
            print(f"⚠️ Boundary polygon skip: {e}")

    with open(kml_path, "w", encoding="utf-8") as f:
        f.write(k.to_string(prettyprint=True))


# ---------------------------- #
# PROSES UTAMA
# ---------------------------- #
def process_kml_to_dxf(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml_or_kmz(kml_path)
    roads = get_osm_roads(polygon)
    if roads.empty:
        roads = get_here_roads(polygon)
    if roads.empty:
        raise Exception("❌ Tidak ada jalan ditemukan (OSM & HERE kosong).")

    geojson_path = os.path.join(output_dir, "roadmap.geojson")
    dxf_path = os.path.join(output_dir, "roadmap.dxf")
    kml_out = os.path.join(output_dir, "roadmap.kml")

    roads_utm = roads.to_crs(TARGET_EPSG)
    roads_utm.to_file(geojson_path, driver="GeoJSON")

    export_to_dxf(roads_utm, dxf_path, polygon=polygon, polygon_crs=polygon_crs)
    export_to_kml(roads, kml_out, polygon=polygon, polygon_crs=polygon_crs)

    return dxf_path, geojson_path, kml_out, True


# ---------------------------- #
# STREAMLIT UI
# ---------------------------- #
def run_kml_dxf():
    st.title("🌍 KML/KMZ → Road Converter (OSM + HERE fallback)")

    kml_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())

                output_dir = "/tmp/output"
                dxf_path, geojson_path, kml_path, ok = process_kml_to_dxf(temp_input, output_dir)
                if ok:
                    st.success("✅ Berhasil diekspor!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF (UTM 60)", data=f, file_name="roadmap.dxf")
                    with open(kml_path, "rb") as f:
                        st.download_button("🌍 Download KML (Google Earth)", data=f, file_name="roadmap.kml")
                    with open(geojson_path, "rb") as f:
                        st.download_button("📁 Download GeoJSON", data=f, file_name="roadmap.geojson")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")
