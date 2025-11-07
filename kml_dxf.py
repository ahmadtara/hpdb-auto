# ======================================================
# kml_converter_fixed_v2.py
# ======================================================
import os
import zipfile
import requests
import geopandas as gpd
import streamlit as st
import osmnx as ox
from shapely.geometry import shape, LineString, MultiLineString, Polygon, MultiPolygon
from shapely.ops import unary_union
import simplekml

HERE_API_KEY = "jGCMpa59MeURAH39Vzk94kutVqC3vl714_ZvcHodX14"

# ---------------------- KLASIFIKASI JALAN ----------------------
def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']:
        return 'HIGHWAYS'
    elif hwy in ['secondary', 'tertiary']:
        return 'MAJOR_ROADS'
    elif hwy in ['residential', 'unclassified', 'service']:
        return 'MINOR_ROADS'
    elif hwy in ['footway', 'path', 'cycleway']:
        return 'PATHS'
    return 'OTHER'

COLOR_MAP = {
    "HIGHWAYS": {"color": simplekml.Color.yellow, "width": 4},
    "MAJOR_ROADS": {"color": simplekml.Color.orange, "width": 3},
    "MINOR_ROADS": {"color": simplekml.Color.white, "width": 2},
    "PATHS": {"color": simplekml.Color.green, "width": 1.5},
    "OTHER": {"color": simplekml.Color.gray, "width": 1}
}

# ---------------------- EKSTRAK POLIGON ----------------------
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
        raise Exception("❌ Tidak ada Polygon di KML/KMZ.")
    if polygons.crs is not None:
        polygons = polygons.to_crs("EPSG:4326")
    return unary_union(polygons.geometry), polygons.crs or "EPSG:4326"

# ---------------------- OSM DATA ----------------------
def get_osm_roads(polygon):
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(polygon, tags=tags)
    except Exception:
        return gpd.GeoDataFrame()
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    try:
        roads = roads.explode(index_parts=False)
    except TypeError:
        roads = roads.explode()
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    if roads.crs is not None:
        roads = roads.to_crs("EPSG:4326")
    try:
        roads = roads.clip(polygon)
    except Exception:
        roads = roads[roads.intersects(polygon)]
    return roads.reset_index(drop=True)

# ---------------------- HERE API FALLBACK ----------------------
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
            inter = geom.intersection(polygon)
            features.append({"geometry": inter, "highway": f.get("properties", {}).get("class", "")})
    if not features:
        return gpd.GeoDataFrame()
    gdf = gpd.GeoDataFrame(features, crs="EPSG:4326")
    return gdf

# ---------------------- EXPORT KML ----------------------
def export_to_kml_simple(gdf, polygon, output_path, include_geojson=False):
    """
    Buat hasil jalan tebal (buffer kiri-kanan) dan rapi seperti DXF.
    """
    WIDTH_MAP = {
        "HIGHWAYS": 14,
        "MAJOR_ROADS": 10,
        "MINOR_ROADS": 8,
        "PATHS": 4,
        "OTHER": 3
    }

    kml_obj = simplekml.Kml()

    # Tambahkan boundary area
    if polygon is not None:
        if isinstance(polygon, (Polygon, MultiPolygon)):
            for i, p in enumerate(getattr(polygon, "geoms", [polygon])):
                outer = [(x, y) for x, y in p.exterior.coords]
                poly = kml_obj.newpolygon(name=f"Boundary_{i+1}")
                poly.outerboundaryis = outer
                poly.style.polystyle.color = simplekml.Color.changealphaint(60, simplekml.Color.blue)

    # Gabungkan dan buffer per tipe jalan
    for layer in ["HIGHWAYS", "MAJOR_ROADS", "MINOR_ROADS", "PATHS", "OTHER"]:
        subset = gdf[gdf["highway"].apply(lambda h: classify_layer(str(h)) == layer)]
        if subset.empty:
            continue
        try:
            merged = unary_union(subset.geometry)
            buffered = merged.buffer(WIDTH_MAP[layer] / 100000.0, resolution=8, join_style=2)
            color = COLOR_MAP[layer]["color"]
            if isinstance(buffered, (Polygon, MultiPolygon)):
                for poly in getattr(buffered, "geoms", [buffered]):
                    outer = [(x, y) for x, y in poly.exterior.coords]
                    p = kml_obj.newpolygon(name=layer)
                    p.outerboundaryis = outer
                    p.style.polystyle.color = simplekml.Color.changealphaint(140, color)
                    p.style.linestyle.color = color
                    p.style.linestyle.width = 1
        except Exception:
            continue

    kml_obj.save(output_path)

    if include_geojson:
        geojson_path = os.path.splitext(output_path)[0] + ".geojson"
        try:
            gdf.to_file(geojson_path, driver="GeoJSON")
        except Exception:
            pass

# ---------------------- PROSES UTAMA ----------------------
def process_kml_to_kml(kml_path, output_dir, include_geojson=False):
    os.makedirs(output_dir, exist_ok=True)
    polygon, _ = extract_polygon_from_kml_or_kmz(kml_path)
    roads = get_osm_roads(polygon)
    if roads.empty:
        roads = get_here_roads(polygon)
    if roads.empty:
        raise Exception("❌ Tidak ada jalan ditemukan (OSM & HERE kosong).")
    roads = roads.to_crs("EPSG:4326")
    kml_output = os.path.join(output_dir, "roadmap.kml")
    export_to_kml_simple(roads, polygon, kml_output, include_geojson)
    return kml_output, True

# ---------------------- STREAMLIT APP ----------------------
def run_kml_dxf():
    st.title("🌍 KML/KMZ → KML Road Converter (Mirip DXF)")
    st.markdown("Upload file KML/KMZ (POLYGON area) untuk diambil jaringan jalan OSM.")
    include_geojson = st.checkbox("Simpan juga GeoJSON (debug opsional)", value=True)
    kml_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())
                output_dir = "/tmp/output"
                kml_path, ok = process_kml_to_kml(temp_input, output_dir, include_geojson)
                if ok:
                    st.success("✅ Berhasil diekspor ke KML!")
                    with open(kml_path, "rb") as f:
                        st.download_button("⬇️ Download KML", data=f, file_name="roadmap.kml")
                    geojson_path = os.path.splitext(kml_path)[0] + ".geojson"
                    if include_geojson and os.path.exists(geojson_path):
                        with open(geojson_path, "rb") as gf:
                            st.download_button("⬇️ Download GeoJSON", data=gf, file_name="roadmap.geojson")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")
