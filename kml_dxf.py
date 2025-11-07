# kml_converter_fixed_v4.py
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

# ---------------------- KLASIFIKASI ----------------------
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

# ---------------------- OSM ----------------------
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
        roads = roads.clip(polygon.buffer(0.00005))
    except Exception:
        roads = roads[roads.intersects(polygon.buffer(0.00005))]
    return roads.reset_index(drop=True)

# ---------------------- HERE API ----------------------
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
    return gpd.GeoDataFrame(features, crs="EPSG:4326")

# ---------------------- BUFFER DALAM METER ----------------------
def buffer_in_meters(geom, meters):
    """Buffer dalam satuan meter dengan transformasi CRS otomatis (UTM)"""
    lon = geom.centroid.x
    zone = int((lon + 180) / 6) + 1
    epsg_utm = 32700 + zone  # UTM Selatan
    geom_utm = gpd.GeoSeries([geom], crs="EPSG:4326").to_crs(epsg_utm)[0]
    buffered_utm = geom_utm.buffer(meters, resolution=8, join_style=2)
    return gpd.GeoSeries([buffered_utm], crs=epsg_utm).to_crs("EPSG:4326")[0]

# ---------------------- EXPORT KML ----------------------
def export_to_kml_simple(gdf, polygon, output_path):
    kml = simplekml.Kml()

    # Tambahkan boundary
    if isinstance(polygon, (Polygon, MultiPolygon)):
        for i, part in enumerate(getattr(polygon, "geoms", [polygon])):
            outer = [(x, y) for x, y in part.exterior.coords]
            poly = kml.newpolygon(name=f"Boundary_{i+1}")
            poly.outerboundaryis = outer
            poly.style.polystyle.color = simplekml.Color.changealphaint(40, simplekml.Color.blue)

    # Tambahkan jalan
    for idx, row in gdf.iterrows():
        geom = row.geometry
        if geom is None or geom.is_empty:
            continue

        hwy = str(row.get("highway", row.get("type", "")))
        layer = classify_layer(hwy)
        style = COLOR_MAP[layer]

        if isinstance(geom, (LineString, MultiLineString)):
            buffered = buffer_in_meters(geom, style["width"])  # buffer kiri-kanan
            if isinstance(buffered, Polygon):
                outer = [(x, y) for x, y in buffered.exterior.coords]
                pol = kml.newpolygon(name=layer)
                pol.outerboundaryis = outer
                pol.style.polystyle.color = simplekml.Color.changealphaint(120, style["color"])
                pol.style.linestyle.color = style["color"]

    kml.save(output_path)

# ---------------------- PROSES ----------------------
def process_kml_to_kml(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, _ = extract_polygon_from_kml_or_kmz(kml_path)
    roads = get_osm_roads(polygon)
    if roads.empty:
        roads = get_here_roads(polygon)
    if roads.empty:
        raise Exception("❌ Tidak ada jalan ditemukan (OSM & HERE kosong).")
    roads = roads.to_crs("EPSG:4326")
    kml_output = os.path.join(output_dir, "roadmap.kml")
    export_to_kml_simple(roads, polygon, kml_output)
    return kml_output, True

# ---------------------- STREAMLIT APP ----------------------
def run_kml_dxf():
    st.title("🌍 KML/KMZ → KML Jalan (Presisi Mirip DXF)")
    st.markdown("Upload area POLYGON (KML/KMZ) → hasil KML jalan dua sisi rapi (buffer meter).")
    kml_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())
                output_dir = "/tmp/output"
                kml_path, ok = process_kml_to_kml(temp_input, output_dir)
                if ok:
                    st.success("✅ Berhasil diekspor ke KML (Presisi tinggi)!")
                    with open(kml_path, "rb") as f:
                        st.download_button("⬇️ Download KML", data=f, file_name="roadmap_buffered.kml")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")
