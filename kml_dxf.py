# kml_converter_fixed.py
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

# warna + width per kategori untuk simplekml
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
    # pastikan geometry dalam EPSG:4326
    if polygons.crs is not None:
        polygons = polygons.to_crs("EPSG:4326")
    # unary_union untuk gabungkan multipolygon/parts
    return unary_union(polygons.geometry), polygons.crs or "EPSG:4326"

# ---------------------- AMBIL DATA JALAN DARI OSM ----------------------
def get_osm_roads(polygon):
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(polygon, tags=tags)
    except Exception:
        return gpd.GeoDataFrame()
    # filter hanya garis
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    # flatten multi-part rows (compatibility across geopandas versions)
    try:
        roads = roads.explode(index_parts=False)
    except TypeError:
        roads = roads.explode()
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    # pastikan CRS 4326
    if roads.crs is not None:
        roads = roads.to_crs("EPSG:4326")
    # clip atau fallback intersects
    try:
        roads = roads.clip(polygon)
    except Exception:
        roads = roads[roads.intersects(polygon)]
    return roads.reset_index(drop=True)

# ---------------------- OPSI FALLBACK: HERE API ----------------------
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

# ---------------------- EXPORT TO KML (simplekml) ----------------------
def export_to_kml_simple(gdf, polygon, output_path, include_geojson=False):
    """
    Export roads GeoDataFrame (EPSG:4326) dan polygon boundary ke KML (pakai buffer area biar mirip DXF)
    """
    from shapely.ops import unary_union
    from shapely.geometry import MultiPolygon

    # Lebar relatif antar tipe jalan
    WIDTH_MAP = {
        "HIGHWAYS": 14,
        "MAJOR_ROADS": 10,
        "MINOR_ROADS": 8,
        "PATHS": 4,
        "OTHER": 3
    }

    kml_obj = simplekml.Kml()

    # Boundary
    if polygon is not None:
        if isinstance(polygon, Polygon):
            outer = [(coord[0], coord[1]) for coord in polygon.exterior.coords]
            pol = kml_obj.newpolygon(name="Boundary")
            pol.outerboundaryis = outer
            pol.style.polystyle.color = simplekml.Color.changealphaint(60, simplekml.Color.blue)
        elif isinstance(polygon, MultiPolygon):
            for i, p in enumerate(polygon.geoms):
                outer = [(coord[0], coord[1]) for coord in p.exterior.coords]
                pol = kml_obj.newpolygon(name=f"Boundary_{i+1}")
                pol.outerboundaryis = outer
                pol.style.polystyle.color = simplekml.Color.changealphaint(60, simplekml.Color.blue)

    # Buffer tiap jalan (seperti DXF)
    buffered_polygons = []
    for _, row in gdf.iterrows():
        geom = row.geometry
        if geom is None or geom.is_empty or not geom.is_valid:
            continue
        hwy = str(row.get("highway", row.get("type", "")))
        layer = classify_layer(hwy)
        width = WIDTH_MAP.get(layer, 6)
        try:
            buffered = geom.buffer(width / 100000.0, resolution=6, join_style=2)  # width kecil (deg)
            buffered_polygons.append((buffered, layer))
        except Exception:
            continue

    # Gabungkan & ekspor ke KML sebagai Polygon
    for geom, layer in buffered_polygons:
        color = COLOR_MAP.get(layer, COLOR_MAP["OTHER"])["color"]
        if isinstance(geom, (Polygon, MultiPolygon)):
            for poly in getattr(geom, "geoms", [geom]):
                outer = [(coord[0], coord[1]) for coord in poly.exterior.coords]
                pol = kml_obj.newpolygon(name=layer)
                pol.outerboundaryis = outer
                pol.style.polystyle.color = simplekml.Color.changealphaint(120, color)
                pol.style.linestyle.color = color
                pol.style.linestyle.width = 1

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
    polygon, polygon_crs = extract_polygon_from_kml_or_kmz(kml_path)
    roads = get_osm_roads(polygon)
    if roads.empty:
        roads = get_here_roads(polygon)
    if roads.empty:
        raise Exception("❌ Tidak ada jalan ditemukan (OSM & HERE kosong).")
    # pastikan roads CRS 4326
    if roads.crs is not None:
        roads = roads.to_crs("EPSG:4326")

    kml_output = os.path.join(output_dir, "roadmap.kml")
    export_to_kml_simple(roads, polygon, kml_output, include_geojson=include_geojson)
    return kml_output, True

# ---------------------- STREAMLIT APP ----------------------
def run_kml_dxf():
    st.title("🌍 KML/KMZ → KML Road Converter (OSM + HERE Fallback)")
    st.markdown("Upload file KML/KMZ yang berisi POLYGON area yang akan diambil jalan-jalannya.")
    include_geojson = st.checkbox("Simpan juga GeoJSON (opsional, untuk debugging)", value=True)
    kml_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())
                output_dir = "/tmp/output"
                kml_path, ok = process_kml_to_kml(temp_input, output_dir, include_geojson=include_geojson)
                if ok:
                    st.success("✅ Berhasil diekspor ke KML!")
                    with open(kml_path, "rb") as f:
                        st.download_button("⬇️ Download KML", data=f, file_name="roadmap.kml")
                    # show geojson link if exists
                    geojson_path = os.path.splitext(kml_path)[0] + ".geojson"
                    if include_geojson and os.path.exists(geojson_path):
                        with open(geojson_path, "rb") as gf:
                            st.download_button("⬇️ Download GeoJSON", data=gf, file_name="roadmap.geojson")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")


