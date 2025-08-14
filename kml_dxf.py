import os
import zipfile
import pandas as pd
import geopandas as gpd
import streamlit as st
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, box
from shapely.ops import unary_union, linemerge, snap, polygonize
import osmnx as ox
from shapely import wkt

TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10

# =====================
# LAYER / WIDTH
# =====================
def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']:
        return 'HIGHWAYS', 10
    elif hwy in ['secondary', 'tertiary']:
        return 'MAJOR_ROADS', 10
    elif hwy in ['residential', 'unclassified', 'service']:
        return 'MINOR_ROADS', 10
    elif hwy in ['footway', 'path', 'cycleway']:
        return 'PATHS', 10
    return 'OTHER', DEFAULT_WIDTH

# =====================
# HELPER CRS
# =====================
def ensure_wgs84_polygon(polygon, crs_hint):
    src_crs = crs_hint if crs_hint is not None else "EPSG:4326"
    poly_ll = gpd.GeoSeries([polygon], crs=src_crs).to_crs("EPSG:4326").iloc[0]
    return poly_ll

# =====================
# EXTRACT POLYGON KML
# =====================
def extract_polygon_from_kml(kml_path):
    gdf = gpd.read_file(kml_path)
    polygons = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polygons.empty:
        raise Exception("No Polygon found in KML")
    poly_union = unary_union(polygons.geometry)
    # simplify geometry to prevent linemerge errors
    poly_union = poly_union.simplify(0.00001, preserve_topology=True)
    return poly_union, (polygons.crs if polygons.crs is not None else "EPSG:4326")

# =====================
# EXTRACT KMZ → KML
# =====================
def extract_kmz_to_kml(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kml_files = [f for f in kmz_file.namelist() if f.lower().endswith('.kml')]
        if not kml_files:
            raise Exception("❌ Tidak ada file .KML di dalam KMZ")
        kmz_file.extract(kml_files[0], extract_dir)
        return os.path.join(extract_dir, kml_files[0])

# =====================
# OSM ROADS
# =====================
@st.cache_data
def get_osm_roads(polygon, polygon_crs):
    poly_ll = ensure_wgs84_polygon(polygon, polygon_crs)
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(poly_ll, tags=tags)
    except Exception:
        return gpd.GeoDataFrame(geometry=[], crs="EPSG:4326")
    if roads.empty:
        return roads
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    if roads.empty:
        return roads
    roads = roads.explode(index_parts=False)
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    roads = roads.clip(poly_ll)
    roads["geometry"] = roads["geometry"].apply(lambda g: snap(g, g, tolerance=0.0001))
    roads = roads.reset_index(drop=True)
    # simplify lines to reduce processing
    roads["geometry"] = roads["geometry"].apply(lambda g: g.simplify(0.5, preserve_topology=True))
    return roads

# =====================
# OSM BUILDINGS
# =====================
@st.cache_data
def get_osm_buildings(polygon, polygon_crs):
    poly_ll = ensure_wgs84_polygon(polygon, polygon_crs)
    tags = {"building": True}
    try:
        buildings = ox.features_from_polygon(poly_ll, tags=tags)
    except Exception:
        return gpd.GeoDataFrame(geometry=[], crs="EPSG:4326")
    if buildings.empty:
        return buildings
    buildings = buildings[buildings.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if buildings.empty:
        return buildings
    buildings = buildings.explode(index_parts=False)
    buildings = buildings[~buildings.geometry.is_empty & buildings.geometry.notnull()]
    buildings = buildings.clip(poly_ll)
    buildings = buildings.reset_index(drop=True)
    # simplify building polygons
    buildings["geometry"] = buildings["geometry"].apply(lambda g: g.simplify(0.5, preserve_topology=True))
    return buildings

# =====================
# GOOGLE OPEN BUILDINGS
# =====================
def load_google_buildings(csv_gz_path, polygon=None, polygon_crs="EPSG:4326"):
    df = pd.read_csv(csv_gz_path, compression='gzip')
    if 'x' in df.columns and 'y' in df.columns:
        df = df.rename(columns={'x':'longitude', 'y':'latitude'})
    elif 'longitude' not in df.columns or 'latitude' not in df.columns:
        raise Exception("CSV Google Open Buildings tidak ada kolom longitude/latitude")
    gdf = gpd.GeoDataFrame(
        df,
        geometry=gpd.points_from_xy(df.longitude, df.latitude),
        crs="EPSG:4326"
    )
    if polygon is not None:
        poly_ll = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs("EPSG:4326").iloc[0]
        gdf = gdf[gdf.geometry.within(poly_ll)]
    # Ubah point ke kotak kecil (5m)
    gdf['geometry'] = gdf['geometry'].buffer(5)
    return gdf

# =====================
# LOAD CSV GEO
# =====================
def load_csv_to_gdf(csv_path, geom_col="geometry", crs="EPSG:4326"):
    df = pd.read_csv(csv_path)
    if geom_col not in df.columns:
        raise Exception(f"CSV tidak ada kolom geometry bernama {geom_col}")
    gdf = gpd.GeoDataFrame(df, geometry=df[geom_col].apply(wkt.loads), crs=crs)
    return gdf

# =====================
# STRIP Z
# =====================
def strip_z(geom):
    if geom.geom_type == "LineString" and hasattr(geom, "has_z") and geom.has_z:
        return LineString([(x, y) for x, y, *_ in geom.coords])
    elif geom.geom_type == "MultiLineString":
        new_geoms = []
        for line in geom.geoms:
            if hasattr(line, "has_z") and line.has_z:
                new_geoms.append(LineString([(x, y) for x, y, *_ in line.coords]))
            else:
                new_geoms.append(line)
        return MultiLineString(new_geoms)
    return geom

# =====================
# ADD BUILDINGS KE DXF
# =====================
def add_buildings_to_dxf(msp, buildings, min_x, min_y):
    for geom in buildings.geometry:
        if geom.is_empty:
            continue
        minx, miny, maxx, maxy = geom.bounds
        rect = box(minx, miny, maxx, maxy)
        coords = [(x - min_x, y - min_y) for x, y in rect.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "BUILDINGS"})

# =====================
# EXPORT DXF
# =====================
def export_to_dxf(gdf_roads, dxf_path, polygon=None, polygon_crs=None, buildings=None):
    doc = ezdxf.new()
    msp = doc.modelspace()
    all_buffers = []

    for _, row in gdf_roads.iterrows():
        geom = strip_z(row.geometry)
        _, width = classify_layer(str(row.get("highway", "")))
        if geom.is_empty or not geom.is_valid:
            continue

        # SAFE MERGE & SIMPLIFY
        if isinstance(geom, (LineString, MultiLineString)):
            merged = linemerge(geom) if isinstance(geom, MultiLineString) else geom
            merged = merged.simplify(0.5, preserve_topology=True)
            buffered = merged.buffer(width / 2, resolution=8, join_style=2)
            all_buffers.append(buffered)
        elif isinstance(geom, (Polygon, MultiPolygon)):
            all_buffers.append(geom.boundary)

    if all_buffers:
        all_union = unary_union(all_buffers)
        outlines = list(polygonize(all_union.boundary))
        bounds = [(pt[0], pt[1]) for geom in outlines for pt in geom.exterior.coords]
        min_x = min(x for x, y in bounds)
        min_y = min(y for x, y in bounds)
        for outline in outlines:
            coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})
    else:
        min_x = min_y = 0

    # BOUNDARY
    if polygon is not None and polygon_crs is not None:
        poly = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs(TARGET_EPSG).iloc[0]
        if poly.geom_type == 'Polygon':
            coords = [(pt[0] - min_x, pt[1] - min_y) for pt in poly.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})
        elif poly.geom_type == 'MultiPolygon':
            for p in poly.geoms:
                coords = [(pt[0] - min_x, pt[1] - min_y) for pt in p.exterior.coords]
                msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})

    # BUILDINGS
    if buildings is not None and not buildings.empty:
        add_buildings_to_dxf(msp, buildings, min_x, min_y)

    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)

# =====================
# PROSES UTAMA KML → DXF
# =====================
def process_kml_to_dxf(kml_path, output_dir, google_building_csv=None):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml(kml_path)
    roads = get_osm_roads(polygon, polygon_crs)
    buildings = get_osm_buildings(polygon, polygon_crs)
    if buildings.empty and google_building_csv is not None:
        buildings = load_google_buildings(google_building_csv, polygon, polygon_crs)
    geojson_path = os.path.join(output_dir, "roadmap_osm.geojson")
    dxf_path = os.path.join(output_dir, "roadmap_osm.dxf")
    if not roads.empty or not buildings.empty:
        roads_utm = roads.to_crs(TARGET_EPSG) if not roads.empty else gpd.GeoDataFrame(geometry=[], crs=TARGET_EPSG)
        buildings_utm = buildings.to_crs(TARGET_EPSG) if not buildings.empty else None
        if not roads_utm.empty:
            roads_utm.to_file(geojson_path, driver="GeoJSON")
        export_to_dxf(roads_utm, dxf_path, polygon, polygon_crs, buildings_utm)
        return dxf_path, geojson_path, True
    else:
        raise Exception("Tidak ada jalan atau bangunan ditemukan di dalam area polygon.")

# =====================
# STREAMLIT
# =====================
def run_kml_dxf():
    st.title("🌍 KML/KMZ/CSV → Jalan & Kotak Bangunan dari Boundary")
    st.markdown("""
<h2>👋 Hai, <span style='color:#0A84FF'>bro</span></h2>
✅ <span style='font-weight:bold;'>CATATAN PENTING :</span><br><br>
1️⃣ Boleh upload `.KML` atau `.KMZ`.<br>
2️⃣ Atau upload CSV (WKT geometry).<br>
3️⃣ Sistem pastikan polygon ke EPSG:4326 sebelum query OSM.<br>
4️⃣ Bangunan digambar sebagai kotak (bounding box) di layer <code>BUILDINGS</code>.<br>
5️⃣ Bisa juga upload Google Open Buildings CSV (.csv.gz) jika OSM kosong.<br><br>
""", unsafe_allow_html=True)

    kml_file = st.file_uploader("Upload file .KML atau .KMZ", type=["kml", "kmz"])
    google_csv = st.file_uploader("Opsional: Upload Google Open Buildings CSV (.csv.gz)", type=["csv.gz"])
    csv_file = st.file_uploader("Atau upload CSV langsung (WKT geometry)", type=["csv"])

    # ==== KML / KMZ ====
    if kml_file:
        if st.button("Mulai Proses KML/KMZ"):
            with st.spinner("💫 Memproses file..."):
                try:
                    temp_input = f"/tmp/{kml_file.name}"
                    with open(temp_input, "wb") as f:
                        f.write(kml_file.read())
                    if temp_input.lower().endswith(".kmz"):
                        temp_input = extract_kmz_to_kml(temp_input, "/tmp")
                    output_dir = "/tmp/output"
                    dxf_path, geojson_path, ok = process_kml_to_dxf(temp_input, output_dir, google_building_csv=google_csv)
                    if ok:
                        st.success("✅ Berhasil diekspor ke DXF!")
                        with open(dxf_path, "rb") as f:
                            st.download_button("⬇️ Download DXF (UTM 60)", data=f, file_name="roadmap_osm.dxf")
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {e}")

    # ==== CSV ====
    if csv_file:
        if st.button("Mulai Proses CSV"):
            with st.spinner("💫 Memproses CSV..."):
                try:
                    temp_csv = f"/tmp/{csv_file.name}"
                    with open(temp_csv, "wb") as f:
                        f.write(csv_file.read())
                    gdf = load_csv_to_gdf(temp_csv)
                    output_dir = "/tmp/output_csv"
                    os.makedirs(output_dir, exist_ok=True)
                    dxf_path = os.path.join(output_dir, "from_csv.dxf")
                    export_to_dxf(gdf_roads=gdf, dxf_path=dxf_path, buildings=gdf)
                    st.success("✅ CSV berhasil diekspor ke DXF!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF", data=f, file_name="from_csv.dxf")
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {e}")

# =====================
# RUN APP
# =====================
if __name__ == "__main__":
    run_kml_dxf()
