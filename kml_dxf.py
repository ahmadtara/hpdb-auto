import os
import zipfile
import pandas as pd
import geopandas as gpd
import streamlit as st
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, box
from shapely.ops import unary_union, linemerge, snap, polygonize
import osmnx as ox
import gdown

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
    return roads

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
# GOOGLE OPEN BUILDINGS LANGSUNG DARI DRIVE
# =====================
def load_google_buildings_from_drive(google_drive_link, polygon=None, polygon_crs="EPSG:4326"):
    # download CSV sementara dari Google Drive
    file_id = google_drive_link.split("/d/")[1].split("/")[0]
    url = f"https://drive.google.com/uc?id={file_id}"
    temp_path = "/tmp/31d_buildings.csv"
    gdown.download(url, temp_path, quiet=False)
    
    df = pd.read_csv(temp_path)
    
    # pastikan kolom latitude/longitude
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
        merged = geom if isinstance(geom, LineString) else linemerge(geom)
        if isinstance(merged, (LineString, MultiLineString)):
            buffered = merged.buffer(width / 2, resolution=8, join_style=2)
            all_buffers.append(buffered)
    if not all_buffers:
        raise Exception("❌ Tidak ada garis valid untuk diekspor.")
    all_union = unary_union(all_buffers)
    outlines = list(polygonize(all_union.boundary))
    if not outlines:
        raise Exception("❌ Polygonize gagal menghasilkan outline.")
    bounds = [(pt[0], pt[1]) for geom in outlines for pt in geom.exterior.coords]
    min_x = min(x for x, y in bounds)
    min_y = min(y for x, y in bounds)
    # ROADS
    for outline in outlines:
        coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})
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
# PROSES UTAMA
# =====================
def process_kml_to_dxf(kml_path, output_dir, google_drive_link=None):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml(kml_path)
    roads = get_osm_roads(polygon, polygon_crs)
    buildings = None
    if google_drive_link:
        buildings = load_google_buildings_from_drive(google_drive_link, polygon, polygon_crs)
    geojson_path = os.path.join(output_dir, "roadmap_osm.geojson")
    dxf_path = os.path.join(output_dir, "roadmap_osm.dxf")
    if (roads.empty if False else True) or (buildings is not None and not buildings.empty):
        roads_utm = roads.to_crs(TARGET_EPSG) if not roads.empty else gpd.GeoDataFrame(geometry=[], crs=TARGET_EPSG)
        buildings_utm = buildings.to_crs(TARGET_EPSG) if buildings is not None else None
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
    st.title("🌍 KML/KMZ → Jalan & Kotak Bangunan Pekanbaru")
    st.markdown("""
<h2>👋 Hai, <span style='color:#0A84FF'>bro</span></h2>
✅ <span style='font-weight:bold;'>CATATAN PENTING :</span><br><br>
1️⃣ Boleh upload `.KML` atau `.KMZ`.<br>
2️⃣ Sistem filter area Pekanbaru berdasarkan polygon boundary.<br>
3️⃣ Bangunan digambar sebagai kotak (bounding box) di layer <code>BUILDINGS</code>.<br>
4️⃣ Bisa juga ambil Google Open Buildings CSV langsung via Google Drive link.<br><br>
""", unsafe_allow_html=True)

    kml_file = st.file_uploader("Upload file .KML atau .KMZ", type=["kml", "kmz"])
    google_drive_link = st.text_input("Opsional: Masukkan link Google Drive CSV 31d_buildings.csv")

    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())
                if temp_input.lower().endswith(".kmz"):
                    temp_input = extract_kmz_to_kml(temp_input, "/tmp")
                output_dir = "/tmp/output"
                dxf_path, geojson_path, ok = process_kml_to_dxf(temp_input, output_dir, google_drive_link=google_drive_link)
                if ok:
                    st.success("✅ Berhasil diekspor ke DXF!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF (UTM 60)", data=f, file_name="roadmap_osm.dxf")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

# =====================
# RUN APP
# =====================
if __name__ == "__main__":
    run_kml_dxf()
