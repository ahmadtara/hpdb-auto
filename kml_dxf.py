import os
import zipfile
import geopandas as gpd
import streamlit as st
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, box
from shapely.ops import unary_union, linemerge, snap, polygonize
import osmnx as ox

TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10

# =====================
# CLASSIFY ROAD LAYERS
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
# LOAD POLYGON FROM KML
# =====================
def extract_polygon_from_kml(kml_path):
    gdf = gpd.read_file(kml_path)
    polygons = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polygons.empty:
        raise Exception("No Polygon found in KML")
    return unary_union(polygons.geometry), polygons.crs

# =====================
# GET ROADS FROM OSM
# =====================
def get_osm_roads(polygon):
    tags = {"highway": True}
    roads = ox.features_from_polygon(polygon, tags=tags)
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    roads = roads.explode(index_parts=False)
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    roads = roads.clip(polygon)
    roads["geometry"] = roads["geometry"].apply(lambda g: snap(g, g, tolerance=0.0001))
    roads = roads.reset_index(drop=True)
    return roads

# =====================
# GET BUILDINGS FROM OSM
# =====================
def get_osm_buildings(polygon):
    tags = {"building": True}
    buildings = ox.features_from_polygon(polygon, tags=tags)
    buildings = buildings[buildings.geometry.type.isin(["Polygon", "MultiPolygon"])]
    buildings = buildings.explode(index_parts=False)
    buildings = buildings[~buildings.geometry.is_empty & buildings.geometry.notnull()]
    buildings = buildings.clip(polygon)
    buildings = buildings.reset_index(drop=True)
    return buildings

# =====================
# REMOVE Z COORDS
# =====================
def strip_z(geom):
    if geom.geom_type == "LineString" and geom.has_z:
        return LineString([(x, y) for x, y, *_ in geom.coords])
    elif geom.geom_type == "MultiLineString":
        return MultiLineString([
            LineString([(x, y) for x, y, *_ in line.coords]) if line.has_z else line
            for line in geom.geoms
        ])
    return geom

# =====================
# ADD BUILDINGS AS BOXES
# =====================
def add_buildings_to_dxf(msp, buildings, min_x, min_y):
    for geom in buildings.geometry:
        if geom.is_empty:
            continue
        bbox = geom.bounds  # (minx, miny, maxx, maxy)
        rect = box(*bbox)
        coords = [(x - min_x, y - min_y) for x, y in rect.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "BUILDINGS"})

# =====================
# EXPORT TO DXF
# =====================
def export_to_dxf(gdf, dxf_path, polygon=None, polygon_crs=None, buildings=None):
    doc = ezdxf.new()
    msp = doc.modelspace()

    all_buffers = []

    for _, row in gdf.iterrows():
        geom = strip_z(row.geometry)
        hwy = str(row.get("highway", ""))
        layer, width = classify_layer(hwy)

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

    # Tambahkan ROADS
    for outline in outlines:
        coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})

    # Tambahkan BOUNDARY
    if polygon is not None and polygon_crs is not None:
        poly = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs(TARGET_EPSG).iloc[0]
        if poly.geom_type == 'Polygon':
            coords = [(pt[0] - min_x, pt[1] - min_y) for pt in poly.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})
        elif poly.geom_type == 'MultiPolygon':
            for p in poly.geoms:
                coords = [(pt[0] - min_x, pt[1] - min_y) for pt in p.exterior.coords]
                msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})

    # Tambahkan BUILDINGS (kotak bounding box)
    if buildings is not None and not buildings.empty:
        add_buildings_to_dxf(msp, buildings, min_x, min_y)

    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)

# =====================
# MAIN PROCESS
# =====================
def process_kml_to_dxf(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml(kml_path)
    
    roads = get_osm_roads(polygon)
    buildings = get_osm_buildings(polygon)

    geojson_path = os.path.join(output_dir, "roadmap_osm.geojson")
    dxf_path = os.path.join(output_dir, "roadmap_osm.dxf")

    if not roads.empty:
        roads_utm = roads.to_crs(TARGET_EPSG)
        buildings_utm = buildings.to_crs(TARGET_EPSG) if not buildings.empty else None
        roads_utm.to_file(geojson_path, driver="GeoJSON")
        export_to_dxf(roads_utm, dxf_path, polygon=polygon, polygon_crs=polygon_crs, buildings=buildings_utm)
        return dxf_path, geojson_path, True
    else:
        raise Exception("Tidak ada jalan ditemukan di dalam area polygon.")

# =====================
# KMZ EXTRACT
# =====================
def extract_kmz_to_kml(kmz_path, extract_dir):
    """Ekstrak file .kml pertama dari .kmz"""
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kml_files = [f for f in kmz_file.namelist() if f.lower().endswith('.kml')]
        if not kml_files:
            raise Exception("❌ Tidak ada file .KML di dalam KMZ")
        kmz_file.extract(kml_files[0], extract_dir)
        return os.path.join(extract_dir, kml_files[0])

# =====================
# STREAMLIT UI
# =====================
def run_kml_dxf():
    st.title("🌍 KML/KMZ → Buat Jalan & Bangunan dari Boundary Cluster")
    st.markdown("""
<h2>👋 Hai, <span style='color:#0A84FF'>bro</span></h2>
✅ <span style='font-weight:bold;'>CATATAN PENTING :</span><br><br>
1️⃣ Format file boleh `.KML` langsung atau `.KMZ`.<br>
2️⃣ Jika `.KMZ`, sistem akan otomatis ekstrak `.KML` utama di dalamnya.<br>
3️⃣ Jalan akan dibuat otomatis berdasarkan boundary cluster.<br>
4️⃣ Bangunan akan digambar sebagai kotak bounding box di layer `BUILDINGS`.<br><br>
""", unsafe_allow_html=True)

    st.caption("Upload file .KML atau .KMZ (area batas cluster)")
    uploaded_file = st.file_uploader("Upload file", type=["kml", "kmz"])

    if uploaded_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{uploaded_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(uploaded_file.read())

                # Jika KMZ → ekstrak KML
                if temp_input.lower().endswith(".kmz"):
                    temp_input = extract_kmz_to_kml(temp_input, "/tmp")

                output_dir = "/tmp/output"
                dxf_path, geojson_path, ok = process_kml_to_dxf(temp_input, output_dir)

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
