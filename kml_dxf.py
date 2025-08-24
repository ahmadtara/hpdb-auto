import os
import zipfile
import tempfile
from fastkml import kml
import geopandas as gpd
import streamlit as st
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString
from shapely.ops import unary_union, linemerge, snap
import osmnx as ox

TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 6   # default jalan kecil

# 🔹 Klasifikasi jalan berdasarkan tipe OSM
def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']:
        return 'HIGHWAYS', 18
    elif hwy in ['secondary', 'tertiary']:
        return 'MAJOR_ROADS', 12
    elif hwy in ['residential', 'unclassified', 'service']:
        return 'MINOR_ROADS', 8
    elif hwy in ['living_street']:
        return 'RESIDENTIAL', 6
    elif hwy in ['footway', 'path', 'cycleway', 'pedestrian']:
        return 'PATHS', 3
    return 'OTHER', DEFAULT_WIDTH


# 🔹 Extract polygon dari KML atau KMZ
def extract_polygon_from_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".kmz":
        with zipfile.ZipFile(filepath, "r") as z:
            kml_files = [f for f in z.namelist() if f.endswith(".kml")]
            if not kml_files:
                raise Exception("❌ Tidak ada file .kml di dalam KMZ")
            tempdir = tempfile.mkdtemp()
            z.extract(kml_files[0], tempdir)
            filepath = os.path.join(tempdir, kml_files[0])

    gdf = gpd.read_file(filepath)
    polygons = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polygons.empty:
        raise Exception("❌ Tidak ada Polygon ditemukan di file")
    return unary_union(polygons.geometry), polygons.crs


# 🔹 Ambil jalan dari OSM
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


# 🔹 Hilangkan Z koordinat
def strip_z(geom):
    if geom.geom_type == "LineString" and geom.has_z:
        return LineString([(x, y) for x, y, *_ in geom.coords])
    elif geom.geom_type == "MultiLineString":
        return MultiLineString([
            LineString([(x, y) for x, y, *_ in line.coords]) if line.has_z else line
            for line in geom.geoms
        ])
    return geom


# 🔹 Export ke DXF (tanpa tabrakan garis)
def export_to_dxf(gdf, dxf_path, polygon=None, polygon_crs=None):
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
            buffered = merged.buffer(width / 2, resolution=8, join_style=1)
            all_buffers.append(buffered)

    if not all_buffers:
        raise Exception("❌ Tidak ada garis valid untuk diekspor.")

    # 🔹 Union langsung → tidak ada tabrakan
    all_union = unary_union(all_buffers)

    # Outline hasil union
    outlines = []
    if all_union.geom_type == "Polygon":
        outlines = [all_union]
    elif all_union.geom_type == "MultiPolygon":
        outlines = list(all_union.geoms)

    # Normalisasi posisi supaya (0,0) di kiri bawah
    bounds = [(pt[0], pt[1]) for geom in outlines for pt in geom.exterior.coords]
    min_x = min(x for x, y in bounds)
    min_y = min(y for x, y in bounds)

    for outline in outlines:
        coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})

    # Tambah boundary polygon
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


# 🔹 Proses utama
def process_kml_to_dxf(input_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_file(input_path)
    roads = get_osm_roads(polygon)

    geojson_path = os.path.join(output_dir, "roadmap_osm.geojson")
    dxf_path = os.path.join(output_dir, "roadmap_osm.dxf")

    if not roads.empty:
        roads_utm = roads.to_crs(TARGET_EPSG)
        roads_utm.to_file(geojson_path, driver="GeoJSON")
        export_to_dxf(roads_utm, dxf_path, polygon=polygon, polygon_crs=polygon_crs)
        return dxf_path, geojson_path, True
    else:
        raise Exception("Tidak ada jalan ditemukan di dalam area polygon.")


# 🔹 Streamlit app
def run_kml_dxf():
    st.title("🌍 KML/KMZ → Road Converter")
    st.markdown("""
    <h2>👋 Hai, <span style='color:#0A84FF'>obi</span></h2>
    ✅ <span style='font-weight:bold;'>CATATAN PENTING :</span><br><br>
    1️⃣ Upload file **KML/KMZ** berisi polygon batas.<br>
    2️⃣ Sistem akan otomatis ambil jalan dari OSM.<br>
    3️⃣ Hasil bisa didownload dalam format **DXF (UTM 60S)**.<br>
    """, unsafe_allow_html=True)
    st.caption("Upload file .KML atau .KMZ (area batas cluster)")

    input_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])

    if input_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{input_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(input_file.read())

                output_dir = "/tmp/output"
                dxf_path, geojson_path, ok = process_kml_to_dxf(temp_input, output_dir)

                if ok:
                    st.success("✅ Berhasil diekspor ke DXF!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download Jalan Autocad UTM 60", data=f, file_name="roadmap_osm.dxf")
             
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")
