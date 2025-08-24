import os
import requests
import mercantile
import mapbox_vector_tile
import geopandas as gpd
import streamlit as st
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, shape
from shapely.ops import unary_union, linemerge, polygonize, snap

# --------------------------
# KONFIGURASI
# --------------------------
TARGET_EPSG = "EPSG:32760"  # UTM 60S
DEFAULT_WIDTH = 10
HERE_API_KEY = "ISI_API_KEY_HERE"   # <<== GANTI DENGAN API KEY MU


# --------------------------
# HELPER: Mapping jalan
# --------------------------
def classify_layer(hwy):
    """Mapping HERE functionalClass -> Layer DXF"""
    if str(hwy) in ["1", "2"]:       # FC1, FC2 = jalan besar
        return "HIGHWAYS", 12
    elif str(hwy) in ["3"]:          # FC3 = primary
        return "MAJOR_ROADS", 10
    elif str(hwy) in ["4"]:          # FC4 = minor
        return "MINOR_ROADS", 8
    elif str(hwy) in ["5"]:          # FC5 = service / lokal
        return "PATHS", 6
    return "OTHER", DEFAULT_WIDTH


# --------------------------
# Ekstrak polygon dari KML
# --------------------------
def extract_polygon_from_kml(kml_path):
    gdf = gpd.read_file(kml_path)
    polygons = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polygons.empty:
        raise Exception("❌ No Polygon found in KML")
    return unary_union(polygons.geometry), polygons.crs


# --------------------------
# Ambil jalan dari HERE
# --------------------------
def get_here_roads(polygon):
    """Ambil data jalan dari HERE Maps API berdasarkan boundary polygon"""
    minx, miny, maxx, maxy = polygon.bounds

    zoom = 14  # detail (bisa 15 kalau perlu lebih halus)
    tiles = list(mercantile.tiles(minx, miny, maxx, maxy, zoom))

    features = []
    for tile in tiles:
        url = f"https://vector.hereapi.com/v2/vectortiles/base/mc/{zoom}/{tile.x}/{tile.y}/omv?apiKey={HERE_API_KEY}"
        resp = requests.get(url)
        if resp.status_code != 200:
            continue

        data = resp.content
        tile_data = mapbox_vector_tile.decode(data)

        # cari layer jalan
        for layer_name in tile_data.keys():
            if "ROAD_GEOM" in layer_name:
                for feat in tile_data[layer_name]["features"]:
                    try:
                        geom = shape(feat["geometry"])
                        hwy = feat["properties"].get("functionalClass", "5")
                        features.append({"geometry": geom, "highway": hwy})
                    except Exception:
                        continue

    if not features:
        raise Exception("❌ Tidak ada jalan ditemukan dari HERE API")

    gdf = gpd.GeoDataFrame(features, crs="EPSG:4326")
    # Clip ke polygon
    gdf = gpd.clip(gdf, gpd.GeoSeries([polygon], crs="EPSG:4326"))
    return gdf


# --------------------------
# Bersihkan Z
# --------------------------
def strip_z(geom):
    if geom.geom_type == "LineString" and geom.has_z:
        return LineString([(x, y) for x, y, *_ in geom.coords])
    elif geom.geom_type == "MultiLineString":
        return MultiLineString([
            LineString([(x, y) for x, y, *_ in line.coords]) if line.has_z else line
            for line in geom.geoms
        ])
    return geom


# --------------------------
# Export ke DXF
# --------------------------
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

    for outline in outlines:
        coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})

    # gambar boundary polygon
    if polygon is not None and polygon_crs is not None:
        poly = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs(TARGET_EPSG).iloc[0]
        if poly.geom_type == "Polygon":
            coords = [(pt[0] - min_x, pt[1] - min_y) for pt in poly.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})
        elif poly.geom_type == "MultiPolygon":
            for p in poly.geoms:
                coords = [(pt[0] - min_x, pt[1] - min_y) for pt in p.exterior.coords]
                msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})

    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)


# --------------------------
# Main process
# --------------------------
def process_kml_to_dxf(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml(kml_path)
    roads = get_here_roads(polygon)

    geojson_path = os.path.join(output_dir, "roadmap_here.geojson")
    dxf_path = os.path.join(output_dir, "roadmap_here.dxf")

    if not roads.empty:
        roads_utm = roads.to_crs(TARGET_EPSG)
        roads_utm.to_file(geojson_path, driver="GeoJSON")
        export_to_dxf(roads_utm, dxf_path, polygon=polygon, polygon_crs=polygon_crs)
        return dxf_path, geojson_path, True
    else:
        raise Exception("❌ Tidak ada jalan ditemukan di dalam area polygon.")


# --------------------------
# Streamlit UI
# --------------------------
def run_kml_dxf():
    st.title("🌍 KML → Road Converter (HERE API)")
    st.caption("Upload file .KML (area batas cluster)")

    kml_file = st.file_uploader("Upload file .KML", type=["kml"])

    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())

                output_dir = "/tmp/output"
                dxf_path, geojson_path, ok = process_kml_to_dxf(temp_input, output_dir)

                if ok:
                    st.success("✅ Berhasil diekspor ke DXF (HERE API)!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF", data=f, file_name="roadmap_here.dxf")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")
