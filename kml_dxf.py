import os
import zipfile
import requests
import pandas as pd
import geopandas as gpd
import streamlit as st
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString
from shapely.ops import unary_union, linemerge, polygonize
import mercantile
from mapbox_vector_tile import decode as mvt_decode

TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10
HERE_API_KEY = "iWCrFicKYt9_AOCtg76h76MlqZkVTn94eHbBl_cE8m0"
HERE_TILE_ZOOM = 15
HERE_V3_URL = "https://vector.hereapi.com/v3/tiles/base/mc/{z}/{x}/{y}/mvt?apiKey={api_key}"

# =====================
# CLASSIFY LAYER
# =====================
def classify_layer(kind_or_fc):
    val = (kind_or_fc or "").lower()
    if val in ["motorway", "trunk", "highway", "freeway", "fc1"]:
        return 'HIGHWAYS', 10
    elif val in ["primary", "primary_link", "fc2", "secondary", "tertiary", "main"]:
        return 'MAJOR_ROADS', 10
    elif val in ["residential", "street", "service", "unclassified", "fc3", "fc4", "local"]:
        return 'MINOR_ROADS', 10
    elif val in ["path", "footway", "cycleway", "track", "trail"]:
        return 'PATHS', 5
    return 'OTHER', DEFAULT_WIDTH

# =====================
# KML / KMZ POLYGON
# =====================
def extract_polygon_from_kml(kml_path):
    gdf = gpd.read_file(kml_path)
    polygons = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polygons.empty:
        raise Exception("No Polygon found in KML")
    return polygons.unary_union, (polygons.crs if polygons.crs else "EPSG:4326")

def extract_kmz_to_kml(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kml_files = [f for f in kmz_file.namelist() if f.lower().endswith('.kml')]
        if not kml_files:
            raise Exception("Tidak ada file .KML di dalam KMZ")
        kmz_file.extract(kml_files[0], extract_dir)
        return os.path.join(extract_dir, kml_files[0])

# =====================
# HERE ROADS
# =====================
def _tile_bounds_lonlat(x, y, z):
    b = mercantile.bounds(x, y, z)
    return (b.west, b.south, b.east, b.north)

def _geom_coords_from_mvt(rings, extent, x, y, z):
    minx, miny, maxx, maxy = _tile_bounds_lonlat(x, y, z)
    coords_list = []
    for ring in rings:
        ll = []
        for px, py in ring:
            lon = minx + (maxx - minx) * (px / extent)
            lat = maxy - (maxy - miny) * (py / extent)  # flip Y
            ll.append((lon, lat))
        coords_list.append(ll)
    return coords_list

def get_here_roads(polygon, polygon_crs, zoom=HERE_TILE_ZOOM):
    poly_ll = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs("EPSG:4326").iloc[0]
    minx, miny, maxx, maxy = poly_ll.bounds
    tiles = list(mercantile.tiles(minx, miny, maxx, maxy, zooms=[zoom]))

    features = []
    for t in tiles:
        url = HERE_V3_URL.format(z=t.z, x=t.x, y=t.y, api_key=HERE_API_KEY)
        r = requests.get(url, timeout=20)
        if r.status_code != 200: 
            continue
        mvt = mvt_decode(r.content)
        for lname, layer in mvt.items():
            if "road" not in lname.lower() and "transport" not in lname.lower():
                continue
            extent = layer.get("extent", 4096)
            for feat in layer["features"]:
                if feat["geometry"]["type"] == "LineString":
                    lines = _geom_coords_from_mvt([feat["geometry"]["coordinates"]], extent, t.x, t.y, t.z)
                    if lines and len(lines[0]) > 1:
                        features.append({"geometry": LineString(lines[0]), "kind": feat.get("properties", {}).get("kind","")})
                elif feat["geometry"]["type"] == "MultiLineString":
                    lines = _geom_coords_from_mvt(feat["geometry"]["coordinates"], extent, t.x, t.y, t.z)
                    ml = MultiLineString([LineString(l) for l in lines if len(l) > 1])
                    if len(ml.geoms) > 0:
                        features.append({"geometry": ml, "kind": feat.get("properties", {}).get("kind","")})

    if not features:
        return gpd.GeoDataFrame(geometry=[], crs="EPSG:4326")

    gdf = gpd.GeoDataFrame(features, geometry=[f["geometry"] for f in features], crs="EPSG:4326")
    gdf["highway"] = gdf["kind"]
    gdf = gdf.clip(poly_ll).reset_index(drop=True)
    return gdf

# =====================
# EXPORT DXF
# =====================
def export_to_dxf(gdf_roads, dxf_path, polygon=None, polygon_crs=None):
    doc = ezdxf.new()
    msp = doc.modelspace()
    buffers = []

    for _, row in gdf_roads.iterrows():
        geom = row.geometry
        layer, width = classify_layer(str(row.get("highway", "")))
        if geom.is_empty:
            continue
        if isinstance(geom, (LineString, MultiLineString)):
            merged = geom if isinstance(geom, LineString) else linemerge(geom)
            buffered = merged.buffer(width/2, resolution=8, join_style=2)
            buffers.append(buffered)

    if buffers:
        all_union = unary_union(buffers)
        outlines = list(polygonize(all_union.boundary))
        all_pts = [(x,y) for g in outlines for x,y in g.exterior.coords]
        min_x, min_y = min(x for x,y in all_pts), min(y for x,y in all_pts)
        for g in outlines:
            coords = [(x-min_x, y-min_y) for x,y in g.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer":"ROADS"})

    if polygon is not None and polygon_crs is not None:
        poly = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs(TARGET_EPSG).iloc[0]
        if poly.geom_type == "Polygon":
            coords = [(x, y) for x,y in poly.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer":"BOUNDARY"})

    doc.saveas(dxf_path)

# =====================
# PROSES KML → DXF
# =====================
def process_kml_to_dxf(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml(kml_path)
    roads = get_here_roads(polygon, polygon_crs)
    if roads.empty:
        raise Exception("Tidak ada jalan ditemukan")
    roads_utm = roads.to_crs(TARGET_EPSG)
    dxf_path = os.path.join(output_dir, "roadmap_here.dxf")
    export_to_dxf(roads_utm, dxf_path, polygon, polygon_crs)
    return dxf_path, True

# =====================
# STREAMLIT
# =====================
def run_kml_dxf():
    st.title("🌍 KML/KMZ → Jalan (HERE API)")
    kml_file = st.file_uploader("Upload file .KML atau .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())
                if temp_input.lower().endswith(".kmz"):
                    temp_input = extract_kmz_to_kml(temp_input, "/tmp")
                output_dir = "/tmp/output"
                dxf_path, ok = process_kml_to_dxf(temp_input, output_dir)
                if ok:
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF", data=f, file_name="roadmap_here.dxf")
            except Exception as e:
                st.error(f"❌ Error: {e}")

if __name__ == "__main__":
    run_kml_dxf()
