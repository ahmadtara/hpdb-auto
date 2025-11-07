import os
import zipfile
import requests
import geopandas as gpd
import streamlit as st
import osmnx as ox
import json
from shapely.geometry import shape, LineString, MultiLineString, Polygon, MultiPolygon
from shapely.ops import unary_union
import simplekml

HERE_API_KEY = "jGCMpa59MeURAH39Vzk94kutVqC3vl714_ZvcHodX14"

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
    return unary_union(polygons.geometry), polygons.crs or "EPSG:4326"

def get_osm_roads(polygon):
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(polygon, tags=tags)
    except Exception:
        return gpd.GeoDataFrame()
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    roads = roads.explode(index_parts=False)
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    # pastikan CRS 4326
    if roads.crs is not None:
        roads = roads.to_crs("EPSG:4326")
    try:
        roads = roads.clip(polygon)
    except Exception:
        # fallback: filter by intersects if clip gagal
        roads = roads[roads.intersects(polygon)]
    return roads.reset_index(drop=True)

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

def _linestring_coords_for_kml(geom):
    """Return list of (lon, lat) coordinates appropriate for simplekml."""
    if isinstance(geom, LineString):
        return [(x, y) for x, y in geom.coords]
    elif isinstance(geom, MultiLineString):
        # return list of lists for MultiLineString (we will create MultiGeometry)
        parts = []
        for line in geom.geoms:
            parts.append([(x, y) for x, y in line.coords])
        return parts
    return []

def export_to_kml_simple(gdf, polygon, output_path):
    """
    Export roads GeoDataFrame (EPSG:4326) and polygon boundary to a KML file using simplekml.
    """
    kml_obj = simplekml.Kml()
    # Boundary as a Polygon Placemark (outer boundary only)
    if polygon is not None:
        # polygon may be Polygon or MultiPolygon
        if isinstance(polygon, Polygon):
            outer = [(x, y) for x, y in polygon.exterior.coords]
            pol = kml_obj.newpolygon(name="Boundary")
            pol.outerboundaryis = outer
        elif isinstance(polygon, MultiPolygon):
            # create one polygon placemark per part
            for i, p in enumerate(polygon.geoms):
                outer = [(x, y) for x, y in p.exterior.coords]
                pol = kml_obj.newpolygon(name=f"Boundary_{i+1}")
                pol.outerboundaryis = outer

    # Roads: create one placemark per LineString (or one MultiGeometry for MultiLine)
    for idx, row in gdf.iterrows():
        geom = row.geometry
        if geom is None or geom.is_empty or not geom.is_valid:
            continue
        hwy = str(row.get("highway", row.get("type", "")))
        layer = classify_layer(hwy)
        # set simple style (optional): name, description
        name = layer
        desc = f"highway={hwy}"
        if isinstance(geom, LineString):
            ls = kml_obj.newlinestring(name=name, description=desc)
            ls.coords = [(x, y) for x, y in geom.coords]
            ls.style.linestyle.width = 2
        elif isinstance(geom, MultiLineString):
            # add each part as a linestring (so all parts visible)
            for i, part in enumerate(geom.geoms):
                ls = kml_obj.newlinestring(name=f"{name}_{i+1}", description=desc)
                ls.coords = [(x, y) for x, y in part.coords]
                ls.style.linestyle.width = 2
        else:
            # ignore non-line geometries
            continue

    # save
    kml_obj.save(output_path)

def process_kml_to_kml(kml_path, output_dir):
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
    export_to_kml_simple(roads, polygon, kml_output)
    return kml_output, True
# ---------------------- STREAMLIT APP ----------------------

def run_kml_dxf():
    st.title("🌍 KML/KMZ → KML Road Converter (OSM + HERE Fallback)")
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
                    st.success("✅ Berhasil diekspor ke KML!")
                    with open(kml_path, "rb") as f:
                        st.download_button("⬇️ Download KML", data=f, file_name="roadmap.kml")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")


