import os
import zipfile
import requests
from fastkml import kml
import geopandas as gpd
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, shape
from shapely.ops import unary_union, linemerge, polygonize
import osmnx as ox
import json
from shapely import wkt


# ----------------------------
# KONFIGURASI
# ----------------------------
TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10
HERE_API_KEY = "jGCMpa59MeURAH39Vzk94kutVqC3vl714_ZvcHodX14"


# ----------------------------
# UTILITIES
# ----------------------------
def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']:
        return 'HIGHWAYS', 14
    elif hwy in ['secondary', 'tertiary']:
        return 'MAJOR_ROADS', 10
    elif hwy in ['residential', 'unclassified', 'service']:
        return 'MINOR_ROADS', 8
    elif hwy in ['footway', 'path', 'cycleway']:
        return 'PATHS', 4
    return 'OTHER', DEFAULT_WIDTH


def strip_z(geom):
    """Remove Z coords safely."""
    if geom is None:
        return geom
    try:
        if geom.geom_type == "LineString" and getattr(geom, "has_z", False):
            return LineString([(x, y) for x, y, *_ in geom.coords])
        if geom.geom_type == "MultiLineString":
            parts = []
            for g in geom.geoms:
                if getattr(g, "has_z", False):
                    parts.append(LineString([(x, y) for x, y, *_ in g.coords]))
                else:
                    parts.append(g)
            return MultiLineString(parts)
        if geom.geom_type == "Polygon" and getattr(geom, "has_z", False):
            exterior = [(x, y) for x, y, *_ in geom.exterior.coords]
            interiors = [[(x, y) for x, y, *_ in ring.coords] for ring in geom.interiors]
            return Polygon(exterior, interiors)
    except Exception:
        pass
    return geom


# ----------------------------
# EXTRACT POLYGON FROM KML/KMZ
# ----------------------------
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
        raise Exception("❌ No Polygon found in KML/KMZ")
    return unary_union(polygons.geometry), polygons.crs


# ----------------------------
# GET ROADS (OSM)
# ----------------------------
def get_osm_roads(polygon):
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(polygon, tags=tags)
    except Exception:
        return gpd.GeoDataFrame()
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    # explode available in geopandas >=0.10: use index_parts=False for consistent behavior
    try:
        roads = roads.explode(index_parts=False)
    except TypeError:
        roads = roads.explode()
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    roads = roads.clip(polygon)
    return roads.reset_index(drop=True)


# ----------------------------
# GET ROADS (HERE fallback)
# ----------------------------
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
        try:
            geom = shape(f["geometry"])
        except Exception:
            continue
        if geom.intersects(polygon):
            features.append({"geometry": geom.intersection(polygon), "properties": f.get("properties", {})})
    if not features:
        return gpd.GeoDataFrame()
    return gpd.GeoDataFrame(features, crs="EPSG:4326")


# ----------------------------
# EXPORT TO DXF
# ----------------------------
def export_to_dxf(gdf, dxf_path, polygon=None, polygon_crs=None):
    doc = ezdxf.new()
    msp = doc.modelspace()
    all_buffers = []
    for _, row in gdf.iterrows():
        geom = strip_z(row.geometry)
        hwy = str(row.get("highway", row.get("type", "")))
        layer, width = classify_layer(hwy)
        if geom is None or getattr(geom, "is_empty", False) or not getattr(geom, "is_valid", True):
            continue
        merged = linemerge(geom) if isinstance(geom, MultiLineString) else geom
        if isinstance(merged, (LineString, MultiLineString)):
            buffered = merged.buffer(width / 2, resolution=8, join_style=2)
            all_buffers.append(buffered)
    if not all_buffers:
        raise Exception("❌ Tidak ada garis valid untuk diekspor.")
    all_union = unary_union(all_buffers)
    outlines = list(polygonize(all_union.boundary))
    min_x = min(pt[0] for geom in outlines for pt in geom.exterior.coords)
    min_y = min(pt[1] for geom in outlines for pt in geom.exterior.coords)
    for outline in outlines:
        coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})
    # add boundary
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


# ----------------------------
# EXPORT TO KML (SAFE)
# ----------------------------
def export_to_kml(gdf, kml_path):
    """
    Export GeoDataFrame to KML safely:
    - ensures EPSG:4326
    - parses WKT or JSON-string geometries
    - splits Multi* into parts
    - uses constructor geometry=... (fastkml)
    """
    from fastkml import kml as fkml
    from shapely.geometry import mapping, shape as shapely_shape, Polygon as ShapelyPolygon, LineString as ShapelyLineString

    ns = '{http://www.opengis.net/kml/2.2}'
    doc_k = fkml.KML()
    doc = fkml.Document(ns, 'docid', 'Road Export', 'Roadmap Exported from Python')
    doc_k.append(doc)
    folder = fkml.Folder(ns, 'f1', 'Roads', 'Extracted Roads')
    doc.append(folder)

    # ensure CRS
    if getattr(gdf, "crs", None) is not None and gdf.crs.to_string() != "EPSG:4326":
        try:
            gdf = gdf.to_crs("EPSG:4326")
        except Exception:
            pass

    for idx, row in gdf.iterrows():
        geom = row.geometry

        # parse if geometry is string
        if isinstance(geom, str):
            parsed = None
            try:
                parsed = wkt.loads(geom)
            except Exception:
                try:
                    parsed = shapely_shape(json.loads(geom))
                except Exception:
                    parsed = None
            if parsed is None:
                print(f"⚠️ Skip feature {idx}: geometry string parse failed")
                continue
            geom = parsed

        if geom is None or getattr(geom, "is_empty", False) or not getattr(geom, "is_valid", True):
            continue

        geom = strip_z(geom)

        # handle Multi* types
        parts = []
        if isinstance(geom, (MultiLineString, MultiPolygon)):
            parts = list(geom.geoms)
        else:
            parts = [geom]

        name = str(row.get("highway", row.get("type", "road")))

        for part in parts:
            # final check
            if part is None:
                continue
            try:
                pm = fkml.Placemark(ns, None, name, "", geometry=part)
                folder.append(pm)
            except Exception as e:
                print(f"⚠️ Could not append placemark idx={idx}: {e}")

    # write file
    with open(kml_path, "w", encoding="utf-8") as f:
        f.write(doc_k.to_string(prettyprint=True))


# ----------------------------
# PROCESS MAIN
# ----------------------------
def process_kml_to_dxf(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml_or_kmz(kml_path)

    # try OSM first
    roads = get_osm_roads(polygon)

    # fallback to HERE
    if roads.empty:
        roads = get_here_roads(polygon)

    if roads.empty:
        raise Exception("❌ Tidak ada jalan ditemukan (OSM & HERE kosong).")

    geojson_path = os.path.join(output_dir, "roadmap.geojson")
    dxf_path = os.path.join(output_dir, "roadmap.dxf")
    kml_path_out = os.path.join(output_dir, "roadmap.kml")

    # save geojson (EPSG:4326 or whatever roads currently have -> convert to TARGET_EPSG for DXF)
    try:
        roads_utm = roads.to_crs(TARGET_EPSG)
    except Exception:
        roads_utm = roads  # if fail, keep as-is
    roads_utm.to_file(geojson_path, driver="GeoJSON")

    # export DXF (expects UTM)
    export_to_dxf(roads_utm, dxf_path, polygon=polygon, polygon_crs=polygon_crs)

    # export KML (use original roads GDF; ensure lat/lon inside func)
    export_to_kml(roads, kml_path_out)

    return dxf_path, geojson_path, kml_path_out, True


# ----------------------------
# STREAMLIT WRAPPER (call from streamlit_app)
# ----------------------------
def run_kml_dxf():
    import streamlit as st
    st.title("🌍 KML/KMZ → Road Converter (OSM + HERE fallback)")
    kml_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())

                output_dir = "/tmp/output"
                dxf_path, geojson_path, kml_path_out, ok = process_kml_to_dxf(temp_input, output_dir)
                if ok:
                    st.success("✅ Berhasil diekspor ke DXF & KML!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF (UTM 60)", data=f, file_name="roadmap.dxf")
                    with open(kml_path_out, "rb") as f:
                        st.download_button("⬇️ Download KML", data=f, file_name="roadmap.kml")
                    with open(geojson_path, "rb") as f:
                        st.download_button("📁 Download GeoJSON", data=f, file_name="roadmap.geojson")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

# -------------------------------------------------------------------
# 🔹 Klasifikasi layer berdasarkan jenis jalan
# -------------------------------------------------------------------
def classify_layer(hwy):
    if hwy in ['motorway', 'trunk', 'primary']:
        return 'HIGHWAYS', 14
    elif hwy in ['secondary', 'tertiary']:
        return 'MAJOR_ROADS', 10
    elif hwy in ['residential', 'unclassified', 'service']:
        return 'MINOR_ROADS', 8
    elif hwy in ['footway', 'path', 'cycleway']:
        return 'PATHS', 4
    return 'OTHER', DEFAULT_WIDTH


# -------------------------------------------------------------------
# 🔹 Ambil poligon dari KML atau KMZ
# -------------------------------------------------------------------
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
        raise Exception("❌ No Polygon found in KML/KMZ")
    return unary_union(polygons.geometry), polygons.crs


# -------------------------------------------------------------------
# 🔹 Ambil jalan dari OSM
# -------------------------------------------------------------------
def get_osm_roads(polygon):
    tags = {"highway": True}
    try:
        roads = ox.features_from_polygon(polygon, tags=tags)
    except Exception:
        return gpd.GeoDataFrame()
    roads = roads[roads.geometry.type.isin(["LineString", "MultiLineString"])]
    roads = roads.explode(index_parts=False)
    roads = roads[~roads.geometry.is_empty & roads.geometry.notnull()]
    roads = roads.clip(polygon)
    return roads.reset_index(drop=True)


# -------------------------------------------------------------------
# 🔹 Fallback: Ambil jalan dari HERE API
# -------------------------------------------------------------------
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
            features.append({"geometry": geom.intersection(polygon), "properties": f["properties"]})
    if not features:
        return gpd.GeoDataFrame()
    return gpd.GeoDataFrame(features, crs="EPSG:4326")


# -------------------------------------------------------------------
# 🔹 Bersihkan koordinat Z
# -------------------------------------------------------------------
def strip_z(geom):
    if geom.geom_type == "LineString" and geom.has_z:
        return LineString([(x, y) for x, y, *_ in geom.coords])
    elif geom.geom_type == "MultiLineString":
        return MultiLineString([
            LineString([(x, y) for x, y, *_ in line.coords]) if line.has_z else line
            for line in geom.geoms
        ])
    return geom


# -------------------------------------------------------------------
# 🔹 Export ke DXF
# -------------------------------------------------------------------
def export_to_dxf(gdf, dxf_path, polygon=None, polygon_crs=None):
    doc = ezdxf.new()
    msp = doc.modelspace()
    all_buffers = []
    for _, row in gdf.iterrows():
        geom = strip_z(row.geometry)
        hwy = str(row.get("highway", row.get("type", "")))
        layer, width = classify_layer(hwy)
        if geom.is_empty or not geom.is_valid:
            continue
        merged = linemerge(geom) if isinstance(geom, MultiLineString) else geom
        if isinstance(merged, (LineString, MultiLineString)):
            buffered = merged.buffer(width / 2, resolution=8, join_style=2)
            all_buffers.append(buffered)
    if not all_buffers:
        raise Exception("❌ Tidak ada garis valid untuk diekspor.")
    all_union = unary_union(all_buffers)
    outlines = list(polygonize(all_union.boundary))
    min_x = min(pt[0] for geom in outlines for pt in geom.exterior.coords)
    min_y = min(pt[1] for geom in outlines for pt in geom.exterior.coords)
    for outline in outlines:
        coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})
    # add boundary
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


# -------------------------------------------------------------------
# 🔹 Export ke KML
# -------------------------------------------------------------------
def export_to_kml(gdf, kml_path):
    """
    Export GeoDataFrame -> KML safely.
    Handles: CRS -> EPSG:4326, WKT-string geometries, Multi* splitting.
    """
    from fastkml import kml
    from shapely.geometry import shape as shapely_shape, MultiLineString, MultiPolygon, Polygon, LineString
    from shapely import wkt

    ns = '{http://www.opengis.net/kml/2.2}'
    k = kml.KML()
    doc = kml.Document(ns, 'docid', 'Road Export', 'Roadmap Exported from Python')
    k.append(doc)
    folder = kml.Folder(ns, 'f1', 'Roads', 'Extracted Roads')
    doc.append(folder)

    # pastikan ada kolom geometry dan CRS
    if getattr(gdf, "crs", None) is not None and gdf.crs.to_string() != "EPSG:4326":
        try:
            gdf = gdf.to_crs("EPSG:4326")
        except Exception:
            # kalau gagal, lanjutkan — kita coba tetap proses tiap fitur
            pass

    for idx, row in gdf.iterrows():
        geom = row.geometry

        # jika geometry adalah string (WKT atau geojson string), coba parse
        if isinstance(geom, str):
            try:
                geom = wkt.loads(geom)
            except Exception:
                # coba interpret sebagai geojson mapping string -> skip kalau gagal
                try:
                    import json
                    gm = json.loads(geom)
                    geom = shapely_shape(gm)
                except Exception:
                    # skip feature
                    print(f"⚠️ Skip feature {idx}: geometry string parse failed")
                    continue

        # jika None atau empty skip
        if geom is None:
            continue
        try:
            geom = strip_z(geom)
        except Exception:
            pass
        if getattr(geom, "is_empty", False) or not getattr(geom, "is_valid", True):
            continue

        name = str(row.get("highway", row.get("type", "road")))

        # jika MultiLineString/MultiPolygon: buat multiple placemark
        if isinstance(geom, (MultiLineString, MultiPolygon)):
            parts = geom.geoms
        else:
            parts = [geom]

        for part in parts:
            # final safety: ensure part is shapely geom
            if isinstance(part, (LineString, Polygon)) or hasattr(part, "__geo_interface__"):
                try:
                    pm = kml.Placemark(ns, None, name, "", geometry=part)
                    folder.append(pm)
                except Exception as e:
                    print(f"⚠️ Could not append placemark for feature {idx}: {e}")
            else:
                print(f"⚠️ Skipping non-geometry part for feature {idx}")

    # tulis file
    with open(kml_path, "w", encoding="utf-8") as f:
        f.write(k.to_string(prettyprint=True))



# -------------------------------------------------------------------
# 🔹 Streamlit UI
# -------------------------------------------------------------------
def run_kml_dxf():
    st.title("🌍 KML/KMZ → Road Converter (OSM + HERE fallback)")
    kml_file = st.file_uploader("Upload file .KML / .KMZ", type=["kml", "kmz"])
    if kml_file:
        with st.spinner("💫 Memproses file..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())

                output_dir = "/tmp/output"
                dxf_path, geojson_path, kml_path_out, ok = process_kml_to_dxf(temp_input, output_dir)
                if ok:
                    st.success("✅ Berhasil diekspor ke DXF & KML!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF (UTM 60)", data=f, file_name="roadmap.dxf")
                    with open(kml_path_out, "rb") as f:
                        st.download_button("⬇️ Download KML", data=f, file_name="roadmap.kml")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")





