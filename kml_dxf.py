import os
import zipfile
import requests
import mercantile
from mapbox_vector_tile import decode as mvt_decode

import geopandas as gpd
import streamlit as st
import ezdxf

from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString
from shapely.ops import unary_union, linemerge, polygonize

# =========================
# KONFIGURASI
# =========================
TARGET_EPSG = "EPSG:32760"   # UTM 60S (sesuaikan jika perlu)
DEFAULT_WIDTH = 10

# Pakai API key milikmu
HERE_API_KEY = "iWCrFicKYt9_AOCtg76h76MlqZkVTn94eHbBl_cE8m0"

# Vector Tile endpoint (v3 mvt & fallback v2 omv)
HERE_V3_URL = "https://vector.hereapi.com/v3/tiles/base/mc/{z}/{x}/{y}/mvt?apiKey={api_key}"
HERE_V2_URL = "https://vector.hereapi.com/v2/vectortiles/base/mc/{z}/{x}/{y}/omv?apiKey={api_key}"

# Level detail tile. 15 cukup detail untuk jalan lingkungan.
HERE_TILE_ZOOM = 15

# =========================
# KLASIFIKASI LAYER & LEBAR
# (Mapping untuk atribut HERE: kind/functionalClass)
# =========================
def classify_layer(here_kind_or_fc: str):
    s = (here_kind_or_fc or "").strip().lower()
    # Functional Class (fc1 paling besar s/d fc5 lokal)
    if s in ["fc1", "motorway", "freeway", "trunk", "highway"]:
        return "HIGHWAYS", 14
    if s in ["fc2", "primary", "primary_link", "main", "trunk_link"]:
        return "MAJOR_ROADS", 12
    if s in ["fc3", "fc4", "secondary", "tertiary", "residential", "street", "service", "unclassified", "local"]:
        return "MINOR_ROADS", 8
    if s in ["fc5", "path", "footway", "cycleway", "track", "trail"]:
        return "PATHS", 5
    return "OTHER", DEFAULT_WIDTH

# =========================
# UTIL POLYGON & KML
# =========================
def extract_polygon_from_kml(kml_path):
    gdf = gpd.read_file(kml_path)
    polys = gdf[gdf.geometry.type.isin(["Polygon", "MultiPolygon"])]
    if polys.empty:
        raise Exception("Tidak ada Polygon/MultiPolygon pada KML.")
    return unary_union(polys.geometry), (polys.crs if polys.crs is not None else "EPSG:4326")

def extract_kmz_to_kml(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, "r") as zf:
        kml_files = [f for f in zf.namelist() if f.lower().endswith(".kml")]
        if not kml_files:
            raise Exception("❌ Tidak ada file .KML di dalam KMZ.")
        zf.extract(kml_files[0], extract_dir)
        return os.path.join(extract_dir, kml_files[0])

# =========================
# KONVERSI KOORDINAT MVT → LON/LAT
# (mapbox-vector-tile memberi koordinat tile (0..extent); perlu diubah ke lon/lat)
# =========================
def _tile_bounds_lonlat(x, y, z):
    b = mercantile.bounds(x, y, z)
    return (b.west, b.south, b.east, b.north)  # minx, miny, maxx, maxy (lon/lat)

def _mvt_ring_to_lonlat(ring, extent, x, y, z):
    minx, miny, maxx, maxy = _tile_bounds_lonlat(x, y, z)
    ll = []
    for px, py in ring:
        lon = minx + (maxx - minx) * (px / extent)
        lat = maxy - (maxy - miny) * (py / extent)  # Y dibalik
        ll.append((lon, lat))
    return ll

def _mvt_lines_to_geom(geom_type, coords, extent, x, y, z):
    # Return LineString atau MultiLineString dalam lon/lat
    if geom_type == "LineString":
        line = _mvt_ring_to_lonlat(coords, extent, x, y, z)
        return LineString(line) if len(line) >= 2 else None
    elif geom_type == "MultiLineString":
        lines = []
        for ring in coords:
            line = _mvt_ring_to_lonlat(ring, extent, x, y, z)
            if len(line) >= 2:
                lines.append(LineString(line))
        if lines:
            return MultiLineString(lines) if len(lines) > 1 else lines[0]
    elif geom_type == "Polygon":
        # Ambil boundary ring sebagai garis (supaya kompatibel pipeline buffer)
        line = _mvt_ring_to_lonlat(coords[0], extent, x, y, z) if coords else []
        return LineString(line) if len(line) >= 2 else None
    elif geom_type == "MultiPolygon":
        lines = []
        for poly in coords:
            if not poly:
                continue
            ring = _mvt_ring_to_lonlat(poly[0], extent, x, y, z)
            if len(ring) >= 2:
                lines.append(LineString(ring))
        if lines:
            return MultiLineString(lines) if len(lines) > 1 else lines[0]
    return None

# =========================
# HERE TILES → GeoDataFrame JALAN
# =========================
def _fetch_here_tile_bytes(x, y, z):
    # Coba v3 mvt
    url_v3 = HERE_V3_URL.format(z=z, x=x, y=y, api_key=HERE_API_KEY)
    r = requests.get(url_v3, timeout=20)
    if r.status_code == 200 and r.content:
        return r.content, "mvt"
    # Fallback v2 omv
    url_v2 = HERE_V2_URL.format(z=z, x=x, y=y, api_key=HERE_API_KEY)
    r = requests.get(url_v2, timeout=20)
    if r.status_code == 200 and r.content:
        return r.content, "omv"
    return None, None

def get_here_roads(polygon_wgs84, zoom=HERE_TILE_ZOOM):
    """
    Ambil jalan dari HERE Vector Tiles untuk bbox polygon, lalu clip ke boundary.
    Output: GeoDataFrame EPSG:4326 dengan kolom 'highway' (diisi dari kind/fc).
    """
    # Pastikan polygon dalam WGS84
    poly_ll = polygon_wgs84

    # BBOX polygon → daftar tile di zoom tertentu
    minx, miny, maxx, maxy = poly_ll.bounds
    tiles = list(mercantile.tiles(minx, miny, maxx, maxy, zooms=[zoom]))

    feats = []
    for t in tiles:
        blob, fmt = _fetch_here_tile_bytes(t.x, t.y, t.z)
        if not blob:
            continue
        decoded = mvt_decode(blob)  # dict layer → {extent, features:[{geometry, properties}]}

        # Cari layer yang mengandung data jalan
        candidate_layers = [k for k in decoded.keys() if "road" in k.lower() or "transport" in k.lower() or "geom_fc" in k.lower()]
        if not candidate_layers:
            candidate_layers = list(decoded.keys())

        for lname in candidate_layers:
            layer = decoded[lname]
            extent = layer.get("extent", 4096)
            for feat in layer.get("features", []):
                g = feat.get("geometry", {})
                geom_type = g.get("type")
                coords = g.get("coordinates")
                if geom_type not in ("LineString", "MultiLineString", "Polygon", "MultiPolygon"):
                    continue

                geom_ll = _mvt_lines_to_geom(geom_type, coords, extent, t.x, t.y, t.z)
                if geom_ll is None or geom_ll.is_empty:
                    continue

                props = feat.get("properties", {}) or {}
                # Ambil info kelas/kind jalan
                here_kind = None
                for k in ["kind", "class", "functionalClass", "fclass", "fc", "road_class"]:
                    if k in props and props[k] not in [None, ""]:
                        here_kind = str(props[k])
                        break

                feats.append({"geometry": geom_ll, "highway": here_kind if here_kind else ""})

    if not feats:
        return gpd.GeoDataFrame(geometry=[], crs="EPSG:4326")

    gdf = gpd.GeoDataFrame(feats, geometry=[f["geometry"] for f in feats], crs="EPSG:4326")
    # Clip ke boundary polygon
    gdf = gpd.clip(gdf, gpd.GeoSeries([poly_ll], crs="EPSG:4326"))
    gdf = gdf[~gdf.geometry.is_empty & gdf.geometry.notnull()].reset_index(drop=True)
    return gdf

# =========================
# STRIP Z (jaga-jaga)
# =========================
def strip_z(geom):
    if geom.geom_type == "LineString" and hasattr(geom, "has_z") and geom.has_z:
        return LineString([(x, y) for x, y, *_ in geom.coords])
    elif geom.geom_type == "MultiLineString":
        parts = []
        for ln in geom.geoms:
            if hasattr(ln, "has_z") and ln.has_z:
                parts.append(LineString([(x, y) for x, y, *_ in ln.coords]))
            else:
                parts.append(ln)
        return MultiLineString(parts)
    return geom

# =========================
# EKSPOR DXF (jalan + boundary)
# =========================
def export_to_dxf(gdf_roads, dxf_path, polygon=None, polygon_crs=None):
    doc = ezdxf.new()
    msp = doc.modelspace()

    buffers = []
    for _, row in gdf_roads.iterrows():
        geom = strip_z(row.geometry)
        layer_name, width = classify_layer(str(row.get("highway", "")))

        if geom.is_empty or not geom.is_valid:
            continue

        merged = geom if isinstance(geom, LineString) else (
            linemerge(geom) if isinstance(geom, MultiLineString) else None
        )
        if merged is None:
            continue

        # Buffer line → area jalan (mendekati batas aspal)
        buffered = merged.buffer(width / 2, resolution=8, join_style=2)
        buffers.append(buffered)

    if not buffers:
        raise Exception("❌ Tidak ada garis valid untuk diekspor.")

    union = unary_union(buffers)
    outlines = list(polygonize(union.boundary))
    if not outlines:
        raise Exception("❌ Polygonize gagal menghasilkan outline.")

    # Offset koordinat supaya mulai dari (0,0)
    bounds_pts = [(x, y) for poly in outlines for (x, y) in poly.exterior.coords]
    min_x = min(x for x, y in bounds_pts)
    min_y = min(y for x, y in bounds_pts)

    for poly in outlines:
        coords = [(x - min_x, y - min_y) for (x, y) in poly.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})

    # Gambar boundary (opsional) – dalam TARGET_EPSG
    if polygon is not None and polygon_crs is not None:
        poly_target = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs(TARGET_EPSG).iloc[0]
        if poly_target.geom_type == "Polygon":
            coords = [(x - min_x, y - min_y) for (x, y) in poly_target.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})
        elif poly_target.geom_type == "MultiPolygon":
            for p in poly_target.geoms:
                coords = [(x - min_x, y - min_y) for (x, y) in p.exterior.coords]
                msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})

    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)

# =========================
# PIPELINE: KML → HERE ROADS → DXF
# =========================
def process_kml_to_dxf(kml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml(kml_path)

    # Pastikan polygon di WGS84 untuk HERE
    poly_ll = gpd.GeoSeries([polygon], crs=polygon_crs if polygon_crs else "EPSG:4326").to_crs("EPSG:4326").iloc[0]

    # Ambil jalan dari HERE dan clip ke boundary
    roads_ll = get_here_roads(poly_ll, zoom=HERE_TILE_ZOOM)
    if roads_ll.empty:
        raise Exception("Tidak ada jalan ditemukan dari HERE pada area boundary.")

    # Simpan GeoJSON (opsional) dalam WGS84
    geojson_path = os.path.join(output_dir, "roadmap_here.geojson")
    roads_ll.to_file(geojson_path, driver="GeoJSON")

    # Proyeksikan ke TARGET_EPSG untuk ekspor DXF (satuan meter)
    roads_utm = roads_ll.to_crs(TARGET_EPSG)

    dxf_path = os.path.join(output_dir, "roadmap_here.dxf")
    export_to_dxf(roads_utm, dxf_path, polygon=polygon, polygon_crs=polygon_crs if polygon_crs else "EPSG:4326")
    return dxf_path, geojson_path, True

# =========================
# STREAMLIT APP
# =========================
def run_kml_dxf():
    st.title("🌍 KML → Jalan (HERE API) → DXF")
    st.markdown("""
    <h3>Proses:</h3>
    <ol>
      <li>Ambil boundary dari KML.</li>
      <li>Unduh jalan dari HERE Vector Tiles (zoom default 15).</li>
      <li>Clip ke boundary polygon.</li>
      <li>Buffer line → area jalan, lalu ekspor DXF.</li>
    </ol>
    """, unsafe_allow_html=True)

    kml_file = st.file_uploader("Upload file .KML", type=["kml", "kmz"])

    zoom = st.slider("Zoom HERE tiles (detail)", min_value=12, max_value=17, value=HERE_TILE_ZOOM, step=1,
                     help="Semakin besar semakin detail (lebih banyak request). 15-16 untuk jalan lingkungan.")
    if zoom != HERE_TILE_ZOOM:
        global HERE_TILE_ZOOM
        HERE_TILE_ZOOM = zoom

    if kml_file:
        with st.spinner("💫 Mengambil jalan dari HERE & memproses..."):
            try:
                temp_input = f"/tmp/{kml_file.name}"
                with open(temp_input, "wb") as f:
                    f.write(kml_file.read())

                # KMZ → ekstrak KML
                if temp_input.lower().endswith(".kmz"):
                    temp_input = extract_kmz_to_kml(temp_input, "/tmp")

                output_dir = "/tmp/output"
                dxf_path, geojson_path, ok = process_kml_to_dxf(temp_input, output_dir)

                if ok:
                    st.success("✅ Berhasil diekspor!")
                    with open(dxf_path, "rb") as f:
                        st.download_button("⬇️ Download DXF (UTM 60S)", data=f, file_name="roadmap_here.dxf")
                    with open(geojson_path, "rb") as f:
                        st.download_button("⬇️ Download GeoJSON (WGS84)", data=f, file_name="roadmap_here.geojson")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

if __name__ == "__main__":
    run_kml_dxf()
