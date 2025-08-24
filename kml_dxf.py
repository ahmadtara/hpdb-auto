import os
import zipfile
import requests
import pandas as pd
import geopandas as gpd
import streamlit as st
import ezdxf
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, box, shape, mapping
from shapely.ops import unary_union, linemerge, snap, polygonize
from shapely import wkt
import mercantile
import math
from mapbox_vector_tile import decode as mvt_decode

# =====================
# KONFIG
# =====================
TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10
MAX_CSV_ROWS = 50000  # batasi row CSV agar tidak crash

# ====== MASUKKAN API KEY HERE DI SINI ======
HERE_API_KEY = "iWCrFicKYt9_AOCtg76h76MlqZkVTn94eHbBl_cE8m0"

# Zoom tile default untuk tingkat detail jalan.
# 14–16 umumnya cukup. Makin besar = makin detail tapi lebih banyak tile.
HERE_TILE_ZOOM = 15

# Endpoint Vector Tile HERE (v3) – fallback ke v2 jika perlu.
# v3 (mvt): https://vector.hereapi.com/v3/tiles/base/mc/{z}/{x}/{y}/mvt?apiKey=...
# v2 (omv): https://vector.hereapi.com/v2/vectortiles/base/mc/{z}/{x}/{y}/omv?apiKey=...
HERE_V3_URL = "https://vector.hereapi.com/v3/tiles/base/mc/{z}/{x}/{y}/mvt?apiKey={api_key}"
HERE_V2_URL = "https://vector.hereapi.com/v2/vectortiles/base/mc/{z}/{x}/{y}/omv?apiKey={api_key}"

# =====================
# LAYER / WIDTH
# =====================
def classify_layer(kind_or_fc):
    """
    Mapping sederhana dari atribut HERE (kind / functionalClass)
    ke layer internal + lebar garis (untuk buffer).
    """
    val = (kind_or_fc or "").lower()
    if val in ["motorway", "trunk", "highway", "freeway", "fc1"]:
        return 'HIGHWAYS', 10
    elif val in ["primary", "primary_link", "fc2", "secondary", "tertiary", "main"]:
        return 'MAJOR_ROADS', 10
    elif val in ["residential", "street", "service", "unclassified", "fc3", "fc4", "local"]:
        return 'MINOR_ROADS', 10
    elif val in ["path", "footway", "cycleway", "track", "trail"]:
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
# KONVERSI KOORDINAT MVT TILE → LON/LAT
# =====================
def _interp(a0, a1, t):
    return a0 + (a1 - a0) * t

def _tile_bounds_lonlat(x, y, z):
    """Batas tile dalam lon/lat menggunakan mercantile."""
    b = mercantile.bounds(x, y, z)
    return (b.west, b.south, b.east, b.north)  # (minx, miny, maxx, maxy) in lon/lat

def _geom_coords_from_mvt(commands, extent, x, y, z):
    """
    Ubah koordinat MVT (0..extent) ke lon/lat berdasarkan bbox tile.
    Return list of LineString koord (list of tuples).
    """
    minx, miny, maxx, maxy = _tile_bounds_lonlat(x, y, z)
    # MVT origin (0,0) di kiri atas tile; sumbu y ke bawah.
    # Kita perlu membalik sumbu Y.
    lines = []
    for ring in commands:
        ll = []
        for (px, py) in ring:
            lon = _interp(minx, maxx, px / extent)
            lat = _interp(maxy, miny, py / extent)  # flip Y
            ll.append((lon, lat))
        lines.append(ll)
    return lines

# =====================
# HERE VECTOR TILE → GeoDataFrame (roads)
# =====================
def _fetch_here_tile(x, y, z):
    # Coba v3 (mvt) dulu, fallback ke v2 (omv)
    url_v3 = HERE_V3_URL.format(z=z, x=x, y=y, api_key=HERE_API_KEY)
    r = requests.get(url_v3, timeout=20)
    if r.status_code == 200 and r.content:
        return r.content, "mvt"
    # fallback v2
    url_v2 = HERE_V2_URL.format(z=z, x=x, y=y, api_key=HERE_API_KEY)
    r = requests.get(url_v2, timeout=20)
    if r.status_code == 200 and r.content:
        return r.content, "omv"
    return None, None

def get_here_roads(polygon, polygon_crs, zoom=HERE_TILE_ZOOM):
    """
    Ambil geometri jalan dari HERE Vector Tile API dalam area polygon.
    Hasil GeoDataFrame EPSG:4326, hanya LineString/MultiLineString,
    terklip ke polygon.
    """
    poly_ll = ensure_wgs84_polygon(polygon, polygon_crs)
    minx, miny, maxx, maxy = poly_ll.bounds

    # Kumpulkan daftar tile yang menutupi bbox polygon
    tiles = list(mercantile.tiles(minx, miny, maxx, maxy, zooms=[zoom]))
    features = []
    for t in tiles:
        blob, fmt = _fetch_here_tile(t.x, t.y, t.z)
        if not blob:
            continue
        # Decode MVT ke dict: {layer: {features:[{geometry, properties, ...}]}}
        mvt = mvt_decode(blob)
        # Layer roads di HERE bisa muncul sebagai "roads", "transportation", dll.
        # Kita iter semua layer yang mengandung data garis jalan.
        candidate_layers = [k for k in mvt.keys() if "road" in k.lower() or "transport" in k.lower()]
        if not candidate_layers:
            # Kalau tidak ketemu, coba semua layer, nanti difilter by geometry type LineString
            candidate_layers = list(mvt.keys())

        # extent default mvt = 4096; ada di metadata tiap layer
        for lname in candidate_layers:
            layer = mvt[lname]
            extent = layer.get("extent", 4096)
            for feat in layer["features"]:
                geom_type = feat["geometry"]["type"]
                coords_cmd = feat["geometry"]["coordinates"]
                props = feat.get("properties", {}) or {}

                # Kita fokus ke garis (LineString/MultiLineString) atau boundary polygonal roads
                if geom_type in ("LineString", "MultiLineString"):
                    # Koordinat MVT bisa nested: list of list of (x,y)
                    if geom_type == "LineString":
                        lines = _geom_coords_from_mvt([coords_cmd], extent, t.x, t.y, t.z)
                        line = LineString(lines[0]) if len(lines[0]) >= 2 else None
                        if line is not None and not line.is_empty:
                            features.append({"geometry": line, **props})
                    else:
                        lines = _geom_coords_from_mvt(coords_cmd, extent, t.x, t.y, t.z)
                        mls = MultiLineString([ln for ln in map(LineString, lines) if len(ln.coords) >= 2])
                        if len(mls.geoms) > 0:
                            features.append({"geometry": mls, **props})
                elif geom_type in ("Polygon", "MultiPolygon"):
                    # Beberapa wilayah (Polygonal roads) – gunakan boundary agar jadi garis tepi jalan
                    # supaya kompatibel dengan proses buffer kamu.
                    if geom_type == "Polygon":
                        lines = _geom_coords_from_mvt(coords_cmd, extent, t.x, t.y, t.z)
                        for ring in lines:
                            if len(ring) >= 2:
                                features.append({"geometry": LineString(ring), **props})
                    else:
                        # MultiPolygon => banyak ring
                        for poly_rings in coords_cmd:
                            lines = _geom_coords_from_mvt(poly_rings, extent, t.x, t.y, t.z)
                            for ring in lines:
                                if len(ring) >= 2:
                                    features.append({"geometry": LineString(ring), **props})

    if not features:
        return gpd.GeoDataFrame(geometry=[], crs="EPSG:4326")

    gdf = gpd.GeoDataFrame(features, geometry=[f["geometry"] for f in features], crs="EPSG:4326")
    # Tambahkan kolom "here_kind" dari properti yang umum (kind / class / functionalClass)
    def _kind(props_row):
        for k in ["kind", "class", "functionalClass", "road_class", "fclass", "fc"]:
            if k in props_row and props_row[k]:
                return str(props_row[k])
        return None

    # Simpan properti di kolom agar bisa diklasifikasikan
    # (GeoPandas menaruh properti di kolom kecuali 'geometry' – sudah disusun di atas)
    if "here_kind" not in gdf.columns:
        gdf["here_kind"] = None
    for i, row in gdf.iterrows():
        # row adalah Series; cari di raw dict
        # (karena di atas kita flatten props sebagai kolom, gunakan prioritas kolom jika ada)
        if pd.isna(row.get("here_kind")):
            hk = None
            for k in ["kind", "class", "functionalClass", "road_class", "fclass", "fc"]:
                if k in gdf.columns and pd.notna(row.get(k)):
                    hk = str(row.get(k)); break
            gdf.at[i, "here_kind"] = hk

    # Potong ke polygon boundary
    gdf = gdf.clip(poly_ll)
    gdf = gdf[~gdf.geometry.is_empty & gdf.geometry.notnull()]
    gdf = gdf.reset_index(drop=True)

    # Untuk kompatibilitas fungsi downstream yang pakai kolom "highway"
    gdf["highway"] = gdf["here_kind"].fillna("")

    return gdf

# =====================
# OSM BUILDINGS (tetap sebagai fallback / sumber bangunan)
# =====================
def get_osm_buildings(polygon, polygon_crs):
    import osmnx as ox  # impor lokal bila tersedia
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
# LOAD CSV GEO DENGAN CACHE
# =====================
@st.cache_data
def load_csv_safe(csv_file):
    temp_csv = f"/tmp/{csv_file.name}"
    with open(temp_csv, "wb") as f:
        f.write(csv_file.read())
    df = pd.read_csv(temp_csv, nrows=MAX_CSV_ROWS)
    gdf = gpd.GeoDataFrame(df, geometry=df['geometry'].apply(wkt.loads), crs="EPSG:4326")
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
        # pakai kolom "highway" (sudah diisi dari HERE kind)
        layer_name, width = classify_layer(str(row.get("highway", "")))
        if geom.is_empty or not geom.is_valid:
            continue

        if isinstance(geom, (LineString, MultiLineString)):
            merged = geom if isinstance(geom, LineString) else linemerge(geom)
            buffered = merged.buffer(width / 2, resolution=8, join_style=2)
            all_buffers.append(buffered)
        elif isinstance(geom, (Polygon, MultiPolygon)):
            all_buffers.append(geom.boundary)

    if all_buffers:
        all_union = unary_union(all_buffers)
        outlines = list(polygonize(all_union.boundary))
        # hitung offset supaya koordinat DXF mulai dari (0,0)
        bounds = [(pt[0], pt[1]) for geom in outlines for pt in geom.exterior.coords]
        min_x = min(x for x, y in bounds)
        min_y = min(y for x, y in bounds)
        for outline in outlines:
            coords = [(pt[0] - min_x, pt[1] - min_y) for pt in outline.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})
    else:
        min_x = min_y = 0

    if polygon is not None and polygon_crs is not None:
        poly = gpd.GeoSeries([polygon], crs=polygon_crs).to_crs(TARGET_EPSG).iloc[0]
        if poly.geom_type == 'Polygon':
            coords = [(pt[0] - min_x, pt[1] - min_y) for pt in poly.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})
        elif poly.geom_type == 'MultiPolygon':
            for p in poly.geoms:
                coords = [(pt[0] - min_x, pt[1] - min_y) for pt in p.exterior.coords]
                msp.add_lwpolyline(coords, dxfattribs={"layer": "BOUNDARY"})

    if buildings is not None and not buildings.empty:
        add_buildings_to_dxf(msp, buildings, min_x, min_y)

    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)

# =====================
# PROSES KML → DXF
# =====================
def process_kml_to_dxf(kml_path, output_dir, google_building_csv=None):
    os.makedirs(output_dir, exist_ok=True)
    polygon, polygon_crs = extract_polygon_from_kml(kml_path)

    # ==== GANTI: pakai HERE untuk ROADS ====
    roads = get_here_roads(polygon, polygon_crs, zoom=HERE_TILE_ZOOM)

    # ==== BUILDINGS tetap dari OSM / Google Open Buildings (opsional) ====
    buildings = get_osm_buildings(polygon, polygon_crs)
    if buildings.empty and google_building_csv is not None:
        buildings = load_google_buildings(google_building_csv, polygon, polygon_crs)

    geojson_path = os.path.join(output_dir, "roadmap_here.geojson")
    dxf_path = os.path.join(output_dir, "roadmap_here.dxf")

    if not roads.empty or (buildings is not None and not buildings.empty):
        roads_utm = roads.to_crs(TARGET_EPSG) if not roads.empty else gpd.GeoDataFrame(geometry=[], crs=TARGET_EPSG)
        buildings_utm = buildings.to_crs(TARGET_EPSG) if buildings is not None and not buildings.empty else None
        if not roads_utm.empty:
            roads.to_file(geojson_path, driver="GeoJSON")  # simpan tetap dalam WGS84 supaya ringan
        export_to_dxf(roads_utm if not roads_utm.empty else roads, dxf_path, polygon, polygon_crs, buildings_utm)
        return dxf_path, geojson_path, True
    else:
        raise Exception("Tidak ada jalan atau bangunan ditemukan di dalam area polygon.")

# =====================
# STREAMLIT
# =====================
def run_kml_dxf():
    st.title("🌍 KML/KMZ/CSV → Jalan & Kotak Bangunan dari Boundary (Sumber Jalan: HERE)")
    st.markdown("""
<h2>👋 Hai, <span style='color:#0A84FF'>bro</span></h2>
✅ <span style='font-weight:bold;'>CATATAN PENTING :</span><br><br>
1️⃣ Bisa upload `.KML` atau `.KMZ`.<br>
2️⃣ Atau upload CSV (WKT geometry).<br>
3️⃣ Sumber **jalan** dari <b>HERE Vector Tile API</b> (lebih sesuai kondisi aktual).<br>
4️⃣ Bangunan digambar sebagai kotak (bounding box) di layer <code>BUILDINGS</code>.<br>
5️⃣ Bisa juga upload Google Open Buildings CSV (.csv.gz) jika OSM kosong.<br><br>
""", unsafe_allow_html=True)

    kml_file = st.file_uploader("Upload file .KML atau .KMZ", type=["kml", "kmz"])
    google_csv = st.file_uploader("Opsional: Upload Google Open Buildings CSV (.csv.gz)", type=["csv.gz"])
    csv_file = st.file_uploader("Atau upload CSV langsung (WKT geometry)", type=["csv"])

    # ==== KML / KMZ ====
    if kml_file:
        with st.spinner("💫 Memproses file (mengambil roads dari HERE)..."):
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
                        st.download_button("⬇️ Download DXF (UTM 60)", data=f, file_name="roadmap_here.dxf")
                    with open(geojson_path, "rb") as f:
                        st.download_button("⬇️ Download GeoJSON (WGS84)", data=f, file_name="roadmap_here.geojson")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

    # ==== CSV ====
    if csv_file:
        with st.spinner("💫 Memproses CSV..."):
            try:
                gdf = load_csv_safe(csv_file)
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
