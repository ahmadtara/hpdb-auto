import os
import streamlit as st
import pandas as pd
import geopandas as gpd
import ezdxf
from shapely.geometry import Point, box, LineString, MultiLineString
from shapely.ops import unary_union, linemerge, snap, polygonize
import gdown  # pip install gdown

TARGET_EPSG = "EPSG:32760"
DEFAULT_WIDTH = 10

# ---------------------
# Batas Pekanbaru (lon/lat EPSG:4326)
# ---------------------
PEKANBARU_BBOX = {
    "min_lon": 101.35,
    "max_lon": 101.55,
    "min_lat": 0.50,
    "max_lat": 0.65
}

# =====================
# Load Google Open Buildings CSV via local file or Google Drive link
# =====================
def load_google_buildings(csv_path_or_url):
    # Download dari Google Drive jika url
    if csv_path_or_url.startswith("http"):
        temp_file = "/tmp/building.csv.gz"
        gdown.download(csv_path_or_url, temp_file, quiet=False)
        csv_path_or_url = temp_file

    chunksize = 200_000
    gdf_list = []

    for chunk in pd.read_csv(csv_path_or_url, compression="gzip", chunksize=chunksize):
        # Pastikan kolom lon/lat
        if 'x' in chunk.columns and 'y' in chunk.columns:
            chunk = chunk.rename(columns={'x':'longitude', 'y':'latitude'})
        elif 'longitude' not in chunk.columns or 'latitude' not in chunk.columns:
            continue

        # Filter Pekanbaru
        chunk = chunk[
            (chunk.longitude >= PEKANBARU_BBOX["min_lon"]) &
            (chunk.longitude <= PEKANBARU_BBOX["max_lon"]) &
            (chunk.latitude >= PEKANBARU_BBOX["min_lat"]) &
            (chunk.latitude <= PEKANBARU_BBOX["max_lat"])
        ]
        if chunk.empty:
            continue

        gdf_chunk = gpd.GeoDataFrame(
            chunk,
            geometry=gpd.points_from_xy(chunk.longitude, chunk.latitude),
            crs="EPSG:4326"
        )
        gdf_list.append(gdf_chunk)

    if not gdf_list:
        raise Exception("❌ CSV tidak ada data valid untuk Pekanbaru")

    gdf = pd.concat(gdf_list, ignore_index=True)
    gdf = gdf.to_crs(TARGET_EPSG)
    return gdf

# =====================
# Buat jalan dari point
# =====================
def create_roads_from_points(gdf, width=DEFAULT_WIDTH):
    lines = []
    for pt in gdf.geometry:
        x, y = pt.x, pt.y
        line_h = LineString([(x - width, y), (x + width, y)])
        line_v = LineString([(x, y - width), (x, y + width)])
        lines.extend([line_h, line_v])
    roads_gdf = gpd.GeoDataFrame(geometry=lines, crs=gdf.crs)
    roads = roads_gdf.unary_union
    if isinstance(roads, LineString):
        roads = MultiLineString([roads])
    return gpd.GeoDataFrame(geometry=[roads], crs=gdf.crs)

# =====================
# Strip Z
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
# Export DXF
# =====================
def export_to_dxf(gdf_roads, gdf_buildings, dxf_path):
    doc = ezdxf.new()
    msp = doc.modelspace()
    # ROADS
    all_buffers = []
    for _, row in gdf_roads.iterrows():
        geom = strip_z(row.geometry)
        if geom.is_empty:
            continue
        buffered = geom.buffer(DEFAULT_WIDTH / 2, resolution=8, join_style=2)
        all_buffers.append(buffered)
    if all_buffers:
        all_union = unary_union(all_buffers)
        outlines = list(polygonize(all_union.boundary))
        bounds = [(pt[0], pt[1]) for geom in outlines for pt in geom.exterior.coords]
        min_x = min(x for x, y in bounds)
        min_y = min(y for x, y in bounds)
        for outline in outlines:
            coords = [(pt[0]-min_x, pt[1]-min_y) for pt in outline.exterior.coords]
            msp.add_lwpolyline(coords, dxfattribs={"layer": "ROADS"})
    else:
        min_x = min_y = 0

    # BUILDINGS
    for geom in gdf_buildings.geometry:
        if geom.is_empty:
            continue
        minx, miny, maxx, maxy = geom.bounds
        rect = box(minx, miny, maxx, maxy)
        coords = [(x - min_x, y - min_y) for x, y in rect.exterior.coords]
        msp.add_lwpolyline(coords, dxfattribs={"layer": "BUILDINGS"})

    doc.set_modelspace_vport(height=10000)
    doc.saveas(dxf_path)

# =====================
# Streamlit App
# =====================
def run_app():
    st.title("📦 Google Open Buildings → Jalan & Bangunan DXF")
    st.markdown("""
    ✅ Bisa upload CSV atau pakai link **Google Drive** (.csv.gz)<br>
    ✅ Filter otomatis: hanya data Pekanbaru<br>
    ✅ Jalan dibuat otomatis dari point, bangunan sebagai kotak bounding box.<br>
    """, unsafe_allow_html=True)

    csv_file = st.file_uploader("Upload CSV Google Open Buildings", type=["csv.gz"])
    drive_url = st.text_input("Atau masukkan link Google Drive (shareable link)")

    if csv_file or drive_url:
        with st.spinner("💫 Memproses..."):
            try:
                if csv_file:
                    temp_csv = f"/tmp/{csv_file.name}"
                    with open(temp_csv, "wb") as f:
                        f.write(csv_file.read())
                    gdf_buildings = load_google_buildings(temp_csv)
                else:
                    gdf_buildings = load_google_buildings(drive_url)

                gdf_roads = create_roads_from_points(gdf_buildings)
                dxf_path = "/tmp/buildings_roads.dxf"
                export_to_dxf(gdf_roads, gdf_buildings, dxf_path)

                st.success("✅ DXF berhasil dibuat!")
                with open(dxf_path, "rb") as f:
                    st.download_button("⬇️ Download DXF", data=f, file_name="buildings_roads.dxf")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {e}")

if __name__ == "__main__":
    run_app()
