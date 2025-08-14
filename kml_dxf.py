# =========================
# IMPORT LIBRARY
# =========================
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point, Polygon

# =========================
# DEKLARASI POLYGON PEKANBARU
# =========================
# Koordinat kasar Pekanbaru (lat/lon)
pekanbaru_poly = Polygon([
    (101.3, 0.3),   # SW
    (101.8, 0.3),   # SE
    (101.8, 0.6),   # NE
    (101.3, 0.6),   # NW
])

# =========================
# UPLOAD CSV BESAR
# =========================
# Misal di Jupyter: bisa drag & drop file CSV ke folder
csv_path = "31d_buildings.csv"  # Ganti dengan path file asli

# =========================
# FILTER CSV DENGAN CHUNK
# =========================
chunksize = 500_000  # Jumlah baris per chunk
filtered_chunks = []

for chunk in pd.read_csv(csv_path, chunksize=chunksize):
    # Pastikan kolom longitude & latitude ada
    if 'x' in chunk.columns and 'y' in chunk.columns:
        chunk = chunk.rename(columns={'x':'longitude', 'y':'latitude'})
    if 'longitude' not in chunk.columns or 'latitude' not in chunk.columns:
        raise Exception("CSV tidak memiliki kolom longitude/latitude")

    # Buat GeoDataFrame dari chunk
    gdf = gpd.GeoDataFrame(
        chunk,
        geometry=gpd.points_from_xy(chunk.longitude, chunk.latitude),
        crs="EPSG:4326"
    )

    # Filter hanya yang ada di Pekanbaru
    filtered_chunk = gdf[gdf.geometry.within(pekanbaru_poly)]
    if not filtered_chunk.empty:
        filtered_chunks.append(filtered_chunk)

# =========================
# CONCAT SEMUA CHUNK
# =========================
if filtered_chunks:
    gdf_pekanbaru = pd.concat(filtered_chunks)
else:
    gdf_pekanbaru = gpd.GeoDataFrame(columns=['longitude','latitude','geometry'], crs="EPSG:4326")

# =========================
# SAVE HASIL CSV
# =========================
output_csv = "31d_buildings_pekanbaru.csv"
gdf_pekanbaru.drop(columns="geometry").to_csv(output_csv, index=False)

print(f"✅ Selesai! File Pekanbaru tersimpan di: {output_csv}")
