import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
import math
from shapely.geometry import Point, LineString
from statistics import mean

# ----------------------------
# Config / Transformer
# ----------------------------
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

target_folders = {
    'FDT', 'FAT', 'HP COVER', 'HP UNCOVER', 'NEW POLE 7-3', 'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY', 'DISTRIBUTION CABLE', 'SLING WIRE', 'KOTAK', 'JALAN'
}

# ----------------------------
# KML / KMZ parsing
# ----------------------------
def extract_kmz(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kmz_file.extractall(extract_dir)
    return os.path.join(extract_dir, "doc.kml")

def parse_kml(kml_path):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    with open(kml_path, 'rb') as f:
        tree = ET.parse(f)
    root = tree.getroot()
    folders = root.findall('.//kml:Folder', ns)
    items = []
    for folder in folders:
        folder_name_tag = folder.find('kml:name', ns)
        if folder_name_tag is None:
            continue
        folder_name = folder_name_tag.text.strip().upper()
        if folder_name not in target_folders:
            continue
        placemarks = folder.findall('.//kml:Placemark', ns)
        for pm in placemarks:
            name = pm.find('kml:name', ns)
            name_text = name.text.strip() if name is not None else ""

            point_coord = pm.find('.//kml:Point/kml:coordinates', ns)
            if point_coord is not None and point_coord.text and point_coord.text.strip():
                lon, lat, *_ = point_coord.text.strip().split(',')
                items.append({
                    'type': 'point',
                    'name': name_text,
                    'latitude': float(lat),
                    'longitude': float(lon),
                    'folder': folder_name
                })
                continue

            line_coord = pm.find('.//kml:LineString/kml:coordinates', ns)
            if line_coord is not None and line_coord.text and line_coord.text.strip():
                coords = []
                for c in line_coord.text.strip().split():
                    lon, lat, *_ = c.split(',')
                    coords.append((float(lat), float(lon)))
                items.append({
                    'type': 'path',
                    'name': name_text,
                    'coords': coords,
                    'folder': folder_name
                })
                continue

            poly_coord = pm.find('.//kml:Polygon//kml:coordinates', ns)
            if poly_coord is not None and poly_coord.text and poly_coord.text.strip():
                coords = []
                for c in poly_coord.text.strip().split():
                    lon, lat, *_ = c.split(',')
                    coords.append((float(lat), float(lon)))
                items.append({
                    'type': 'path',
                    'name': name_text,
                    'coords': coords,
                    'folder': folder_name
                })
    return items

# ----------------------------
# Utilities
# ----------------------------
def latlon_to_xy(lat, lon):
    return transformer.transform(lon, lat)

def apply_offset(points_xy):
    xs = [x for x, y in points_xy]
    ys = [y for x, y in points_xy]
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    return [(x - cx, y - cy) for x, y in points_xy], (cx, cy)

def classify_items(items):
    classified = {name: [] for name in [
        "FDT", "FAT", "HP_COVER", "HP_UNCOVER",
        "NEW_POLE_7_3", "NEW_POLE_7_4", "EXISTING_POLE", "POLE",
        "BOUNDARY", "DISTRIBUTION_CABLE", "SLING_WIRE", "KOTAK", "JALAN"
    ]}
    for it in items:
        folder = it['folder']
        if "FDT" in folder:
            classified["FDT"].append(it)
        elif "FAT" in folder and folder != "FAT AREA":
            classified["FAT"].append(it)
        elif "HP COVER" in folder:
            classified["HP_COVER"].append(it)
        elif "HP UNCOVER" in folder:
            classified["HP_UNCOVER"].append(it)
        elif "NEW POLE 7-3" in folder:
            classified["NEW_POLE_7_3"].append(it)
        elif "NEW POLE 7-4" in folder:
            classified["NEW_POLE_7_4"].append(it)
        elif "EXISTING POLE" in folder or "EMR" in folder:
            classified["EXISTING_POLE"].append(it)
        elif "BOUNDARY" in folder:
            classified["BOUNDARY"].append(it)
        elif "DISTRIBUTION CABLE" in folder:
            classified["DISTRIBUTION_CABLE"].append(it)
        elif "SLING WIRE" in folder:
            classified["SLING_WIRE"].append(it)
        elif "KOTAK" in folder:
            classified["KOTAK"].append(it)
        elif "JALAN" in folder:
            classified["JALAN"].append(it)
        else:
            classified["POLE"].append(it)
    return classified

def segment_angle_xy(p1, p2):
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    ang = math.degrees(math.atan2(dy, dx))
    if ang <= -180:
        ang += 360
    if ang > 180:
        ang -= 360
    return ang

# ----------------------------
# Cable helpers
# ----------------------------
def build_cable_lines(classified):
    from shapely.ops import nearest_points
    cables = []
    for item in classified.get("DISTRIBUTION_CABLE", []):
        if item['type'] != 'path' or not item.get('coords'):
            continue
        xy = [latlon_to_xy(lat, lon) for lat, lon in item['coords']]
        if len(xy) >= 2:
            cables.append({'orig': item, 'xy_path': xy, 'line': LineString(xy)})
    return cables

def nearest_segment_angle_with_minlen(pt_xy, cable_line: LineString, min_seg_len):
    coords = list(cable_line.coords)
    px, py = pt_xy
    best_dist = float("inf")
    best_angle = None
    for i in range(len(coords) - 1):
        p1, p2 = coords[i], coords[i+1]
        seg = LineString([p1, p2])
        seg_len = seg.length
        if seg_len < min_seg_len:
            continue
        dist = seg.distance(Point(px, py))
        if dist < best_dist:
            best_dist = dist
            best_angle = segment_angle_xy(p1, p2)
    if best_angle is not None:
        return best_angle
    best_dist = float("inf")
    for i in range(len(coords) - 1):
        p1, p2 = coords[i], coords[i+1]
        seg = LineString([p1, p2])
        dist = seg.distance(Point(px, py))
        if dist < best_dist:
            best_dist = dist
            best_angle = segment_angle_xy(p1, p2)
    return best_angle if best_angle is not None else 0.0

# ----------------------------
# Group HP
# ----------------------------
def group_hp_by_cable_and_along(hp_xy_list, cables, max_gap_along=20.0):
    per_cable_hp = {i: [] for i in range(len(cables))}
    hp_meta = []
    for idx, hp in enumerate(hp_xy_list):
        pt = Point(hp['xy'])
        best_c = None
        best_dist = float("inf")
        best_along = None
        for c_idx, c in enumerate(cables):
            line = c['line']
            d = line.distance(pt)
            if d < best_dist:
                best_dist = d
                best_c = c_idx
                best_along = line.project(pt)
        if best_c is None:
            best_c, best_along = 0, 0.0
        per_cable_hp[best_c].append((idx, best_along))
        hp_meta.append({'cable_idx': best_c, 'along': best_along, 'dist_to_cable': best_dist})

    groups = []
    for c_idx, hp_list in per_cable_hp.items():
        if not hp_list: continue
        hp_list.sort(key=lambda x: x[1])
        current_group = [hp_list[0][0]]
        last_along = hp_list[0][1]
        group_alongs = [last_along]
        for idx_i, along_i in hp_list[1:]:
            if abs(along_i - last_along) > max_gap_along:
                groups.append({'cable_idx': c_idx,'indices': current_group.copy(),'along_vals': group_alongs.copy()})
                current_group = [idx_i]
                group_alongs = [along_i]
            else:
                current_group.append(idx_i)
                group_alongs.append(along_i)
            last_along = along_i
        groups.append({'cable_idx': c_idx,'indices': current_group.copy(),'along_vals': group_alongs.copy()})
    return groups, hp_meta

# ----------------------------
# Main DXF builder (UPDATED: includes automatic HP UNCOVER box + ANSI31 hatch)
# ----------------------------
def build_dxf_with_smart_hp(classified, template_path, output_path,
                            min_seg_len=15.0, max_gap_along=20.0,
                            rotate_hp=True, uncover_box_size=10.0):
    # load template or create new
    if template_path and os.path.exists(template_path):
        doc = ezdxf.readfile(template_path)
    else:
        doc = ezdxf.new('R2010')
    msp = doc.modelspace()

    block_mapping = {
        "FDT": "FDT",
        "FAT": "FAT",
        "POLE": "A$C14dd5346",
        "NEW_POLE_7_3": "A$C14dd5346",
        "NEW_POLE_7_4": "np9",
        "EXISTING_POLE": "A$Cdb6fd7d1"
    }

    # --- detect block references ---
    matchblock_fat = None
    matchblock_fdt = None
    matchblock_pole = None

    # scan block definitions in template
    for b in doc.blocks:
        try:
            bname = b.name.upper()
        except Exception:
            continue
        if not matchblock_fat and "FAT" in bname:
            matchblock_fat = b
        if not matchblock_fdt and "FDT" in bname:
            matchblock_fdt = b
        if not matchblock_pole and ("POLE" in bname or bname.startswith("A$") or "NEW_POLE" in bname or "EXISTING_POLE" in bname):
            matchblock_pole = b

    # collect XY coords and offset them
    all_xy = []
    for _, cat_items in classified.items():
        for obj in cat_items:
            if obj['type'] == 'point':
                all_xy.append(latlon_to_xy(obj['latitude'], obj['longitude']))
            elif obj['type'] == 'path':
                all_xy.extend([latlon_to_xy(lat, lon) for lat, lon in obj['coords']])
    if not all_xy:
        st.error("❌ Tidak ada data dari KMZ!")
        return None
    shifted_all, _ = apply_offset(all_xy)
    idx = 0
    for _, cat_items in classified.items():
        for obj in cat_items:
            if obj['type'] == 'point':
                obj['xy'] = shifted_all[idx]; idx += 1
            elif obj['type'] == 'path':
                obj['xy_path'] = shifted_all[idx: idx + len(obj['coords'])]; idx += len(obj['coords'])

    layer_mapping = {
        "BOUNDARY": "FAT AREA",
        "DISTRIBUTION_CABLE": "FO 36 CORE",
        "SLING_WIRE": "STRAND UG",
        "KOTAK": "GARIS HOMEPASS",
        "JALAN": "JALAN"
    }

    cables = build_cable_lines(classified)
    hp_items = []
    for obj in classified.get("HP_COVER", []) + classified.get("HP_UNCOVER", []):
        if 'xy' in obj:
            hp_items.append({'obj': obj, 'xy': obj['xy']})

    groups, hp_meta = group_hp_by_cable_and_along(hp_items, cables, max_gap_along=max_gap_along)
    for group in groups:
        c_idx = group['cable_idx']
        if c_idx >= len(cables):
            continue
        line = cables[c_idx]['line']
        rep_along = sorted(group['along_vals'])[len(group['along_vals'])//2]
        rep_point = line.interpolate(rep_along)
        angle = nearest_segment_angle_with_minlen((rep_point.x, rep_point.y), line, min_seg_len)
        for hp_idx in group['indices']:
            hp_items[hp_idx]['rotation'] = angle
    for hp in hp_items:
        if 'rotation' not in hp:
            best_angle, best_dist = 0.0, float("inf")
            for c in cables:
                dist = c['line'].distance(Point(hp['xy']))
                if dist < best_dist:
                    best_dist = dist
                    best_angle = nearest_segment_angle_with_minlen(hp['xy'], c['line'], min_seg_len)
            hp['rotation'] = best_angle

    # Gambar polyline & block lain
    for layer_name, cat_items in classified.items():
        true_layer = layer_mapping.get(layer_name, layer_name)
        for obj in cat_items:
            if obj['type'] == "path":
                if len(obj.get('xy_path', [])) >= 2:
                    msp.add_lwpolyline(obj['xy_path'], dxfattribs={"layer": true_layer})
                elif len(obj.get('xy_path', [])) == 1:
                    msp.add_circle(center=obj['xy_path'][0], radius=0.5, dxfattribs={"layer": true_layer})

    # ----------------------------
    # AUTO HATCH HP UNCOVER (buat kotak di titik HP UNCOVER lalu hatch ANSI31)
    # ----------------------------
    # uncover_box_size adalah panjang sisi kotak dalam meter (nilai diambil setelah apply_offset => unit sama)
    half_s = uncover_box_size / 2.0
    for obj in classified.get("HP_UNCOVER", []):
        if obj['type'] == "point" and 'xy' in obj:
            x, y = obj['xy']
            s = half_s
            # buat polygon kotak (tutup ulang)
            square = [
                (x - s, y - s),
                (x + s, y - s),
                (x + s, y + s),
                (x - s, y + s),
                (x - s, y - s)
            ]
            try:
            
                # tambahkan hatch ANSI31 di layer HP UNCOVER
                hatch = msp.add_hatch(dxfattribs={"layer": "HP UNCOVER"})
                # set pattern ANSI31; scale dan angle dapat disesuaikan jika perlu
                try:
                    hatch.set_pattern_fill("ANSI31", scale=11.0000, angle=0)
                except Exception:
                    # fallback: gunakan solid jika pattern tidak tersedia
                    try:
                        hatch.set_solid_fill()
                    except Exception:
                        pass
                # tambahkan path polygon ke hatch
                # ezdxf expects a list of points for add_polyline_path
                try:
                    hatch.paths.add_polyline_path(square, is_closed=True)
                except Exception:
                    # beberapa versi ezdxf mungkin butuh format lain, fallback ke edge path
                    try:
                        hatch.paths.add_edge_path([("L", (square[0], square[1], square[2], square[3]))])
                    except Exception:
                        pass
            except Exception as e:
                st.warning(f"Gagal buat kotak/hatch HP UNCOVER di ({x:.3f},{y:.3f}): {e}")

    # Gambar point/blocks/text lain (kecuali HP_COVER / HP_UNCOVER yang sudah diproses)
    for layer_name, cat_items in classified.items():
        true_layer = layer_mapping.get(layer_name, layer_name)
        if layer_name in ("HP_COVER", "HP_UNCOVER"):
            continue
        for obj in cat_items:
            if obj['type'] != "point":
                continue
            x, y = obj['xy']
            block_name = None
            # prefer using block_mapping for new/existing poles
            if layer_name == "FAT" and matchblock_fat:
                block_name = matchblock_fat.name if hasattr(matchblock_fat, "name") else (matchblock_fat.dxf.name if hasattr(matchblock_fat, "dxf") else None)
            elif layer_name == "FDT" and matchblock_fdt:
                block_name = matchblock_fdt.name if hasattr(matchblock_fdt, "name") else (matchblock_fdt.dxf.name if hasattr(matchblock_fdt, "dxf") else None)
            elif layer_name in ["NEW_POLE_7_3", "NEW_POLE_7_4", "EXISTING_POLE", "POLE"]:
                block_name = block_mapping.get(layer_name)

            if block_name:
                try:
                    if layer_name in ["FDT", "FAT"]:
                        scale = 0.0025
                    elif layer_name in ["NEW_POLE_7_3", "POLE", "EXISTING_POLE"]:
                        scale = 1.0
                    elif layer_name in ["NEW_POLE_7_4"]:
                        scale = 0.001
                    else:
                        scale = 1.0

                    msp.add_blockref(
                        block_name, (x, y),
                        dxfattribs={
                            "layer": true_layer,
                            "xscale": scale,
                            "yscale": scale,
                            "zscale": scale
                        }
                    )
                except Exception as e:
                    st.warning(f"Gagal insert block {block_name}: {e}")
                    # fallback: circle if block insertion failed
                    msp.add_circle(center=(x, y), radius=2, dxfattribs={"layer": true_layer})
            else:
                # fallback kalau tidak ada block
                msp.add_circle(center=(x, y), radius=2, dxfattribs={"layer": true_layer})

            # Tambah teks untuk point (FDT, FAT, POLE, dll.)
            text_layer = "FEATURE_LABEL" if "POLE" in layer_name else true_layer

            # Atur warna teks sesuai kategori
            if layer_name == "FAT":
                color_val = 2   # kuning
            elif layer_name in ["FDT", "NEW_POLE_7_3", "NEW_POLE_7_4", "POLE"]:
                color_val = 1   # merah
            elif layer_name == "EXISTING_POLE":
                color_val = 7   # putih
            else:
                color_val = 256 # bylayer

            msp.add_text(
                obj.get("name", ""),
                dxfattribs={
                    "height": 5.0 if layer_name in ["FDT", "FAT", "NEW_POLE_7_3", "NEW_POLE_7_4", "EXISTING_POLE"] else 1.5,
                    "layer": text_layer,
                    "color": color_val,
                    "insert": (x + 2, y)
                }
            )

    # ----------------------------
    # Draw HP teks di titik tengah (dengan opsi rotasi)
    # ----------------------------
    for hp in hp_items:
        x, y = hp['xy']
        # Sudut rotasi dari kabel
        rot_deg = hp['rotation'] if rotate_hp else 0

        # Koreksi orientasi agar teks tidak terbalik di AutoCAD
        rot_deg = (rot_deg + 360) % 360  # normalisasi 0–360

        # Jika miring terbalik (menghadap bawah), balik 180°
        if 90 < rot_deg < 270:
            rot_deg = (rot_deg + 180) % 360

        rot = math.radians(rot_deg)
        name = hp['obj'].get("name", "")

        h = 6 if "HP COVER" in hp['obj']['folder'] else 3
        c = 6 if "HP COVER" in hp['obj']['folder'] else 7

        # Estimasi lebar teks
        text_width = len(name) * h * 0.6

        text = msp.add_text(
            name,
            dxfattribs={
                "layer": "FEATURE_LABEL",
                "color": c,
                "height": h,
            }
        )
        text.dxf.rotation = float(rot_deg)
        try:
            text.dxf.halign = 1   # center
            text.dxf.valign = 2   # middle
        except Exception:
            pass
        text.dxf.insert = (float(x), float(y))
        try:
            text.dxf.align_point = (float(x), float(y))
        except Exception:
            pass

    # save dxf
    doc.saveas(output_path)
    return output_path

# ----------------------------
# Streamlit UI
# ----------------------------
def run_kmz_to_dwg():
    st.title("🏗️ KMZ → AUTOCAD (Smart HP rotation + block insertion + Auto Hatch HP UNCOVER)")
    uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
    uploaded_template = st.file_uploader("📀 Upload Template DXF (optional)", type=["dxf"])
    st.sidebar.header("Rotation parameters")
    min_seg_len = st.sidebar.slider("Min seg length (m)", 5.0, 100.0, 15.0, 1.0)
    max_gap_along = st.sidebar.slider("Max gap along (m)", 5.0, 200.0, 20.0, 1.0)
    rotate_hp = st.sidebar.checkbox("Rotate HP Text", value=False)
    uncover_box_size = st.sidebar.slider("Ukuran kotak HP UNCOVER (m)", 1.0, 100.0, 10.0, 1.0)

    if uploaded_kmz:
        tmpdir = "temp_extract"; os.makedirs(tmpdir, exist_ok=True)
        kmz_path = os.path.join(tmpdir, "uploaded.kmz")
        with open(kmz_path, "wb") as f: f.write(uploaded_kmz.read())
        kml_path = extract_kmz(kmz_path, tmpdir)
        items = parse_kml(kml_path)
        classified = classify_items(items)

        template_path = None
        if uploaded_template:
            template_path = os.path.join(tmpdir, "template.dxf")
            with open(template_path, "wb") as f: f.write(uploaded_template.read())

        out_path = "output_smart_hp.dxf"
        res = build_dxf_with_smart_hp(
            classified, template_path, out_path,
            min_seg_len=min_seg_len, max_gap_along=max_gap_along,
            rotate_hp=rotate_hp, uncover_box_size=uncover_box_size
        )
        if res:
            with open(res, "rb") as f:
                st.download_button("⬇️ Download DXF", f, res)

if __name__ == "__main__":
    run_kmz_to_dwg()


