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
        "FDT", "FAT", "HP_COVER", "HP_UNCOVER", "NEW_POLE", "EXISTING_POLE", "POLE",
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
        elif "NEW POLE" in folder:
            classified["NEW_POLE"].append(it)
        elif "EXISTING" in folder or "EMR" in folder:
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

def find_nearest_cable_segment(pt_xy, cables, min_seg_len=1.0):
    """
    Untuk titik pt_xy cari kabel terdekat, kembalikan tuple:
    (best_cable_index, angle_degrees, projected_point (Point), distance)
    """
    px, py = pt_xy
    best_dist = float("inf")
    best_info = (None, 0.0, Point(px, py), best_dist)
    for idx, c in enumerate(cables):
        line = c['line']
        # jarak dan proyeksi
        dist = line.distance(Point(px, py))
        if dist < best_dist:
            # cari segmen terbaik di cable yang memperhatikan min_seg_len
            coords = list(line.coords)
            seg_best_dist = float("inf")
            seg_best_angle = 0.0
            seg_best_proj = None
            # proyeksi sepanjang line (global) - kita gunakan proyeksi pada line (bukan seg)
            proj_along = line.project(Point(px, py))
            seg_best_proj = line.interpolate(proj_along)
            # cari segmen terdekat (tanpa min len filter ketat, preferensi)
            for i in range(len(coords) - 1):
                a = coords[i]; b = coords[i+1]
                seg = LineString([a, b])
                d = seg.distance(Point(px, py))
                if d < seg_best_dist:
                    seg_best_dist = d
                    seg_best_angle = segment_angle_xy(a, b)
            if seg_best_proj is None:
                seg_best_proj = Point(px, py)
            best_dist = dist
            best_info = (idx, seg_best_angle, seg_best_proj, dist)
    return best_info  # (cable_idx, angle, projected_point, distance)

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
# Helper bbox & collision (approx)
# ----------------------------
def estimate_text_bbox(cx, cy, text, height, rotation_degrees):
    """
    Kira-kira axis-aligned bbox (approx) untuk teks.
    Lebih sederhana: lebar ~= len(text) * height * 0.6
    return (xmin, ymin, xmax, ymax)
    """
    w = max(1.0, len(text)) * height * 0.6
    h = height
    # axis-aligned approx (ignore rotation to keep simple)
    xmin = cx - w/2
    xmax = cx + w/2
    ymin = cy - h/2
    ymax = cy + h/2
    return (xmin, ymin, xmax, ymax)

def bbox_overlap(a, b):
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    return not (ax2 < bx1 or bx2 < ax1 or ay2 < by1 or by2 < ay1)

# ----------------------------
# Main DXF builder
# ----------------------------
def build_dxf_with_smart_hp(classified, template_path, output_path,
                            min_seg_len=15.0, max_gap_along=20.0,
                            text_offset=2.5, collision_step=1.0, collision_attempts=6):
    """
    text_offset: jarak (meter) untuk menggeser label tegak lurus dari proyeksi kabel
    collision_step: step perpindahan tambahan tiap percobaan (meter)
    collision_attempts: maksimum percobaan geser untuk menghindari collision
    """
    if template_path and os.path.exists(template_path):
        doc = ezdxf.readfile(template_path)
    else:
        doc = ezdxf.new('R2010')
    msp = doc.modelspace()

    # --- detect block references ---
    matchblock_fat = matchblock_fdt = matchblock_pole = None

    # cek di modelspace
    for e in msp:
        try:
            if e.dxftype() == "INSERT":
                name = e.dxf.name.upper()
                if "FAT" in name and matchblock_fat is None:
                    matchblock_fat = e
                elif "FDT" in name and matchblock_fdt is None:
                    matchblock_fdt = e
                elif "POLE" in name and matchblock_pole is None:
                    matchblock_pole = e
                elif name.startswith("A$") and matchblock_pole is None:
                    matchblock_pole = e
        except Exception:
            continue

    # cek juga di block definition
    for b in doc.blocks:
        bname = b.name.upper()
        if not matchblock_fat and "FAT" in bname:
            matchblock_fat = b
        if not matchblock_fdt and "FDT" in bname:
            matchblock_fdt = b
        if not matchblock_pole and "POLE" in bname:
            matchblock_pole = b

    # shift coords
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

    # mapping
    layer_mapping = {
        "BOUNDARY": "FAT AREA",
        "DISTRIBUTION_CABLE": "FO 36 CORE",
        "SLING_WIRE": "STRAND UG",
        "KOTAK": "GARIS HOMEPASS",
        "JALAN": "JALAN"
    }

    cables = build_cable_lines(classified)

    # HP rotation grouping (tetap dipisah)
    hp_items = []
    for obj in classified.get("HP_COVER", []) + classified.get("HP_UNCOVER", []):
        if 'xy' in obj: hp_items.append({'obj': obj, 'xy': obj['xy']})

    groups, hp_meta = group_hp_by_cable_and_along(hp_items, cables, max_gap_along=max_gap_along)
    for g_idx, group in enumerate(groups):
        if not cables: break
        c_idx = group['cable_idx']; line = cables[c_idx]['line']
        rep_along = sorted(group['along_vals'])[len(group['along_vals'])//2]
        rep_point = line.interpolate(rep_along)
        angle = nearest_segment_angle_with_minlen((rep_point.x, rep_point.y), line, min_seg_len)
        for hp_idx in group['indices']:
            hp_items[hp_idx]['rotation'] = angle
    for hp in hp_items:
        if 'rotation' not in hp and cables:
            best_angle, best_dist = 0.0, float("inf")
            for c in cables:
                dist = c['line'].distance(Point(hp['xy']))
                if dist < best_dist:
                    best_dist = dist
                    best_angle = nearest_segment_angle_with_minlen(hp['xy'], c['line'], min_seg_len)
            hp['rotation'] = best_angle

    # draw polylines
    for layer_name, cat_items in classified.items():
        true_layer = layer_mapping.get(layer_name, layer_name)
        for obj in cat_items:
            if obj['type']=="path":
                if len(obj.get('xy_path',[]))>=2:
                    msp.add_lwpolyline(obj['xy_path'], dxfattribs={"layer": true_layer})
                elif len(obj.get('xy_path',[]))==1:
                    msp.add_circle(center=obj['xy_path'][0], radius=0.5, dxfattribs={"layer": true_layer})

    # draw points (FAT/FDT/POLE) -> sekarang block + teks ROTATE mengikuti kabel
    placed_label_bboxes = []  # simpan bbox axis-aligned dari label yang sudah ditempatkan (approx)

    for layer_name, cat_items in classified.items():
        true_layer = layer_mapping.get(layer_name, layer_name)
        if layer_name in ("HP_COVER", "HP_UNCOVER"):
            continue
        for obj in cat_items:
            if obj['type'] != "point":
                continue
            x, y = obj['xy']
            block_name = None
            if layer_name == "FAT" and matchblock_fat:
                block_name = matchblock_fat.dxf.name if hasattr(matchblock_fat, "dxf") else matchblock_fat.name
            elif layer_name == "FDT" and matchblock_fdt:
                block_name = matchblock_fdt.dxf.name if hasattr(matchblock_fdt, "dxf") else matchblock_fdt.name
            elif layer_name in ["NEW_POLE", "EXISTING_POLE"] and matchblock_pole:
                block_name = matchblock_pole.dxf.name if hasattr(matchblock_pole, "dxf") else matchblock_pole.name
            elif layer_name == "POLE" and matchblock_pole:
                block_name = matchblock_pole.dxf.name if hasattr(matchblock_pole, "dxf") else matchblock_pole.name

            # cari kabel terdekat untuk menentukan rotasi (khusus distribution cable)
            rot_angle = 0.0
            proj_point = Point(x, y)
            if cables:
                c_idx, angle, proj_pt, dist = find_nearest_cable_segment((x, y), cables, min_seg_len=min_seg_len)
                if c_idx is not None:
                    rot_angle = angle
                    proj_point = proj_pt

            inserted = False
            if block_name:
                try:
                    scale = 1.0
                    if layer_name == "FDT":
                        scale = 0.0025
                    elif layer_name == "FAT":
                        scale = 0.0025
                    elif "POLE" in layer_name:
                        scale = 1.0

                    # tambahkan block dengan rotasi sesuai kabel
                    msp.add_blockref(
                        block_name, (x, y),
                        dxfattribs={
                            "layer": true_layer,
                            "xscale": scale,
                            "yscale": scale,
                            "zscale": scale,
                            "rotation": rot_angle
                        }
                    )
                    inserted = True
                except Exception as e:
                    st.warning(f"Gagal insert block {block_name}: {e}")
                    inserted = False

            if not inserted:
                msp.add_circle(center=(x, y), radius=2, dxfattribs={"layer": true_layer})

            # ----------------------------
            # Tambahkan teks yang ROTATE mengikuti kabel juga (khusus untuk POLE/FAT/FDT)
            # ----------------------------
            label_text = obj.get("name", "")
            # pilih tinggi teks
            text_height = 5.0 if layer_name in ["FDT", "FAT", "NEW_POLE", "EXISTING_POLE", "POLE"] else 1.5

            # starting label position: titik proyeksi pada kabel
            base_x = proj_point.x
            base_y = proj_point.y

            # perp unit vector (tegak lurus ke arah kabel)
            rad = math.radians(rot_angle)
            perp_x, perp_y = -math.sin(rad), math.cos(rad)

            # awal offset (tegak lurus) sesuai text_offset
            label_x = base_x + perp_x * text_offset
            label_y = base_y + perp_y * text_offset

            # collision avoidance: jika bbox overlap, geser +/- perp * (collision_step * attempt)
            bbox = estimate_text_bbox(label_x, label_y, label_text, text_height, rot_angle)
            attempt = 0
            direction = 1
            while any(bbox_overlap(bbox, other) for other in placed_label_bboxes) and attempt < collision_attempts:
                attempt += 1
                # berganti sisi setiap percobaan: +, -, +, -
                direction = -direction
                shift = (text_offset + collision_step * attempt) * direction
                label_x = base_x + perp_x * shift
                label_y = base_y + perp_y * shift
                bbox = estimate_text_bbox(label_x, label_y, label_text, text_height, rot_angle)

            # simpan bbox ke list (untuk cek label selanjutnya)
            placed_label_bboxes.append(bbox)

            # layer teks: pakai FEATURE_LABEL untuk pole, else gunakan true_layer
            text_layer = "FEATURE_LABEL" if "POLE" in layer_name or layer_name in ["NEW_POLE", "EXISTING_POLE"] else true_layer
            text_color = 1 if text_layer == "FEATURE_LABEL" else 256

            # tambahkan teks (rotate mengikuti kabel)
            msp.add_text(
                label_text,
                dxfattribs={
                    "height": text_height,
                    "layer": text_layer,
                    "color": text_color,
                    "insert": (label_x, label_y),
                    "rotation": rot_angle
                }
            )

    # draw HP (tetap seperti sebelumnya, terpisah)
    for hp in hp_items:
        x,y=hp['xy'];rot=hp['rotation'];name=hp['obj'].get("name","")
        if "HP COVER" in hp['obj']['folder']:
            msp.add_text(name,dxfattribs={"height":6,"layer":"FEATURE_LABEL","color":6,"insert":(x,y),"rotation":rot})
        else:
            msp.add_text(name,dxfattribs={"height":3,"layer":"FEATURE_LABEL","color":7,"insert":(x,y),"rotation":rot})

    doc.saveas(output_path)
    return output_path

# ----------------------------
# Streamlit UI
# ----------------------------
def run_kmz_to_dwg():
    st.title("🏗️ KMZ → AUTOCAD (Smart HP rotation + block insertion)")
    uploaded_kmz=st.file_uploader("📂 Upload File KMZ",type=["kmz"])
    uploaded_template=st.file_uploader("📀 Upload Template DXF (optional)",type=["dxf"])
    st.sidebar.header("Rotation / Label parameters")
    min_seg_len=st.sidebar.slider("Min seg length (m)",5.0,100.0,15.0,1.0)
    max_gap_along=st.sidebar.slider("Max gap along (m)",5.0,200.0,20.0,1.0)
    text_offset = st.sidebar.slider("Label offset from cable (m)", 0.0, 10.0, 2.5, 0.5)
    collision_step = st.sidebar.slider("Collision shift step (m)", 0.1, 5.0, 1.0, 0.1)
    collision_attempts = st.sidebar.slider("Max collision attempts", 0, 12, 6, 1)

    if uploaded_kmz:
        extract_dir="temp_kmz";os.makedirs(extract_dir,exist_ok=True)
        with open("temp_upload.kmz","wb") as f: f.write(uploaded_kmz.read())
        kml_path=extract_kmz("temp_upload.kmz",extract_dir)
        items=parse_kml(kml_path);classified=classify_items(items)
        template_path="template_ref.dxf"
        if uploaded_template:
            with open(template_path,"wb") as f: f.write(uploaded_template.read())
        else:
            template_path=template_path if os.path.exists(template_path) else ""
        output_dxf="converted_output.dxf"
        try:
            result=build_dxf_with_smart_hp(classified,template_path,output_dxf,
                                           min_seg_len=min_seg_len,max_gap_along=max_gap_along,
                                           text_offset=text_offset, collision_step=collision_step,
                                           collision_attempts=collision_attempts)
            if result and os.path.exists(result):
                st.success("✅ DXF berhasil dibuat.")
                with open(result,"rb") as f:
                    st.download_button("⬇️ Download DXF",f,file_name=os.path.basename(result))
            else: st.error("❌ Gagal membuat DXF.")
        except Exception as e:
            st.error(f"❌ Error: {e}")

if __name__=="__main__":
    run_kmz_to_dwg()
