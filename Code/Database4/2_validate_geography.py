"""
Step 2: Geographic Validation
──────────────────────────────
Reads database4_with_ids.xlsx and validates that each row falls within
the Florida state boundary (from a Census TIGER shapefile) plus a
configurable coastal buffer (default 1 km).

Five cases handled:
  1. Sign error detected, corrected point inside FL   → fix value, keep + log
  2. Lat AND Lon present, inside Florida proper       → keep
  3. Lat AND Lon present, outside FL but within buffer → keep + warning
  4. Lat AND Lon present, outside buffer               → remove + distance from land
  5. Lat OR Lon missing                                → passed through (Step 3)

Output: database4_geo_validated.xlsx
"""

import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
import os
import sys

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database4_config import config

# ── Resolve paths ──────────────────────────────────────────────────────────────
script_dir    = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.normpath(os.path.join(script_dir, config["output"]["folder"]))

input_path   = os.path.join(output_folder, "database4_with_ids.xlsx")
output_path  = os.path.join(output_folder, "database4_geo_validated.xlsx")
removed_path = os.path.join(output_folder, "database4_removed.xlsx")

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 1 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 1")

# ── Geography config ───────────────────────────────────────────────────────────
geo     = config["geography"]
lat_col = geo["lat_column"]
lon_col = geo["lon_column"]

# ── Load Florida boundary ────────────────────────────────────────────────────
shp_path = os.path.normpath(os.path.join(script_dir, geo["shapefile"]))
states   = gpd.read_file(shp_path)
florida  = states[states["STATEFP"] == geo["state_fips"]]

if florida.empty:
    print(f"[ERROR] Could not find state with FIPS={geo['state_fips']} in shapefile")
    sys.exit(1)

# Build polygons:
#   fl_polygon    = exact Florida boundary (WGS84, no buffer)
#   fl_buffered   = Florida + buffer (WGS84)
#   fl_utm_polygon = exact Florida boundary (UTM, for distance calc)
fl_wgs         = florida.to_crs(epsg=4326)
fl_polygon     = fl_wgs.union_all()

fl_utm         = florida.to_crs(epsg=32617)
fl_utm_polygon = fl_utm.union_all()
fl_buf_geom    = fl_utm.buffer(geo["buffer_meters"])
fl_buf_wgs     = gpd.GeoSeries(fl_buf_geom, crs="EPSG:32617").to_crs(epsg=4326)
fl_buffered    = fl_buf_wgs.union_all()

print(f"  Florida boundary loaded with {geo['buffer_meters']}m coastal buffer")

# ── Identify rows with coordinates ───────────────────────────────────────────
has_coords = df[lat_col].notna() & df[lon_col].notna()

# ── Auto-fix sign errors ─────────────────────────────────────────────────────
# Biologists sometimes forget the negative sign on longitude or accidentally
# negate latitude. Try correcting and check if the point falls inside Florida.
corrections = []

for idx in df[has_coords].index:
    lat = df.at[idx, lat_col]
    lon = df.at[idx, lon_col]
    corrected_lat = lat
    corrected_lon = lon
    fixes = []

    # Longitude should be negative for Florida (-80 to -87)
    if lon > 0 and 79 <= lon <= 88:
        corrected_lon = -lon
        fixes.append(f"Longitude sign corrected: {lon} \u2192 {corrected_lon}")

    # Latitude should be positive for Florida (24 to 31)
    if lat < 0 and -31 <= lat <= -24:
        corrected_lat = -lat
        fixes.append(f"Latitude sign corrected: {lat} \u2192 {corrected_lat}")

    if fixes:
        pt = Point(corrected_lon, corrected_lat)
        if fl_polygon.contains(pt) or fl_buffered.contains(pt):
            df.at[idx, lon_col] = corrected_lon
            df.at[idx, lat_col] = corrected_lat
            df.at[idx, "_geo_correction"] = "; ".join(fixes)
            corrections.append({"unique_id": df.at[idx, "unique_id"], "correction": "; ".join(fixes)})

if corrections:
    print(f"  {len(corrections)} coordinate sign error(s) auto-corrected")

# ── Three-way classification ─────────────────────────────────────────────────
# Now runs on potentially-corrected values
inside_fl     = pd.Series(False, index=df.index)
inside_buffer = pd.Series(False, index=df.index)

for idx in df[has_coords].index:
    pt = Point(df.at[idx, lon_col], df.at[idx, lat_col])
    if fl_polygon.contains(pt):
        inside_fl.at[idx] = True
    elif fl_buffered.contains(pt):
        inside_buffer.at[idx] = True

# Flag rows in the buffer zone with a warning
buffer_count = inside_buffer.sum()
if buffer_count > 0:
    df.loc[inside_buffer, "_geo_warning"] = df.loc[inside_buffer].apply(
        lambda r: f"Within {geo['buffer_meters']}m coastal buffer (lat={r[lat_col]}, lon={r[lon_col]})",
        axis=1
    )
    print(f"  {buffer_count} row(s) kept with coastal buffer warning")

# Rows that have coords but are outside both FL and the buffer
outside_mask = has_coords & ~inside_fl & ~inside_buffer

# ── Build removal reason with distance from land ─────────────────────────────
for idx in df[outside_mask].index:
    lat = df.at[idx, lat_col]
    lon = df.at[idx, lon_col]

    # Project point to UTM for meter-accurate distance
    pt_gdf = gpd.GeoSeries([Point(lon, lat)], crs="EPSG:4326").to_crs(epsg=32617)
    dist_m = pt_gdf.iloc[0].distance(fl_utm_polygon)
    dist_km = round(dist_m / 1000, 1)

    df.at[idx, "_removal_reason"] = (
        f"Outside Florida boundaries (lat={lat}, lon={lon}) "
        f"\u2014 {dist_km} km from land"
    )

# ── Split clean vs removed ─────────────────────────────────────────────────────
removed = df[outside_mask].copy()
clean   = df[~outside_mask].copy()

# ── Save removed rows (append to existing removed file if it exists) ───────────
if len(removed) > 0:
    if os.path.exists(removed_path):
        existing = pd.read_excel(removed_path)
        removed  = pd.concat([existing, removed], ignore_index=True)
    removed.to_excel(removed_path, index=False)

# ── Save clean output ──────────────────────────────────────────────────────────
clean.to_excel(output_path, index=False)

print(f"\n[DONE] Step 2 complete")
print(f"  Rows kept    : {len(clean)}")
print(f"  Rows removed : {len(outside_mask[outside_mask])}")
print(f"  Output       : {output_path}")
