"""
Step 2: Geographic Validation
──────────────────────────────
Validates that each row falls within a Florida lat/lon bounding box.

Four cases handled:
  1. Sign error detected, corrected point inside box → fix value, keep + log
  2. Lat AND Lon present, inside bounding box        → keep
  3. Lat AND Lon present, outside bounding box        → remove
  4. Lat OR Lon missing                               → passed through (Step 3)

Output: db3_<year>_geo_validated.xlsx  (in Output/<year>/)
"""

import pandas as pd
import os
import sys

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database3_config import config, get_output_folder, get_filename

# ── Resolve paths ──────────────────────────────────────────────────────────────
script_dir    = os.path.dirname(os.path.abspath(__file__))
output_folder = get_output_folder(script_dir)

input_path   = os.path.join(output_folder, get_filename("with_ids"))
output_path  = os.path.join(output_folder, get_filename("geo_validated"))
removed_path = os.path.join(output_folder, get_filename("removed"))

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 1 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
year = config["active_year"]
print(f"  Loaded {len(df)} rows from Step 1 ({year})")

# ── Geography config ───────────────────────────────────────────────────────────
geo     = config["geography"]
lat_col = geo["lat_column"]
lon_col = geo["lon_column"]
lat_min = geo["lat_min"]
lat_max = geo["lat_max"]
lon_min = geo["lon_min"]
lon_max = geo["lon_max"]

print(f"  Bounding box: lat [{lat_min}, {lat_max}], lon [{lon_min}, {lon_max}]")


def in_box(lat, lon):
    """Check if a point falls within the Florida bounding box."""
    return lat_min <= lat <= lat_max and lon_min <= lon <= lon_max


# ── Identify rows with coordinates ───────────────────────────────────────────
has_coords = df[lat_col].notna() & df[lon_col].notna()

# ── Auto-fix sign errors ─────────────────────────────────────────────────────
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
        fixes.append(f"Longitude sign corrected: {lon} → {corrected_lon}")

    # Latitude should be positive for Florida (24 to 31)
    if lat < 0 and -31 <= lat <= -24:
        corrected_lat = -lat
        fixes.append(f"Latitude sign corrected: {lat} → {corrected_lat}")

    if fixes and in_box(corrected_lat, corrected_lon):
        df.at[idx, lon_col] = corrected_lon
        df.at[idx, lat_col] = corrected_lat
        df.at[idx, "_geo_correction"] = "; ".join(fixes)
        corrections.append({"unique_id": df.at[idx, "unique_id"], "correction": "; ".join(fixes)})

if corrections:
    print(f"  {len(corrections)} coordinate sign error(s) auto-corrected")

# ── Bounding box check ──────────────────────────────────────────────────────
inside_box = pd.Series(False, index=df.index)

for idx in df[has_coords].index:
    lat = df.at[idx, lat_col]
    lon = df.at[idx, lon_col]
    if in_box(lat, lon):
        inside_box.at[idx] = True

# Rows that have coords but are outside the bounding box
outside_mask = has_coords & ~inside_box

# ── Build removal reason ─────────────────────────────────────────────────────
for idx in df[outside_mask].index:
    lat = df.at[idx, lat_col]
    lon = df.at[idx, lon_col]
    df.at[idx, "_removal_reason"] = (
        f"Outside Florida bounding box (lat={lat}, lon={lon})"
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

print(f"\n[DONE] Step 2 complete ({year})")
print(f"  Rows kept    : {len(clean)}")
print(f"  Rows removed : {len(outside_mask[outside_mask])}")
print(f"  Output       : {output_path}")
