"""
Step 2: Geographic Validation
──────────────────────────────
Reads database4_with_ids.xlsx and validates that each row falls within
the Florida bounding box defined in the config.

Three cases handled:
  1. Lat AND Lon present, within bounds  → keep
  2. Lat AND Lon present, outside bounds → remove, log reason
  3. Lat OR Lon missing                  → handled in Step 3 (required fields)
                                           passed through here untouched

Removal reason: "Outside Florida boundaries (lat=X, lon=Y)"

Output: database4_geo_validated.xlsx  (clean rows only)
        removed log appended to df["_removal_reason"] on removed rows
"""

import pandas as pd
import os
import sys

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database4_config import config

# ── Resolve paths ──────────────────────────────────────────────────────────────
script_dir    = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.normpath(os.path.join(script_dir, config["output"]["folder"]))

input_path  = os.path.join(output_folder, "database4_with_ids.xlsx")
output_path = os.path.join(output_folder, "database4_geo_validated.xlsx")
removed_path = os.path.join(output_folder, "database4_removed.xlsx")

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 1 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 1")

# ── Geography config ───────────────────────────────────────────────────────────
geo        = config["geography"]
lat_col    = geo["lat_column"]
lon_col    = geo["lon_column"]
lat_min    = geo["lat_min"]
lat_max    = geo["lat_max"]
lon_min    = geo["lon_min"]
lon_max    = geo["lon_max"]

# ── Validate ───────────────────────────────────────────────────────────────────
# Only check rows where both lat and lon are present
has_coords = df[lat_col].notna() & df[lon_col].notna()

in_bounds = (
    (df[lat_col] >= lat_min) & (df[lat_col] <= lat_max) &
    (df[lon_col] >= lon_min) & (df[lon_col] <= lon_max)
)

# Rows that have coords but are outside Florida
outside_mask = has_coords & ~in_bounds

# Build removal reason string for flagged rows
df.loc[outside_mask, "_removal_reason"] = (
    "Outside Florida boundaries (lat=" +
    df.loc[outside_mask, lat_col].astype(str) + ", lon=" +
    df.loc[outside_mask, lon_col].astype(str) + ")"
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
