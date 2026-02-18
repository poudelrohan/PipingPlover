"""
Step 2: Geographic Validation
──────────────────────────────
Validates that each row falls within the Florida bounding box defined in config.

Three cases handled:
  1. Lat AND Lon present, within bounds  → keep
  2. Lat AND Lon present, outside bounds → remove, log reason
  3. Lat OR Lon missing                  → passed through (handled in Step 3)

Removal reason: "Outside Florida boundaries (lat=X, lon=Y)"

Output: database1_geo_validated.xlsx
"""

import pandas as pd
import os
import sys

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database1_config import config

# ── Resolve paths ──────────────────────────────────────────────────────────────
script_dir    = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.normpath(os.path.join(script_dir, config["output"]["folder"]))

input_path   = os.path.join(output_folder, "database1_with_ids.xlsx")
output_path  = os.path.join(output_folder, "database1_geo_validated.xlsx")
removed_path = os.path.join(output_folder, "database1_removed.xlsx")

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

# ── Validate ───────────────────────────────────────────────────────────────────
has_coords = df[lat_col].notna() & df[lon_col].notna()

in_bounds = (
    (df[lat_col] >= geo["lat_min"]) & (df[lat_col] <= geo["lat_max"]) &
    (df[lon_col] >= geo["lon_min"]) & (df[lon_col] <= geo["lon_max"])
)

outside_mask = has_coords & ~in_bounds

df.loc[outside_mask, "_removal_reason"] = df.loc[outside_mask].apply(
    lambda r: f"Outside Florida boundaries (lat={r[lat_col]}, lon={r[lon_col]})", axis=1
)

# ── Split clean vs removed ─────────────────────────────────────────────────────
removed = df[outside_mask].copy()
clean   = df[~outside_mask].copy()

# ── Save removed rows ──────────────────────────────────────────────────────────
if len(removed) > 0:
    if os.path.exists(removed_path):
        existing = pd.read_excel(removed_path)
        removed  = pd.concat([existing, removed], ignore_index=True)
    removed.to_excel(removed_path, index=False)

clean.to_excel(output_path, index=False)

print(f"\n[DONE] Step 2 complete")
print(f"  Rows kept    : {len(clean)}")
print(f"  Rows removed : {len(removed[removed['_removal_reason'].notna()]) if len(removed) > 0 else 0}")
print(f"  Output       : {output_path}")
