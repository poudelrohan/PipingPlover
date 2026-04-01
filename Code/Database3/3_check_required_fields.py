"""
Step 3: Required Fields & Location Check
─────────────────────────────────────────
Two checks run on every row:

  A) Location fields rule:
       - ALL of [Route, Latitude, Longitude] missing → remove row
       - At least ONE present but not all → keep row, warn

  B) Required fields rule:
       - SurveyDate missing/empty → remove row

Output: db3_<year>_required_fields_ok.xlsx  (in Output/<year>/)
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

year = config["active_year"]

input_path   = os.path.join(output_folder, get_filename("geo_validated"))
output_path  = os.path.join(output_folder, get_filename("required_fields_ok"))
removed_path = os.path.join(output_folder, get_filename("removed"))

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 2 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 2 ({year})")

removed_rows = []
warn_partial_location = []

# ── A) Location fields check ───────────────────────────────────────────────────
loc_config = config["location_fields"]
loc_fields = loc_config["fields"]

for idx, row in df.iterrows():
    missing = [f for f in loc_fields if pd.isna(row.get(f)) or str(row.get(f)).strip() == ""]
    if len(missing) == len(loc_fields):
        df.at[idx, "_removal_reason"] = loc_config["removal_reason"]
        removed_rows.append(idx)
    elif len(missing) > 0 and loc_config.get("warn_if_partial", False):
        warn_partial_location.append({
            "unique_id":      row.get("unique_id"),
            "missing_fields": missing
        })

# ── B) Required fields check ───────────────────────────────────────────────────
required = config["required_fields"]

for idx, row in df.iterrows():
    if idx in removed_rows:
        continue
    for field in required:
        if field not in df.columns:
            print(f"  [WARNING] Required field '{field}' not found in data columns — skipping check")
            continue
        if pd.isna(row[field]) or str(row[field]).strip() == "":
            df.at[idx, "_removal_reason"] = f"Missing required field: {field}"
            removed_rows.append(idx)
            break

# ── Split clean vs removed ─────────────────────────────────────────────────────
removed = df.loc[removed_rows].copy()
clean   = df.drop(index=removed_rows).copy()

# ── Save removed rows ─────────────────────────────────────────────────────────
if len(removed) > 0:
    if os.path.exists(removed_path):
        existing = pd.read_excel(removed_path)
        removed  = pd.concat([existing, removed], ignore_index=True)
    removed.to_excel(removed_path, index=False)

clean.to_excel(output_path, index=False)

print(f"\n[DONE] Step 3 complete ({year})")
print(f"  Rows kept    : {len(clean)}")
print(f"  Rows removed : {len(removed_rows)}")

if warn_partial_location:
    print(f"\n  [WARNING] {len(warn_partial_location)} rows have partial location data:")
    for w in warn_partial_location:
        print(f"    unique_id={w['unique_id']} — missing: {w['missing_fields']}")

print(f"\n  Output       : {output_path}")
