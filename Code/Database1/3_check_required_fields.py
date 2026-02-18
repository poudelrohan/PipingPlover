"""
Step 3: Required Fields & Location Check
─────────────────────────────────────────
Three checks run on every row:

  A) Location fields rule (config["location_fields"]):
       - ALL of [LOCALITY, LATITUDE, LONGITUDE] missing → remove row
       - At least ONE present but not all → keep row, warn in Summary_Report
       Removal reason: "Missing all location fields (LOCALITY, LATITUDE, LONGITUDE)"

  B) Required fields rule (config["required_fields"]):
       - Any required field null/empty → remove row
       Removal reason: "Missing required field: <field_name>"

  C) Observation Count special rule (config["observation_count"]):
       - Value = "x" → keep (eBird convention: species present, count unknown)
       - Value = blank/null → remove row
       Removal reason: "Missing required field: OBSERVATION COUNT"

Output: database1_required_fields_ok.xlsx
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

input_path   = os.path.join(output_folder, "database1_geo_validated.xlsx")
output_path  = os.path.join(output_folder, "database1_required_fields_ok.xlsx")
removed_path = os.path.join(output_folder, "database1_removed.xlsx")

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 2 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 2")

removed_indices      = []
warn_partial_location = []

# ── A) Location fields check ───────────────────────────────────────────────────
loc_config = config["location_fields"]
loc_fields = loc_config["fields"]

for idx, row in df.iterrows():
    missing = [f for f in loc_fields if pd.isna(row.get(f)) or str(row.get(f)).strip() == ""]
    if len(missing) == len(loc_fields):
        df.at[idx, "_removal_reason"] = loc_config["removal_reason"]
        removed_indices.append(idx)
    elif len(missing) > 0 and loc_config.get("warn_if_partial", False):
        warn_partial_location.append({
            "unique_id":      row.get("unique_id"),
            "missing_fields": missing,
        })

# ── B) Required fields check ───────────────────────────────────────────────────
for idx, row in df.iterrows():
    if idx in removed_indices:
        continue
    for field in config["required_fields"]:
        if field not in df.columns:
            print(f"  [WARNING] Required field '{field}' not in data columns — skipping")
            continue
        if pd.isna(row[field]) or str(row[field]).strip() == "":
            df.at[idx, "_removal_reason"] = f"Missing required field: {field}"
            removed_indices.append(idx)
            break

# ── C) Observation Count special rule ─────────────────────────────────────────
obs_config = config["observation_count"]
obs_col    = obs_config["column"]
keep_val   = str(obs_config["keep_value"]).strip().lower()

if obs_col not in df.columns:
    print(f"  [WARNING] OBSERVATION COUNT column '{obs_col}' not found — skipping check")
else:
    for idx, row in df.iterrows():
        if idx in removed_indices:
            continue
        val = row[obs_col]
        if pd.isna(val) or str(val).strip() == "":
            # Blank = remove
            df.at[idx, "_removal_reason"] = obs_config["removal_reason"]
            removed_indices.append(idx)
        # "x" = keep silently, no flag needed

# ── Split clean vs removed ─────────────────────────────────────────────────────
removed = df.loc[removed_indices].copy()
clean   = df.drop(index=removed_indices).copy()

# ── Save removed rows ──────────────────────────────────────────────────────────
if len(removed) > 0:
    if os.path.exists(removed_path):
        existing = pd.read_excel(removed_path)
        removed  = pd.concat([existing, removed], ignore_index=True)
    removed.to_excel(removed_path, index=False)

clean.to_excel(output_path, index=False)

# ── Summary ────────────────────────────────────────────────────────────────────
print(f"\n[DONE] Step 3 complete")
print(f"  Rows kept    : {len(clean)}")
print(f"  Rows removed : {len(removed_indices)}")

if warn_partial_location:
    print(f"\n  [WARNING] {len(warn_partial_location)} rows have partial location data:")
    for w in warn_partial_location:
        print(f"    unique_id={w['unique_id']} — missing: {w['missing_fields']}")

print(f"\n  Output : {output_path}")
