"""
Step 5: Duplicate Removal
──────────────────────────
Identifies duplicates based on config["duplicate_criteria"]:
  Route + Latitude + Longitude + SurveyDate + TotalObserved + FlagCode + FlagColor + BandCombo

Rules:
  - First occurrence kept, subsequent removed
  - Null values in criteria treated as equal
  - Criteria columns not present in this year's data are silently skipped

Output: db3_<year>_clean.xlsx  (in Output/<year>/)
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

year = config["year"]

input_path   = os.path.join(output_folder, get_filename("columns_selected"))
output_path  = os.path.join(output_folder, get_filename("clean"))
removed_path = os.path.join(output_folder, get_filename("removed"))

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 4 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 4 ({year})")

# ── Map duplicate criteria (apply any column renames) ─────────────────────────
renames           = config.get("column_rename", {})
criteria_original = config["duplicate_criteria"]
criteria          = [renames.get(c, c) for c in criteria_original]

if not criteria:
    print("  [WARNING] No duplicate_criteria defined in config. Skipping.")
    df.to_excel(output_path, index=False)
    sys.exit(0)

missing = [c for c in criteria if c not in df.columns]
if missing:
    print(f"  [NOTE] Criteria columns not in this year's data (skipping): {missing}")
criteria = [c for c in criteria if c in df.columns]

if not criteria:
    print("  [WARNING] No usable duplicate criteria after filtering. Skipping.")
    df.to_excel(output_path, index=False)
    sys.exit(0)

# Ensure _removal_reason column exists with correct dtype
if "_removal_reason" in df.columns:
    df["_removal_reason"] = df["_removal_reason"].astype("object")
else:
    df["_removal_reason"] = None

# ── Find duplicates ────────────────────────────────────────────────────────────
PLACEHOLDER = "__NULL__"
df_check = df[criteria].fillna(PLACEHOLDER).astype(str)

is_duplicate = df_check.duplicated(keep="first")

first_occurrence_map = {}
for idx, row in df_check.iterrows():
    key = tuple(row)
    if key not in first_occurrence_map:
        first_occurrence_map[key] = df.at[idx, "unique_id"]

removed_indices = []
for idx in df[is_duplicate].index:
    key      = tuple(df_check.loc[idx])
    first_id = first_occurrence_map[key]
    df.at[idx, "_removal_reason"] = f"Duplicate of unique_id: {first_id}"
    removed_indices.append(idx)

# ── Split clean vs removed ─────────────────────────────────────────────────────
removed = df.loc[removed_indices].copy()
clean   = df.drop(index=removed_indices).copy()

if "_removal_reason" in clean.columns:
    clean = clean.drop(columns=["_removal_reason"])

# ── Save removed rows ──────────────────────────────────────────────────────────
if len(removed) > 0:
    if os.path.exists(removed_path):
        existing = pd.read_excel(removed_path)
        removed  = pd.concat([existing, removed], ignore_index=True)
    removed.to_excel(removed_path, index=False)

clean.to_excel(output_path, index=False)

print(f"\n[DONE] Step 5 complete ({year})")
print(f"  Rows kept    : {len(clean)}")
print(f"  Rows removed : {len(removed_indices)}")
print(f"  Output       : {output_path}")
