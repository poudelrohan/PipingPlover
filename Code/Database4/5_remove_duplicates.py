"""
Step 5: Duplicate Removal
──────────────────────────
Identifies duplicates based on config["duplicate_criteria"].
  - Latitude + Longitude + ResightDate + FlagCode

Rules:
  - First occurrence of a duplicate group → kept
  - Subsequent occurrences → removed
  - Null values in criteria columns are treated as equal
    (two rows both missing FlagCode at same location/date = duplicate)

Removal reason: "Duplicate of unique_id: <id of first occurrence>"

Output: database4_clean.xlsx
        Removed rows appended to database4_removed.xlsx
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

input_path   = os.path.join(output_folder, "database4_columns_selected.xlsx")
output_path  = os.path.join(output_folder, "database4_clean.xlsx")
removed_path = os.path.join(output_folder, "database4_removed.xlsx")

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 4 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 4")

# ── Duplicate criteria ─────────────────────────────────────────────────────────
criteria = config["duplicate_criteria"]

if not criteria:
    print("  [WARNING] No duplicate_criteria defined in config. Skipping duplicate removal.")
    df.to_excel(output_path, index=False)
    sys.exit(0)

# Check all criteria columns exist (account for renames)
missing = [c for c in criteria if c not in df.columns]
if missing:
    print(f"  [ERROR] Duplicate criteria columns not found in data: {missing}")
    print(f"          Available columns: {list(df.columns)}")
    sys.exit(1)

# ── Find duplicates ────────────────────────────────────────────────────────────
# Use a placeholder for NaN so groupby treats nulls as equal
PLACEHOLDER = "__NULL__"
df_check = df[criteria].fillna(PLACEHOLDER)

# Mark duplicates — keep first occurrence
is_duplicate = df_check.duplicated(keep="first")

# For each duplicate, find the unique_id of the first occurrence it matches
first_occurrence_map = {}
for idx, row in df_check.iterrows():
    key = tuple(row)
    if key not in first_occurrence_map:
        first_occurrence_map[key] = df.at[idx, "unique_id"]

# Assign removal reasons
removed_indices = []
for idx in df[is_duplicate].index:
    key = tuple(df_check.loc[idx])
    first_id = first_occurrence_map[key]
    df.at[idx, "_removal_reason"] = f"Duplicate of unique_id: {first_id}"
    removed_indices.append(idx)

# ── Split clean vs removed ─────────────────────────────────────────────────────
removed = df.loc[removed_indices].copy()
clean   = df.drop(index=removed_indices).copy()

# Drop the internal _removal_reason column from clean data
if "_removal_reason" in clean.columns:
    clean = clean.drop(columns=["_removal_reason"])

# ── Save removed rows (append to existing removed file) ───────────────────────
if len(removed) > 0:
    if os.path.exists(removed_path):
        existing = pd.read_excel(removed_path)
        removed  = pd.concat([existing, removed], ignore_index=True)
    removed.to_excel(removed_path, index=False)

# ── Save clean output ──────────────────────────────────────────────────────────
clean.to_excel(output_path, index=False)

print(f"\n[DONE] Step 5 complete")
print(f"  Rows kept    : {len(clean)}")
print(f"  Rows removed : {len(removed_indices)}")
print(f"  Output       : {output_path}")
