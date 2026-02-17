"""
Step 4: Column Selection & Rename
───────────────────────────────────
Keeps only biologist-approved columns defined in config["columns_to_keep"].
Also renames columns per config["column_rename"].

Pipeline/tracking columns always kept regardless of config:
  - unique_id, source_file, source_sheet, _removal_reason

Logs which columns were dropped.

Output: database4_columns_selected.xlsx
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

input_path  = os.path.join(output_folder, "database4_required_fields_ok.xlsx")
output_path = os.path.join(output_folder, "database4_columns_selected.xlsx")

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 3 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 3")
print(f"  Columns before selection: {list(df.columns)}")

# ── Always keep these pipeline columns ────────────────────────────────────────
pipeline_cols = ["unique_id", "source_file", "source_sheet", "_removal_reason"]

# ── Build final column list ────────────────────────────────────────────────────
keep = config["columns_to_keep"]

# Warn if any config column doesn't exist in the data
missing_cols = [c for c in keep if c not in df.columns]
if missing_cols:
    print(f"\n  [WARNING] These columns from config not found in data: {missing_cols}")

# Only keep columns that actually exist
keep_existing = [c for c in keep if c in df.columns]

# Add pipeline cols (only if they exist)
for col in pipeline_cols:
    if col in df.columns and col not in keep_existing:
        keep_existing.append(col)

# Log what's being dropped
all_cols    = list(df.columns)
dropped     = [c for c in all_cols if c not in keep_existing]
print(f"\n  Columns dropped ({len(dropped)}): {dropped}")
print(f"  Columns kept   ({len(keep_existing)}): {keep_existing}")

# ── Select columns ─────────────────────────────────────────────────────────────
df = df[keep_existing]

# ── Rename columns ─────────────────────────────────────────────────────────────
renames = config.get("column_rename", {})
# Only rename columns that exist
valid_renames = {old: new for old, new in renames.items() if old in df.columns}
if valid_renames:
    df = df.rename(columns=valid_renames)
    print(f"\n  Renamed columns: {valid_renames}")

# ── Save output ────────────────────────────────────────────────────────────────
df.to_excel(output_path, index=False)

print(f"\n[DONE] Step 4 complete")
print(f"  Rows : {len(df)}")
print(f"  Final columns: {list(df.columns)}")
print(f"  Output : {output_path}")
