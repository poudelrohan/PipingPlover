"""
Step 4: Column Selection, Rename & Ordering
────────────────────────────────────────────
- Keeps only biologist-approved columns defined in config["columns_to_keep"]
- Enforces the column order defined in config["columns_to_keep"]
- Renames columns per config["column_rename"]
- Strips time component from date columns (date only, Excel date format)
- _removal_reason kept internally for Step 5/6 but not shown in final output

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

# ── Handle datetime columns ───────────────────────────────────────────────────
# Date columns: strip time component so Excel renders clean dates
# Time columns: extract just the time (HH:MM:SS), drop the dummy 1900 date
time_columns = {"StartTime", "EndTime"}
for col in df.columns:
    if pd.api.types.is_datetime64_any_dtype(df[col]):
        if col in time_columns:
            df[col] = pd.to_datetime(df[col]).dt.strftime("%H:%M:%S")
            df[col] = df[col].replace("NaT", None)
            print(f"  Extracted time from '{col}' → time only")
        else:
            df[col] = pd.to_datetime(df[col]).dt.normalize()
            print(f"  Stripped time from '{col}' → date only")

# ── Build final column list (order from config) ────────────────────────────────
keep = config["columns_to_keep"]

# Warn if any config column doesn't exist in the data
missing_cols = [c for c in keep if c not in df.columns]
if missing_cols:
    print(f"\n  [WARNING] These columns from config not found in data: {missing_cols}")

# Follow config order exactly, only include columns that exist
keep_ordered = [c for c in keep if c in df.columns]

# Always append internal tracking columns at the very end for pipeline use
for internal_col in ["_removal_reason", "_geo_warning", "_geo_correction"]:
    if internal_col in df.columns and internal_col not in keep_ordered:
        keep_ordered.append(internal_col)

# Log what's being dropped
all_cols = list(df.columns)
dropped  = [c for c in all_cols if c not in keep_ordered]
print(f"\n  Columns dropped ({len(dropped)}): {dropped}")
print(f"  Columns kept   ({len(keep_ordered)}): {keep_ordered}")

# ── Select & reorder columns ───────────────────────────────────────────────────
df = df[keep_ordered]

# ── Rename columns ─────────────────────────────────────────────────────────────
renames = config.get("column_rename", {})
valid_renames = {old: new for old, new in renames.items() if old in df.columns}
if valid_renames:
    df = df.rename(columns=valid_renames)
    print(f"\n  Renamed columns: {valid_renames}")

# ── Save output ────────────────────────────────────────────────────────────────
df.to_excel(output_path, index=False)

print(f"\n[DONE] Step 4 complete")
print(f"  Rows          : {len(df)}")
print(f"  Final columns : {list(df.columns)}")
print(f"  Output        : {output_path}")
