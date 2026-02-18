"""
Step 4: Column Selection, Rename & Ordering
────────────────────────────────────────────
- Keeps only biologist-approved columns defined in config["columns_to_keep"]
- Enforces the column order defined in config["columns_to_keep"]
- Renames ALL CAPS eBird column names to readable titles (config["column_rename"])
- Strips time component from date columns (date only, Excel date format)
- _removal_reason kept internally for Step 5/6, not shown in final output

Output: database1_columns_selected.xlsx
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

input_path  = os.path.join(output_folder, "database1_required_fields_ok.xlsx")
output_path = os.path.join(output_folder, "database1_columns_selected.xlsx")

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 3 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 3")

# ── Strip time from all datetime columns ───────────────────────────────────────
# Normalize to midnight so Excel renders as a clean date (no time component shown)
# Keep as datetime64 — openpyxl writes it as a real Excel date cell
for col in df.columns:
    if pd.api.types.is_datetime64_any_dtype(df[col]):
        df[col] = pd.to_datetime(df[col]).dt.normalize()
        print(f"  Stripped time from '{col}' → date only")

# ── Build final column list (order from config) ────────────────────────────────
keep = config["columns_to_keep"]

missing_cols = [c for c in keep if c not in df.columns]
if missing_cols:
    print(f"\n  [WARNING] These columns from config not found in data: {missing_cols}")

keep_ordered = [c for c in keep if c in df.columns]

# Always append _removal_reason at end for internal pipeline use
if "_removal_reason" in df.columns and "_removal_reason" not in keep_ordered:
    keep_ordered.append("_removal_reason")

all_cols = list(df.columns)
dropped  = [c for c in all_cols if c not in keep_ordered]
print(f"\n  Columns dropped ({len(dropped)}): {dropped}")
print(f"  Columns kept   ({len(keep_ordered)}): {keep_ordered}")

# ── Select & reorder ───────────────────────────────────────────────────────────
df = df[keep_ordered]

# ── Rename columns ─────────────────────────────────────────────────────────────
renames = config.get("column_rename", {})
valid_renames = {old: new for old, new in renames.items() if old in df.columns}
if valid_renames:
    df = df.rename(columns=valid_renames)
    print(f"\n  Renamed {len(valid_renames)} columns")

# ── Save output ────────────────────────────────────────────────────────────────
df.to_excel(output_path, index=False)

print(f"\n[DONE] Step 4 complete")
print(f"  Rows          : {len(df)}")
print(f"  Final columns : {list(df.columns)}")
print(f"  Output        : {output_path}")
