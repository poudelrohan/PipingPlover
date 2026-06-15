"""
Step 4: Column Selection & Ordering
────────────────────────────────────
- Keeps only columns defined in config["columns_to_keep"]
- Enforces column order
- Applies any renames from config["column_rename"]
- Strips time component from date columns

Output: db3_<year>_columns_selected.xlsx  (in Output/<year>/)
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

input_path  = os.path.join(output_folder, get_filename("required_fields_ok"))
output_path = os.path.join(output_folder, get_filename("columns_selected"))

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 3 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 3 ({year})")

# ── Strip time from datetime columns ──────────────────────────────────────────
for col in df.columns:
    if pd.api.types.is_datetime64_any_dtype(df[col]):
        df[col] = pd.to_datetime(df[col]).dt.normalize()
        print(f"  Stripped time from '{col}' → date only")

# ── Build final column list ───────────────────────────────────────────────────
keep = config["columns_to_keep"]

missing_cols = [c for c in keep if c not in df.columns]
if missing_cols:
    print(f"\n  [NOTE] Columns not in this year's data (OK): {missing_cols}")

keep_ordered = [c for c in keep if c in df.columns]

# Append internal tracking columns
for internal_col in ["_removal_reason", "_geo_warning", "_geo_correction"]:
    if internal_col in df.columns and internal_col not in keep_ordered:
        keep_ordered.append(internal_col)

dropped = [c for c in df.columns if c not in keep_ordered]
print(f"\n  Columns dropped ({len(dropped)}): {dropped}")
print(f"  Columns kept   ({len(keep_ordered)}): {keep_ordered}")

df = df[keep_ordered]

# ── Rename columns ─────────────────────────────────────────────────────────────
renames = config.get("column_rename", {})
valid_renames = {old: new for old, new in renames.items() if old in df.columns}
if valid_renames:
    df = df.rename(columns=valid_renames)
    print(f"\n  Renamed columns: {valid_renames}")

# ── Save output ────────────────────────────────────────────────────────────────
df.to_excel(output_path, index=False)

print(f"\n[DONE] Step 4 complete ({year})")
print(f"  Rows          : {len(df)}")
print(f"  Final columns : {list(df.columns)}")
print(f"  Output        : {output_path}")
