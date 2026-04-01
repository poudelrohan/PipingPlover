"""
Step 1: Add Unique IDs
──────────────────────
Reads the extracted PIPL data from Step 0 and adds:
  - unique_id       : sequential integer (1, 2, 3 ...)
  - source_database : human-readable database label
  - source_file     : name of the Excel file the row came from
  - source_sheet    : name of the sheet the row came from

Also clears old output files (except Step 0) to prevent stale data.

Output: db3_<year>_with_ids.xlsx  (in Output/<year>/)
"""

import pandas as pd
import os
import sys
import glob

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database3_config import config, get_output_folder, get_filename

# ── Resolve paths ──────────────────────────────────────────────────────────────
script_dir    = os.path.dirname(os.path.abspath(__file__))
output_folder = get_output_folder(script_dir)

year     = config["active_year"]
year_cfg = config["years"][year]

input_path  = os.path.join(output_folder, get_filename("extracted"))
output_path = os.path.join(output_folder, get_filename("with_ids"))

# ── Clear old output files (except the extracted file from Step 0) ────────────
old_files = [f for f in glob.glob(os.path.join(output_folder, "db3_*.xlsx"))
             if not f.endswith(get_filename("extracted"))]
if old_files:
    for f in old_files:
        os.remove(f)
    print(f"  Cleared {len(old_files)} old output file(s) from previous run")

# ── Load data ──────────────────────────────────────────────────────────────────
if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    print("        Did you run Step 0 first?")
    sys.exit(1)

df = pd.read_excel(input_path)
print(f"  Loaded {len(df)} rows from Step 0")

# ── Add tracking columns ──────────────────────────────────────────────────────
df.insert(0, "unique_id", range(1, len(df) + 1))
df["source_database"] = config["source_database"]
df["source_file"]     = year_cfg["file"]
df["source_sheet"]    = year_cfg["sheet"]

# ── Save output ────────────────────────────────────────────────────────────────
df.to_excel(output_path, index=False)

print(f"\n[DONE] Step 1 complete ({year})")
print(f"  Total rows : {len(df)}")
print(f"  Columns    : {list(df.columns)}")
print(f"  Output     : {output_path}")
