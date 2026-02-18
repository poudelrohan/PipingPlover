"""
Step 1: Add Unique IDs
──────────────────────
Reads all input Excel files defined in the config, combines them into
one DataFrame, and adds:
  - unique_id       : sequential integer (1, 2, 3 ...)
  - source_database : human-readable database label (from config)
  - source_file     : name of the Excel file the row came from
  - source_sheet    : name of the sheet the row came from

Output: database4_with_ids.xlsx
"""

import pandas as pd
import os
import sys
import glob

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database4_config import config

# ── Resolve paths ──────────────────────────────────────────────────────────────
script_dir    = os.path.dirname(os.path.abspath(__file__))
input_folder  = os.path.normpath(os.path.join(script_dir, config["input"]["folder"]))
output_folder = os.path.normpath(os.path.join(script_dir, config["output"]["folder"]))
header_row    = config["input"]["header_row"]

os.makedirs(output_folder, exist_ok=True)

# ── Clear output folder before each run ───────────────────────────────────────
# Prevents stale files from previous runs mixing with current results
old_files = glob.glob(os.path.join(output_folder, "*.xlsx"))
if old_files:
    for f in old_files:
        os.remove(f)
    print(f"  Cleared {len(old_files)} old output file(s) from previous run")

# ── Read all input files ───────────────────────────────────────────────────────
all_frames = []

for filename in config["input"]["files"]:
    filepath = os.path.join(input_folder, filename)

    if not os.path.exists(filepath):
        print(f"  [ERROR] File not found: {filepath}")
        sys.exit(1)

    xl = pd.ExcelFile(filepath)

    for sheet in xl.sheet_names:
        print(f"  Reading '{filename}' → sheet '{sheet}' ...")
        df = pd.read_excel(filepath, sheet_name=sheet, header=header_row)

        df["source_database"] = config["source_database"]
        df["source_file"]     = filename
        df["source_sheet"]    = sheet

        all_frames.append(df)
        print(f"    {len(df)} rows loaded")

if not all_frames:
    print("[ERROR] No data loaded. Check config input files.")
    sys.exit(1)

# ── Combine and add unique_id ──────────────────────────────────────────────────
combined = pd.concat(all_frames, ignore_index=True)
combined.insert(0, "unique_id", range(1, len(combined) + 1))

# ── Save output ────────────────────────────────────────────────────────────────
output_path = os.path.join(output_folder, "database4_with_ids.xlsx")
combined.to_excel(output_path, index=False)

print(f"\n[DONE] Step 1 complete")
print(f"  Total rows : {len(combined)}")
print(f"  Columns    : {list(combined.columns)}")
print(f"  Output     : {output_path}")
