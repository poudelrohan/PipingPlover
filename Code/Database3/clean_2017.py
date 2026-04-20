"""
Clean 2017 Winter Bird Survey — filter to PIPL rows only
─────────────────────────────────────────────────────────
Reads the raw 2017 file and produces Winter Birds '17 Clean.xlsx with:
  - DS1: rows where focal species contains "PIPL" (66 → keep)
  - DS2: rows where Piping Plover count > 0 (route-level summaries)
  - DS3: rows where Species contains "PIPL" (47 → keep)

No data values are changed — this is purely a row/column filter.

Output: Databases/Database3Clean/Winter Birds '17 Clean.xlsx
"""

import pandas as pd
import os

script_dir  = os.path.dirname(os.path.abspath(__file__))
input_path  = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3/Winter Birds '17.xlsx"))
output_path = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3Clean/Winter Birds '17 Clean.xlsx"))

# ── DS1 ────────────────────────────────────────────────────────────────────────
# header_row=1 (row 0 = instructions, row 1 = column names)
# Keep only the 16 named columns; filter to rows with PIPL
ds1 = pd.read_excel(input_path, sheet_name="DATA SHEET 1", header=1)
named_cols_ds1 = [c for c in ds1.columns if not str(c).startswith("Unnamed")]
ds1 = ds1[named_cols_ds1]
focal_col = "Species and number of individuals"
ds1_pipl = ds1[ds1[focal_col].astype(str).str.contains("PIPL", case=False, na=False)].copy()
print(f"DS1: {len(ds1)} rows → {len(ds1_pipl)} PIPL rows kept")

# ── DS2 ────────────────────────────────────────────────────────────────────────
# header_row=0 (column names are the very first row — different from 2016!)
# Keep rows where Piping Plover count > 0
ds2 = pd.read_excel(input_path, sheet_name="DATA SHEET 2", header=0)
named_cols_ds2 = [c for c in ds2.columns if not str(c).startswith("Unnamed")]
ds2 = ds2[named_cols_ds2]
pipl_col = "Piping Plover"
ds2_pipl = ds2[pd.to_numeric(ds2[pipl_col], errors="coerce").fillna(0) > 0].copy()
print(f"DS2: {len(ds2)} rows → {len(ds2_pipl)} PIPL route rows kept")

# ── DS3 ────────────────────────────────────────────────────────────────────────
# header_row=1 (row 0 = instructions, row 1 = column names)
# Drop the 2 trailing empty Unnamed columns; filter to PIPL species rows
ds3 = pd.read_excel(input_path, sheet_name="DATA SHEET 3", header=1)
named_cols_ds3 = [c for c in ds3.columns if not str(c).startswith("Unnamed")]
ds3 = ds3[named_cols_ds3]
ds3_pipl = ds3[ds3["Species"].astype(str).str.contains("PIPL", case=False, na=False)].copy()
print(f"DS3: {len(ds3)} rows → {len(ds3_pipl)} PIPL rows kept")

# ── Write output ────────────────────────────────────────────────────────────────
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    ds1_pipl.to_excel(writer, sheet_name="DATA SHEET 1", index=False)
    ds2_pipl.to_excel(writer, sheet_name="DATA SHEET 2", index=False)
    ds3_pipl.to_excel(writer, sheet_name="DATA SHEET 3", index=False)

print(f"\n[DONE] Saved to: {output_path}")
