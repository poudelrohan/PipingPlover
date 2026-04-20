"""
Clean 2020 Winter Bird Survey — filter to PIPL rows and columns only
─────────────────────────────────────────────────────────────────────
2020 splits into two useful sheets:

  "All Species"         — one row per transect, route-level species totals
                          (equivalent of DS2 / 2018-2019 route total section)
  "Focal Observations"  — one row per transect, per-group detail for focal
                          species across 19 GPS points
                          (equivalent of DS1 / 2018-2019 group section)

Both sheets use the same transect names, so they can be joined later.
"Directions" and "Summary Data" sheets are dropped.

What we keep:
  All Species  → metadata cols + Piping Plover total; rows with PIPL > 0
  Focal        → Transect + County + per-group Lat/Long/PIPL count/PIPL bands;
                 rows with any PIPL > 0; all other species cols dropped

Known issues (do NOT fix here — log for pipeline):
  - Anastasia State Park Beach: route total=1 but no focal group GPS detail
  - Ponce Inlet Parks shorelines: route total=3 but focal group sum=5 (+2)
  - "Z My transect is not listed": catch-all row, route total=1, no GPS detail

Output: Databases/Database3Clean/Winter Birds 2020 Clean.xlsx
"""

import pandas as pd
import os

script_dir  = os.path.dirname(os.path.abspath(__file__))
input_path  = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3/Winter Birds 2020.xlsx"))
output_path = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3Clean/Winter Birds 2020 Clean.xlsx"))

# ── All Species sheet ─────────────────────────────────────────────────────────
all_sp = pd.read_excel(input_path, sheet_name="All Species", header=0)
all_sp = all_sp[all_sp["Transect "] != "TOTALS"].copy()
all_sp = all_sp[all_sp["Transect "].notna()].copy()

meta_cols_sp = [
    "Transect ", "County", "Date you did your survey",
    "Names of observers", "Email of primary observer",
    "Route Start Time", "Weather:  temperature (optional)",
    "Wind (optional)", "Rain (optional)",
    "Did you do a complete survey (all birds) or a focal survey? (focal species are h",
]
pipl_sp_col = [c for c in all_sp.columns if "Piping" in str(c)][0]
comments_sp = [c for c in all_sp.columns if "comment" in str(c).lower()]
keep_sp = [c for c in meta_cols_sp + [pipl_sp_col] + comments_sp if c in all_sp.columns]

all_sp_clean = all_sp[keep_sp].copy()
pipl_mask_sp = all_sp_clean[pipl_sp_col].apply(pd.to_numeric, errors="coerce").fillna(0) > 0
all_sp_pipl  = all_sp_clean[pipl_mask_sp].copy()
print(f"All Species: {len(all_sp)} rows → {len(all_sp_pipl)} PIPL rows, {len(all_sp_pipl.columns)} cols")

# ── Focal Observations sheet ──────────────────────────────────────────────────
# Row 0 = group number labels, Row 1 = actual column headers, Row 2+ = data
focal_raw = pd.read_excel(input_path, sheet_name="Focal Observations", header=None)
focal_raw.columns = focal_raw.iloc[1]
focal = focal_raw.iloc[2:].reset_index(drop=True)
transect_col = focal.columns[0]
focal = focal[focal[transect_col].notna()].copy()

# Keep Transect + County + per-group Lat/Long/PIPL count/PIPL bands
base_cols = [focal.columns[0], focal.columns[1]]   # Transect, County
group_cols = []
for g in range(1, 20):
    lat_m  = [c for c in focal.columns if f"Latitude (point {g})"              in str(c)]
    lon_m  = [c for c in focal.columns if f"Longitude (point {g})"             in str(c)]
    pipl_m = [c for c in focal.columns if f"Number of Piping Plovers (point {g})" in str(c)]
    band_m = [c for c in focal.columns if f"PIPL Band/Flag Codes (point {g})"  in str(c)]
    for matches in [lat_m, lon_m, pipl_m, band_m]:
        if matches:
            group_cols.append(matches[0])

keep_focal = [c for c in list(base_cols) + group_cols if c in focal.columns]
focal_clean = focal[keep_focal].copy()

pipl_focal_cols = [c for c in focal_clean.columns if "Piping" in str(c)]
pipl_sum = focal_clean[pipl_focal_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
focal_pipl = focal_clean[pipl_sum > 0].copy()
print(f"Focal Obs:   {len(focal)} rows → {len(focal_pipl)} PIPL rows, {len(focal_pipl.columns)} cols")

# ── Save ───────────────────────────────────────────────────────────────────────
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    all_sp_pipl.to_excel(writer, sheet_name="All Species",        index=False)
    focal_pipl.to_excel( writer, sheet_name="Focal Observations", index=False)

print(f"\n[DONE] Saved to: {output_path}")
