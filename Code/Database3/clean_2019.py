"""
Clean 2019 Winter Bird Survey — filter to PIPL rows and columns only
─────────────────────────────────────────────────────────────────────
2019 is a Google Form export (sheet: "Form Responses 1").
Same wide format as 2018 — one row per transect submission.

Structure:
  Cols 0-1   : Blank + Email (dropped)
  Col  2     : Transect name (Google Form dropdown)
  Cols 3-14  : Route metadata (date, observers, new-route GPS, time, weather, comments)
  Cols 15-61 : Route-level species totals (like DS2)
  Cols 62+   : Per-group blocks for groups 1-19:
                 Point N Latitude, N Longitude,
                 N Number of Piping Plovers,
                 N Band/Flag Codes for Piping Plovers (free text)
                 + other focal species cols (dropped)

What we keep:
  - Route metadata (transect, date, observers, start/end GPS, time, weather, rain, comments)
  - Route-total Piping Plover column
  - Per group 1-19: Latitude, Longitude, PIPL count, PIPL band codes
  - Only rows (transects) with any PIPL > 0

Known issues (do NOT fix here — log for pipeline):
  - Cedar Key West: route total=3 but no group GPS entries
  - Hillsboro Inlet: group sum=10 but route total left blank
  - New Smyrna Beach: two identical submissions (same date, same count)
  - Route names have ", County=X" appended from Google Form dropdown

Output: Databases/Database3Clean/Winter Birds 2019 Clean.xlsx
"""

import pandas as pd
import os

script_dir  = os.path.dirname(os.path.abspath(__file__))
input_path  = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3/Winter Birds 2019.xlsx"))
output_path = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3Clean/Winter Birds 2019 Clean.xlsx"))

# ── Load ───────────────────────────────────────────────────────────────────────
fr = pd.read_excel(input_path, sheet_name="Form Responses 1", header=0)
print(f"Loaded {len(fr)} rows, {len(fr.columns)} columns")

# ── Route metadata columns to keep ────────────────────────────────────────────
meta_cols = [
    fr.columns[2],   # Transect name
    fr.columns[3],   # Date
    fr.columns[4],   # Observers
    fr.columns[5],   # New route start lat
    fr.columns[6],   # New route start long
    fr.columns[7],   # New route end lat
    fr.columns[8],   # New route end long
    fr.columns[9],   # Route start time
    fr.columns[10],  # Weather temperature
    fr.columns[11],  # Wind
    fr.columns[12],  # Rain
    fr.columns[13],  # Survey type
    fr.columns[14],  # Additional comments
]

# Route-total PIPL column
route_pipl_col = fr.columns[28]  # "Piping Plover  (target species...)"

# ── Per-group PIPL columns (groups 1-19) ─────────────────────────────────────
group_cols = []
for g in range(1, 20):
    lat_matches  = [c for c in fr.columns if f"Point {g} Latitude"    in str(c)]
    lon_matches  = [c for c in fr.columns if f"Point {g} Longitude"   in str(c)]
    pipl_matches = [c for c in fr.columns if f"{g} Number of Piping"  in str(c)]
    band_matches = [c for c in fr.columns if f"{g} Band/Flag Codes for Piping" in str(c)]
    for matches in [lat_matches, lon_matches, pipl_matches, band_matches]:
        if matches:
            group_cols.append(matches[0])

keep_cols = meta_cols + [route_pipl_col] + group_cols
keep_cols = [c for c in keep_cols if c in fr.columns]   # guard against missing

df = fr[keep_cols].copy()

# ── Filter to rows with any PIPL ──────────────────────────────────────────────
pipl_cols_all = [c for c in df.columns if "Piping" in str(c) or "PIPL" in str(c)]
pipl_sum = df[pipl_cols_all].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
df_pipl = df[pipl_sum > 0].copy()

print(f"Rows with PIPL: {len(df_pipl)} out of {len(df)}")
print(f"Columns kept: {len(df_pipl.columns)}")
print(f"  Metadata     : {len(meta_cols)}")
print(f"  Route total  : 1")
print(f"  Group cols   : {len(group_cols)} (19 groups × up to 4 cols each)")

# ── Save ───────────────────────────────────────────────────────────────────────
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df_pipl.to_excel(writer, sheet_name="Form Responses 1", index=False)

print(f"\n[DONE] Saved to: {output_path}")
