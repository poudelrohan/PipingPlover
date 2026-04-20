"""
Clean 2018 Winter Bird Survey — filter to PIPL rows and columns only
─────────────────────────────────────────────────────────────────────
2018 uses a completely different format from 2016/2017.
There is ONE sheet ('Sheet1') with ONE ROW PER TRANSECT (route).
All groups/points for a transect are spread across columns (wide format):

  Cols 0-13  : Route metadata (transect name, start/end GPS, date, observers,
               survey type, start time, weather, wind, comments, rain)
  Cols 14-59 : Route-level TOTALS for every focal species (equivalent of DS2).
               Includes PIPL (col 27) = total PIPL for the whole transect.
  Cols 60+   : Repeated column blocks per group/point (equivalent of DS1):
                 <N>Lat, <N>Long, <N>PIPL, <N>PIPLbands, <N>REKN, ...
               Groups 1-15, 17, 19 have PIPL columns.
               Group 16 exists but has NO PIPL column.
               Group 18 is entirely missing.
               Group 20 only has lat/long/nonfocals.

Note: band info is a COUNT only (no band combos like DS3 in 2016/2017).

What we keep:
  - All 14 route metadata columns
  - Route-total PIPL column
  - For each group with a PIPL column (1-15, 17, 19): Lat, Long, PIPL, PIPLbands
  - Only rows (transects) that have any PIPL > 0

Row 0 is the user's notes — dropped. Rows 1 = headers. Rows 2-3 = empty template rows.

Output: Databases/Database3Clean/Winter Birds 2018 Clean.xlsx
"""

import pandas as pd
import os

script_dir  = os.path.dirname(os.path.abspath(__file__))
input_path  = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3/Winter Birds 2018.xlsx"))
output_path = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3Clean/Winter Birds 2018 Clean.xlsx"))

# ── Load ───────────────────────────────────────────────────────────────────────
# Row 0 = notes, Row 1 = headers, Rows 2-3 = empty template rows
raw = pd.read_excel(input_path, sheet_name="Sheet1", header=None)
raw.columns = raw.iloc[1]
raw = raw.iloc[2:].reset_index(drop=True)           # drop notes + header rows
raw = raw[raw["Transect"].notna()].reset_index(drop=True)  # drop empty template rows
print(f"Loaded {len(raw)} transect rows")

# ── Define columns to keep ────────────────────────────────────────────────────

# Route metadata (cols 0-13)
meta_cols = [
    "Transect",
    "Start Lat", "Start Long", "End Lat", "End Long",
    "Date you did your survey",
    "Names of observers",
    "Lead observer's email",
    "Did you do a complete survey (all birds) or a focal survey?",
    "Route Start Time",
    "Weather: temperature (optional)",
    "Wind (optional)",
    "Any additional survey comments?",
    "Rain (optional)",
]

# Route-total PIPL (col 27)
route_pipl_col = "PIPL"

# Group-level PIPL columns — groups 1-15, 17, 19 have PIPL
# (group 16 has no PIPL col, group 18 missing, group 20 has no PIPL)
groups_with_pipl = list(range(1, 16)) + [17, 19]

group_cols = []
for g in groups_with_pipl:
    # Lat/Long column names vary slightly (e.g. "1Lat" vs "20lat")
    lat_name  = f"{g}Lat"
    lon_name  = f"{g}Long"
    pipl_name = f"{g}PIPL"
    band_name = f"{g}PIPLbands" if f"{g}PIPLbands" in raw.columns else f"{g}PIPL ands"
    for col in [lat_name, lon_name, pipl_name, band_name]:
        if col in raw.columns:
            group_cols.append(col)

keep_cols = [c for c in meta_cols + [route_pipl_col] + group_cols if c in raw.columns]
df = raw[keep_cols].copy()

# ── Filter to PIPL routes ──────────────────────────────────────────────────────
pipl_all_cols = [c for c in df.columns if "PIPL" in str(c)]
pipl_sum = df[pipl_all_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
has_pipl = pipl_sum > 0
df_pipl = df[has_pipl].copy()

print(f"Transects with PIPL: {len(df_pipl)} out of {len(df)}")
print(f"Columns kept: {len(df_pipl.columns)}")
print(f"  Metadata : {len(meta_cols)}")
print(f"  Route total PIPL : 1")
print(f"  Group-level cols : {len(group_cols)} ({len(groups_with_pipl)} groups × up to 4 cols each)")

# ── Save ───────────────────────────────────────────────────────────────────────
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df_pipl.to_excel(writer, sheet_name="Sheet1", index=False)

print(f"\n[DONE] Saved to: {output_path}")
print()
print("NOTE: When building the pipeline, this wide format must be 'melted'")
print("into long format (one row per group per transect), like 2016/2017.")
