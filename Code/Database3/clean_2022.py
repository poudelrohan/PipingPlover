"""
Clean 2022 Winter Bird Survey — filter to PIPL rows and columns only
─────────────────────────────────────────────────────────────────────
Same structure as 2020/2021: two sheets kept.
  "All Species"        — route totals (metadata + PIPL count)
  "Focal Observations" — per-group GPS detail (19 points, MOST IMPORTANT)

"For sites with >19 points" sheet exists but has no PIPL data — skipped.
"Directions", "Summary Data", "transects with no data 2.22.21" — skipped.

Known issues (do NOT fix here — log for pipeline):
  - "Totals" row appears in All Species with PIPL=288 — summary row, not a transect
  - AllSp=42 PIPL rows vs Focal=38 PIPL rows (4 routes have totals but no GPS groups)

Output: Databases/Database3Clean/Winter Birds 2022 Clean.xlsx
"""

import pandas as pd, os, warnings
warnings.filterwarnings('ignore')

script_dir  = os.path.dirname(os.path.abspath(__file__))
input_path  = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3/Winter Birds 2022.xlsx"))
output_path = os.path.normpath(os.path.join(script_dir, "../../Databases/Database3Clean/Winter Birds 2022 Clean.xlsx"))

# ── All Species ───────────────────────────────────────────────────────────────
all_sp = pd.read_excel(input_path, sheet_name="All Species", header=0)
tc_sp  = all_sp.columns[0]   # "Transect "
all_sp = all_sp[~all_sp[tc_sp].isin(["TOTALS", "Totals"])].dropna(subset=[tc_sp]).copy()

meta_cols = list(all_sp.columns[:10])   # Transect, County, Date, Observers, Email, Start Time, Temp, Wind, Rain, Survey type
pipl_col  = [c for c in all_sp.columns if "Piping" in str(c)][0]
comments  = [c for c in all_sp.columns if "comment" in str(c).lower()]
keep_sp   = [c for c in meta_cols + [pipl_col] + comments if c in all_sp.columns]

all_sp_pipl = all_sp[all_sp[pipl_col].apply(pd.to_numeric, errors="coerce").fillna(0) > 0][keep_sp].copy()
print(f"All Species:  {len(all_sp)} rows → {len(all_sp_pipl)} PIPL rows, {len(all_sp_pipl.columns)} cols")

# ── Focal Observations ────────────────────────────────────────────────────────
focal_raw = pd.read_excel(input_path, sheet_name="Focal Observations", header=None)
focal_raw.columns = focal_raw.iloc[1]
focal = focal_raw.iloc[2:].reset_index(drop=True)
tc_f  = focal.columns[0]
focal = focal[focal[tc_f].notna()].copy()

base  = [focal.columns[0], focal.columns[1]]   # Transect, County
group_cols = []
for g in range(1, 20):
    for pattern in [f"Latitude (point {g})", f"Longitude (point {g})",
                    f"Number of Piping Plovers (point {g})", f"PIPL Band/Flag Codes (point {g})"]:
        matches = [c for c in focal.columns if pattern in str(c)]
        if matches:
            group_cols.append(matches[0])

keep_f = [c for c in list(base) + group_cols if c in focal.columns]
focal_clean = focal[keep_f].copy()
pipl_f_cols = [c for c in focal_clean.columns if "Piping" in str(c)]
focal_pipl  = focal_clean[focal_clean[pipl_f_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) > 0].copy()
print(f"Focal Obs:    {len(focal)} rows → {len(focal_pipl)} PIPL rows, {len(focal_pipl.columns)} cols")

# ── Save ──────────────────────────────────────────────────────────────────────
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    all_sp_pipl.to_excel(writer, sheet_name="All Species",        index=False)
    focal_pipl.to_excel( writer, sheet_name="Focal Observations", index=False)
print(f"\n[DONE] {output_path}")
