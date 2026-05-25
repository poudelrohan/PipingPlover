"""
Step 0: Extract PIPL Data — 2018
──────────────────────────────────
Reads the Band Review file (wide format, one row per route) and melts it
into one row per bird per point, applying Option A band expansion.

Band data source priority per (route, point):
  1. Biologist correction  — from Biologist_Band_Review_Completed_2018-2019.xlsx col H
  2. Our interpretation    — from the same file col F (already structured 1)/2)/3))
  3. Raw Band Review text  — fallback if the review file has no entry for this point

Wide → Long logic (per route row):
  For each point N in [1,2,3,...,15,17,19]:
    • Skip if {N}PIPL is null or 0
    • Skip if {N}Lat / {N}Long are both null AND no GPS can be borrowed
      from another year (row logged to Removed sheet)
    • Parse band entries (strip N) prefix, split on newlines)
    • Option A expansion:
        - k banded entries → k rows with TotalObserved=1, BandCombo=band string
        - remainder = {N}PIPL − k → 1 row with TotalObserved=remainder, no band info
        - if k == 0: 1 row with TotalObserved={N}PIPL, no band info

Output: db3_2018_extracted.xlsx  (in Output/2018/)
"""

import pandas as pd
import re
import os
import sys
from pathlib import Path

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database3_config import config, get_output_folder, get_filename

script_dir    = os.path.dirname(os.path.abspath(__file__))
output_folder = get_output_folder(script_dir)
year          = config["year"]

input_dir   = os.path.normpath(os.path.join(script_dir, config["input_folder"]))
input_path  = os.path.join(input_dir, config["file"])
output_path = os.path.join(output_folder, get_filename("extracted"))
removed_path = os.path.join(output_folder, get_filename("removed_gps"))

if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    sys.exit(1)

cols   = config["columns"]
points = config["points"]


# ══════════════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════════════

def normalize_route(s):
    """Lower, strip punctuation, collapse whitespace — for route matching."""
    s = str(s).lower().strip()
    s = re.sub(r'[^\w\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def parse_coord(val):
    """Parse a coordinate value to float. Returns None if unparseable."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        f = float(str(val).strip())
        return f if f != 0.0 else None
    except (ValueError, TypeError):
        return None


def combine_weather(temp, wind, rain):
    """Combine temperature, wind, rain into a single WeatherCondition string."""
    parts = []
    for v in (temp, wind, rain):
        s = str(v).strip() if v is not None and not (isinstance(v, float) and pd.isna(v)) else ""
        if s and s.lower() not in ("nan", "none", ""):
            parts.append(s)
    return ", ".join(parts) if parts else None


def parse_band_entries(text):
    """
    Parse the structured "1) ... \\n2) ..." text from the Band Review file
    into a list of individual bird band strings.
    Returns [] if the cell is null, empty, or "No banded birds".
    """
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return []
    text = str(text).strip()
    if not text or text.lower() == "no banded birds":
        return []

    # Normalise inline numbering: "1) bird1 2) bird2" → "1) bird1\n2) bird2"
    text = re.sub(r'(?<!\n)(\s+)(\d+\)\s*)', r'\n\2', text)

    entries = []
    for line in text.split("\n"):
        line = line.strip()
        m = re.match(r'^\d+\)\s*', line)
        if m:
            entry = line[m.end():].strip()
            if entry:
                entries.append(entry)

    # If the text exists but has no numbered entries it wasn't parsed —
    # treat the whole cell as 1 bird's info (fallback).
    if not entries and text:
        entries = [text]

    return entries


# ══════════════════════════════════════════════════════════════════════════════
# Biologist Corrections — override raw band text with reviewed data
# ══════════════════════════════════════════════════════════════════════════════

def build_biologist_corrections():
    """
    Read the completed biologist review file and build:
      (normalized_route, point_N) → corrected_band_text

    Priority: col H (correction) if filled, else col F (our interpretation).
    If the correction text is a meta-instruction (e.g. "treat as unbanded")
    rather than actual band data, map to empty string → pipeline treats as
    no banded birds at that point.
    """
    review_path = (
        Path(script_dir).parents[2]
        / "Databases" / "Database3BiologistReview"
        / "Biologist_Band_Review_Completed_2018-2019.xlsx"
    )
    if not review_path.exists():
        print(f"  [INFO] No biologist review file found — using raw Band Review data")
        return {}

    import openpyxl
    wb = openpyxl.load_workbook(review_path)
    ws = wb["2018"]
    corrections = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        route      = str(row[0] or "").strip()
        point      = row[1]
        interp     = str(row[5] or "").strip()   # col F — our interpretation
        correction = str(row[7] or "").strip()   # col H — biologist correction

        if not route or not point:
            continue

        # Choose the best available text: correction > interpretation
        text = correction if correction else interp

        if not text:
            continue

        # Detect meta-instructions like "treat as unbanded bird" — not actual band data
        t_lower = text.lower()
        if "unbanded" in t_lower and "1)" not in t_lower:
            text = ""   # empty → all birds at this point treated as unbanded

        key = (normalize_route(route), int(point))
        corrections[key] = text

    n_corrections = sum(1 for v in corrections.values() if v)
    n_unbanded    = sum(1 for v in corrections.values() if not v)
    print(f"  Biologist corrections loaded: {n_corrections} band entries, "
          f"{n_unbanded} points overridden to unbanded")
    return corrections


# ══════════════════════════════════════════════════════════════════════════════
# GPS Lookup — borrow from other years when a point has no coordinates
# ══════════════════════════════════════════════════════════════════════════════

def build_gps_lookup():
    """
    Scan 2019–2024 clean sheets and build:
      (normalized_route, point_N)  →  (lat, lon, source_year)

    Handles two column formats:
      • 2019 "Form Responses 1"  : 'Point N Latitude ...'
      • 2020–2024 "Focal Obs"    : 'Latitude (point N)'
    """
    lookup   = {}
    root     = Path(script_dir).parents[2] / "Databases" / "Database3Clean"

    def try_add(route_val, n, lat, lon, yr):
        if route_val and lat and lon:
            key = (normalize_route(str(route_val)), int(n))
            if key not in lookup:          # first year found wins
                lookup[key] = (lat, lon, yr)

    for yr in range(2019, 2025):
        fpath = root / f"Winter Birds {yr} Clean.xlsx"
        if not fpath.exists():
            continue
        try:
            wb_tmp = pd.ExcelFile(fpath)
        except Exception:
            continue

        # ── 2019: "Form Responses 1" ─────────────────────────────────────────
        if "Form Responses 1" in wb_tmp.sheet_names:
            df_tmp = wb_tmp.parse("Form Responses 1", header=0)
            # Route col: the one whose name contains 'route' (case-insensitive)
            route_col_tmp = next(
                (c for c in df_tmp.columns if 'route' in str(c).lower()), None
            )
            # If no route col, use first col
            if route_col_tmp is None and len(df_tmp.columns) > 0:
                route_col_tmp = df_tmp.columns[0]

            for n in range(1, 20):
                lat_col_tmp = next(
                    (c for c in df_tmp.columns
                     if re.search(rf'\bPoint\s+{n}\s+Lat', str(c), re.IGNORECASE)), None
                )
                lon_col_tmp = next(
                    (c for c in df_tmp.columns
                     if re.search(rf'\bPoint\s+{n}\s+Long', str(c), re.IGNORECASE)), None
                )
                if lat_col_tmp is None or lon_col_tmp is None:
                    continue
                for _, r in df_tmp.iterrows():
                    try_add(r.get(route_col_tmp), n,
                            parse_coord(r.get(lat_col_tmp)),
                            parse_coord(r.get(lon_col_tmp)), yr)

        # ── 2020–2024: "Focal Observations" ──────────────────────────────────
        if "Focal Observations" in wb_tmp.sheet_names:
            df_tmp = wb_tmp.parse("Focal Observations", header=0)
            route_col_tmp = df_tmp.columns[0]   # first col is route/transect

            for n in range(1, 20):
                lat_col_tmp = next(
                    (c for c in df_tmp.columns
                     if re.search(rf'Lat.*\bpoint\s+{n}\b', str(c), re.IGNORECASE)
                     or re.search(rf'\bpoint\s+{n}\b.*Lat', str(c), re.IGNORECASE)), None
                )
                lon_col_tmp = next(
                    (c for c in df_tmp.columns
                     if re.search(rf'Long.*\bpoint\s+{n}\b', str(c), re.IGNORECASE)
                     or re.search(rf'\bpoint\s+{n}\b.*Long', str(c), re.IGNORECASE)), None
                )
                if lat_col_tmp is None or lon_col_tmp is None:
                    continue
                for _, r in df_tmp.iterrows():
                    try_add(r.get(route_col_tmp), n,
                            parse_coord(r.get(lat_col_tmp)),
                            parse_coord(r.get(lon_col_tmp)), yr)

    print(f"  GPS lookup built: {len(lookup)} (route, point) pairs from 2019–2024")
    return lookup


# ══════════════════════════════════════════════════════════════════════════════
# Main extraction
# ══════════════════════════════════════════════════════════════════════════════

print(f"\n{'='*60}")
print(f"Step 0 — 2018 PIPL extraction")
print(f"  Input : {input_path}")
print(f"  Output: {output_path}")
print()

# ── Read Band Review ───────────────────────────────────────────────────────────
df_raw = pd.read_excel(input_path, sheet_name=config["sheet"], header=config["header_row"])
print(f"  Loaded {len(df_raw)} route rows from Band Review")

# ── Load biologist corrections ─────────────────────────────────────────────────
bio_corrections = build_biologist_corrections()

# ── Build GPS lookup from other years ─────────────────────────────────────────
gps_lookup = build_gps_lookup()

# ── Identify {N}PIPL / {N}Lat / {N}Long / {N}PIPLbands columns ───────────────
# Build a mapping: point_N → {pipl_col, lat_col, long_col, bands_col}
point_cols = {}
for col in df_raw.columns:
    col_str = str(col)
    m_lat   = re.fullmatch(r'(\d+)Lat', col_str, re.IGNORECASE)
    m_long  = re.fullmatch(r'(\d+)Long', col_str, re.IGNORECASE)
    m_pipl  = re.fullmatch(r'(\d+)PIPL', col_str, re.IGNORECASE)
    m_bands = re.fullmatch(r'(\d+)PIPL.?ands', col_str, re.IGNORECASE)

    for m, field in [(m_lat,'lat'), (m_long,'long'), (m_pipl,'pipl'), (m_bands,'bands')]:
        if m:
            n = int(m.group(1))
            if n not in point_cols:
                point_cols[n] = {}
            point_cols[n][field] = col

# ── Melt wide → long with Option A expansion ──────────────────────────────────
out_rows      = []
removed_rows  = []
stats = {
    "routes":          len(df_raw),
    "points_skipped_no_pipl": 0,
    "points_dropped_no_gps":  0,
    "points_gps_borrowed":    0,
    "rows_banded":     0,
    "rows_unbanded":   0,
    "total_pipl":      0,
}

for _, route_row in df_raw.iterrows():

    # ── Route-level fields ───────────────────────────────────────────────────
    route    = str(route_row.get(cols["route"], "")).strip()
    date_val = route_row.get(cols["date"])
    time_val = route_row.get(cols["time"])
    weather  = combine_weather(
        route_row.get(cols["temp"]),
        route_row.get(cols["wind"]),
        route_row.get(cols["rain"]),
    )
    observer = route_row.get(cols["observers"])
    email    = route_row.get(cols["email"])
    comments = route_row.get(cols["comments"])

    # Normalize date
    if pd.notna(date_val):
        try:
            date_val = pd.to_datetime(date_val)
        except Exception:
            pass

    # Normalize comments (null → None)
    if pd.isna(comments) if isinstance(comments, float) else False:
        comments = None
    comments = str(comments).strip() if comments and str(comments).strip().lower() not in ("nan","none","") else None

    # ── Loop over each point ─────────────────────────────────────────────────
    for n in points:
        pc = point_cols.get(n, {})

        # PIPL count at this point
        pipl_raw = route_row.get(pc.get("pipl")) if "pipl" in pc else None
        try:
            pipl_count = int(float(pipl_raw)) if pipl_raw is not None and not (isinstance(pipl_raw, float) and pd.isna(pipl_raw)) else 0
        except (ValueError, TypeError):
            pipl_count = 0

        if pipl_count == 0:
            stats["points_skipped_no_pipl"] += 1
            continue

        # GPS at this point
        lat = parse_coord(route_row.get(pc.get("lat"))) if "lat" in pc else None
        lon = parse_coord(route_row.get(pc.get("long"))) if "long" in pc else None

        if lat is None or lon is None:
            # Try borrowing from another year (logged to console only, not added to Comments)
            key = (normalize_route(route), n)
            if key in gps_lookup:
                lat, lon, src_yr = gps_lookup[key]
                stats["points_gps_borrowed"] += 1
                print(f"  [GPS BORROW] {route} pt {n}: lat={lat}, lon={lon} (from {src_yr})")
            else:
                # No GPS available — drop and log
                stats["points_dropped_no_gps"] += 1
                removed_rows.append({
                    "Route":          route,
                    "SurveyDate":     date_val,
                    "GroupNumber":    n,
                    "TotalObserved":  pipl_count,
                    "removal_reason": f"No GPS for point {n} and no matching coordinates found in any other year",
                })
                print(f"  [DROP] {route} pt {n}: PIPL={pipl_count}, no GPS in any year")
                continue

        # Band entries — use biologist correction if available, else raw Band Review text
        bands_raw = route_row.get(pc.get("bands")) if "bands" in pc else None
        bio_key   = (normalize_route(route), n)
        if bio_key in bio_corrections:
            bands_raw = bio_corrections[bio_key]   # "" = unbanded, string = corrected combo
        band_list  = parse_band_entries(bands_raw)
        n_banded   = len(band_list)
        remainder  = pipl_count - n_banded

        if remainder < 0:
            print(f"  [WARNING] {route} pt {n}: {n_banded} banded entries > PIPL count ({pipl_count}). "
                  f"Keeping all band entries, TotalObserved=1 each.")
            remainder = 0

        # Base fields shared by all rows from this point
        base = {
            "SurveyDate":       date_val,
            "SurveyTime":       str(time_val).strip() if pd.notna(time_val) else None,
            "WeatherCondition": weather,
            "Route":            route,
            "Latitude":         lat,
            "Longitude":        lon,
            "GroupNumber":      n,
            "Observer":         str(observer).strip() if pd.notna(observer) else None,
            "ObserverEmail":    str(email).strip() if pd.notna(email) else None,
            "FlagCode":         None,
            "FlagColor":        None,
            "Comments":         comments,   # original observer text only — no pipeline notes
        }

        # ── Banded rows ──────────────────────────────────────────────────────
        for band_text in band_list:
            row = base.copy()
            row["TotalObserved"] = 1
            row["BandCombo"]     = band_text
            out_rows.append(row)
            stats["rows_banded"] += 1

        # ── Unbanded remainder row ───────────────────────────────────────────
        if remainder > 0:
            row = base.copy()
            row["TotalObserved"] = remainder
            row["BandCombo"]     = None
            out_rows.append(row)
            stats["rows_unbanded"] += 1
        elif n_banded == 0:
            # No band info at all — one row for all birds at this point
            row = base.copy()
            row["TotalObserved"] = pipl_count
            row["BandCombo"]     = None
            out_rows.append(row)
            stats["rows_unbanded"] += 1

        stats["total_pipl"] += pipl_count

# ── Build DataFrames ──────────────────────────────────────────────────────────
df_out = pd.DataFrame(out_rows)
df_removed = pd.DataFrame(removed_rows) if removed_rows else pd.DataFrame(
    columns=["Route", "SurveyDate", "GroupNumber", "TotalObserved", "removal_reason"]
)

# ── Save ──────────────────────────────────────────────────────────────────────
df_out.to_excel(output_path, index=False)
df_removed.to_excel(removed_path, index=False)

# ── Summary ───────────────────────────────────────────────────────────────────
print(f"\n{'─'*50}")
print(f"[DONE] Step 0 — 2018")
print(f"  Routes processed:            {stats['routes']}")
print(f"  Points skipped (0 PIPL):     {stats['points_skipped_no_pipl']}")
print(f"  Points dropped (no GPS):     {stats['points_dropped_no_gps']}")
print(f"  Points GPS borrowed:         {stats['points_gps_borrowed']}")
print(f"  Output rows — banded:        {stats['rows_banded']}")
print(f"  Output rows — unbanded:      {stats['rows_unbanded']}")
print(f"  Total output rows:           {len(df_out)}")
print(f"  Total PIPL accounted for:    {stats['total_pipl']}")
print(f"  Extracted file:  {output_path}")
print(f"  Removed file:    {removed_path}")
