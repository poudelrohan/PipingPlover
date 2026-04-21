"""
Step 0: Extract PIPL Data from GPS Sheet — 2017
─────────────────────────────────────────────────
Reads DATA SHEET 1 (flock GPS) and DATA SHEET 3 (band resights), extracts
PIPL counts, normalizes columns, and merges via Option A expansion.

2017-specific handling:
  ① Lanark Reef date typo: cell contains '02/08/2017`' (trailing backtick)
     → backtick stripped before date parsing so the row is not lost to year filter
  ② Outback Key group 2 longitude entered as -82751911 (missing decimal)
     → corrected to -82.751911 before coordinate parsing
  ③ DS1 route "Dunlaton Bridge to Ponce Inlet" is a typo for "Dunlawton"
     → corrected so DS1 rows match DS3 "Dunlawton bridge to Ponce Inlet" entries
  ④ DS1 filtered to year == 2017 — drops any stray wrong-year rows
  ⑤ Wild Goose Lagoon DS3 rows have trailing newline in route name
     → handled automatically by normalize_route (whitespace collapsed + stripped)
  ⑥ Highland Beach group 2 and New Smyrna Beach group 6 have positive longitudes
     → auto-corrected in Step 2 (geography validator); no Step 0 fix needed
  ⑦ Unmatched DS3 rows: if the route exists anywhere in DS1 (same date), GPS is
     borrowed from that DS1 entry and the row is kept with an "approx GPS" note.
     If the route has NO DS1 entry at all, the row is dropped — band info without
     any location is not useful enough to keep.
     Pavilion Key DS3 (group 1, no DS1 group 1): DS1 has Pavilion Key group 4 with
     GPS, so those coordinates are borrowed.
  ⑧ DS3 Navarre Beach Soundside route name → corrected to "Navarre Beach Sound Side"
     to match DS1/DS2 spelling (spacing difference causes match failure)
  ⑨ DS3 Bunche Beach group 1 → reassigned to group 2 to match DS1
     (DS1 only has group 2; DS3 recorded group 1 in error)

Option A expansion (same as 2015–2016):
  - One row per banded bird  (TotalObserved=1, with band details)
  - One row for unbanded remainder  (TotalObserved = flock_count − n_banded)
  - If no banded birds match: keep original DS1 row unchanged

Output: db3_2017_extracted.xlsx  (in Output/2017/)
"""

import pandas as pd
import re
import os
import sys

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database3_config import config, get_output_folder, get_filename

# ── Resolve paths ──────────────────────────────────────────────────────────────
script_dir    = os.path.dirname(os.path.abspath(__file__))
output_folder = get_output_folder(script_dir)

year         = config["year"]
input_folder = os.path.normpath(os.path.join(script_dir, config["input_folder"]))
input_path   = os.path.join(input_folder, config["file"])
output_path  = os.path.join(output_folder, get_filename("extracted"))
cols         = config["columns"]

if not os.path.exists(input_path):
    print(f"[ERROR] Input file not found: {input_path}")
    sys.exit(1)


# ══════════════════════════════════════════════════════════════════════════════
# Helper functions
# ══════════════════════════════════════════════════════════════════════════════

def normalize_route(s):
    """
    Normalize a route name for DS1↔DS3 matching only — NOT for display.
    Strips punctuation, lowercases, collapses whitespace.
    Also handles trailing newlines (e.g. Wild Goose Lagoon DS3 entries).
    """
    s = str(s).lower().strip()
    s = re.sub(r'[^\w\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def normalize_time_str(t):
    """Normalize a single time component to HH:MM."""
    import datetime as dt
    if isinstance(t, dt.time):
        return f"{t.hour:02d}:{t.minute:02d}"
    t = str(t).strip()
    t = re.sub(r'\s+[A-Z]{2,3}$', '', t).strip()
    t = re.sub(r'\s*(am|pm|a|p)\s*$', '', t, flags=re.IGNORECASE).strip()
    t = t.replace(';', ':')

    if ':' in t:
        parts = t.split(':')
        try:
            h = int(parts[0])
            m_str = re.sub(r'\D', '', parts[1]) if len(parts) > 1 else '0'
            m = int(m_str) if m_str else 0
            return f"{h:02d}:{m:02d}"
        except (ValueError, IndexError):
            return t
    elif re.match(r'^\d{3,4}$', t):
        return f"{t[:-2].zfill(2)}:{t[-2:]}"
    return t


def normalize_survey_time(raw):
    """
    Normalize a survey time range to 'HH:MM - HH:MM'.
    Handles: '1021-1220', '9:40 am - 12:10 pm', '09:15 to 12:30',
             '07:45 / 17:30', '07:30-14:30 EST', '10:15a- 1:45p'.
    """
    import datetime as dt
    if pd.isna(raw):
        return raw
    if isinstance(raw, dt.time):
        return normalize_time_str(raw)
    raw = str(raw).strip()
    raw = re.sub(r'\s+to\s+', '-', raw, flags=re.IGNORECASE)
    raw = re.sub(r'\s*/\s*', '-', raw)
    raw = re.sub(r'\s*-\s*', '-', raw)
    parts = raw.split('-')
    if len(parts) == 2:
        return f"{normalize_time_str(parts[0])} - {normalize_time_str(parts[1])}"
    return raw


def extract_pipl_count(text, num_individuals=None):
    """Extract the PIPL count from the focal species text."""
    if pd.isna(text):
        return None
    text = str(text)

    m = re.search(r'(\d+)\s*PIPL', text, re.IGNORECASE)
    if m: return int(m.group(1))

    m = re.search(r'PIPL\s*\((\d+)\)', text, re.IGNORECASE)
    if m: return int(m.group(1))

    m = re.search(r'PIPL\s*\((\d+)[\s,]', text, re.IGNORECASE)
    if m: return int(m.group(1))

    m = re.search(r'PIPL\s*[-:]\s*(\d+)', text, re.IGNORECASE)
    if m: return int(m.group(1))

    m = re.search(r'PIPL\s+(\d+)', text, re.IGNORECASE)
    if m: return int(m.group(1))

    m = re.search(r'[Pp]iping\s*[Pp]lovers?\s*\((\d+)', text)
    if m: return int(m.group(1))

    if re.search(r'\bPIPL\b', text, re.IGNORECASE) or re.search(r'[Pp]iping\s*[Pp]lover', text):
        if pd.notna(num_individuals):
            try: return int(float(num_individuals))
            except (ValueError, TypeError): pass
        return 1

    return None


def parse_coordinate(val, is_longitude=False):
    """Parse coordinate into float. Handles N/S/E/W prefixes, DMS, etc."""
    if pd.isna(val):
        return None
    val = str(val).strip().replace('\xa0', ' ').strip()
    if val == "" or val.lower() == "same":
        return None

    try:
        return float(val)
    except ValueError:
        pass

    m = re.match(r'^([NSEW])\s*([\d.]+)$', val, re.IGNORECASE)
    if m:
        direction = m.group(1).upper()
        num = float(m.group(2))
        if direction in ('S', 'W'):
            num = -num
        return num

    m = re.match(r'^(-?[\d.]+)\s*([NSEW])$', val, re.IGNORECASE)
    if m:
        num = float(m.group(1))
        direction = m.group(2).upper()
        if direction in ('S', 'W'):
            num = -abs(num)
        else:
            num = abs(num)
        return num

    m = re.match(r"^(\d+)\s+([\d.]+)['\u2032]?\s*([NSEW])\s*$", val, re.IGNORECASE)
    if m:
        deg = int(m.group(1))
        minutes = float(m.group(2))
        direction = m.group(3).upper()
        result = deg + minutes / 60.0
        if direction in ('S', 'W'):
            result = -result
        return round(result, 6)

    print(f"    [WARNING] Could not parse coordinate: '{val}'")
    return None


def is_blank_band_val(val):
    """Return True if a band leg value should be treated as missing/blank."""
    if not val:
        return True
    if re.match(r'^-+$', val.strip()):
        return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
# Read DS1 (flock GPS)
# ══════════════════════════════════════════════════════════════════════════════

df = pd.read_excel(
    input_path,
    sheet_name=config["sheet"],
    header=config["header_row"],
)
print(f"  Loaded {len(df)} rows from '{config['file']}' → sheet '{config['sheet']}'")

# ── ① Fix Lanark Reef date backtick: '02/08/2017`' → '02/08/2017' ──────────────
# The backtick causes pd.to_datetime to return NaT, which drops the row from the
# year filter below. Strip it before any date processing.
date_col_raw = cols["date"]
if date_col_raw in df.columns:
    str_mask = df[date_col_raw].apply(lambda v: isinstance(v, str))
    if str_mask.any():
        cleaned = df.loc[str_mask, date_col_raw].str.replace('`', '', regex=False)
        n_fixed = (cleaned != df.loc[str_mask, date_col_raw]).sum()
        df.loc[str_mask, date_col_raw] = pd.to_datetime(cleaned, errors='coerce')
        if n_fixed > 0:
            print(f"  [FIX ①] Stripped backtick from {n_fixed} date cell(s) (Lanark Reef)")

# ── ④ Filter DS1 to survey year only ──────────────────────────────────────────
if date_col_raw in df.columns:
    _orig_len = len(df)
    df = df[pd.to_datetime(df[date_col_raw], errors='coerce').dt.year == int(year)].copy().reset_index(drop=True)
    _dropped = _orig_len - len(df)
    if _dropped > 0:
        print(f"  [FIX ④] Dropped {_dropped} DS1 row(s) with year ≠ {year} (stray wrong-year entries)")

# ── ② Fix Outback Key group 2 longitude: -82751911 → -82.751911 ───────────────
# Longitude was entered without a decimal point. parse_coordinate would return
# -82751911.0, which Step 2 would remove as outside Florida bounds.
loc_col_raw = cols["location"]
grp_col_raw = config.get("group_column")
lon_col_raw = cols["longitude"]

if all(c in df.columns for c in [loc_col_raw, grp_col_raw, lon_col_raw]):
    bad_lon_mask = (
        df[loc_col_raw].astype(str).str.contains("Outback Key", case=False, na=False)
        & (df[grp_col_raw].astype(str) == "2")
        & (pd.to_numeric(df[lon_col_raw], errors='coerce') == -82751911)
    )
    if bad_lon_mask.any():
        df.loc[bad_lon_mask, lon_col_raw] = -82.751911
        print(f"  [FIX ②] Corrected Outback Key group 2 longitude: -82751911 → -82.751911 ({bad_lon_mask.sum()} row(s))")

# ── ③ Fix DS1 route name typo: Dunlaton → Dunlawton Bridge to Ponce Inlet ──────
# DS3 uses "Dunlawton bridge to Ponce Inlet"; DS1 has a typo "Dunlaton Bridge to
# Ponce Inlet". normalize_route cannot bridge a spelling difference, so correct
# DS1 before matching. We use the corrected name in output (the original was wrong).
if loc_col_raw in df.columns:
    dunlaton_mask = df[loc_col_raw].astype(str).str.contains("Dunlaton", case=False, na=False)
    if dunlaton_mask.any():
        df.loc[dunlaton_mask, loc_col_raw] = "Dunlawton Bridge to Ponce Inlet"
        print(f"  [FIX ③] Corrected DS1 route name: 'Dunlaton' → 'Dunlawton Bridge to Ponce Inlet' ({dunlaton_mask.sum()} row(s))")

# ── Parse coordinates ──────────────────────────────────────────────────────────
lat_src = cols["latitude"]
lon_src = cols["longitude"]

if lat_src in df.columns:
    df[lat_src] = df[lat_src].apply(lambda v: parse_coordinate(v, is_longitude=False))
if lon_src in df.columns:
    df[lon_src] = df[lon_src].apply(lambda v: parse_coordinate(v, is_longitude=True))

lat_parsed = df[lat_src].notna().sum() if lat_src in df.columns else 0
lon_parsed = df[lon_src].notna().sum() if lon_src in df.columns else 0
print(f"  Coordinates parsed: {lat_parsed} lat, {lon_parsed} lon")

# ── Extract PIPL count ─────────────────────────────────────────────────────────
focal_col = cols["focal_species"]
has_num   = config.get("has_num_individuals", False)
num_col   = cols.get("num_individuals")

if has_num and num_col and num_col in df.columns:
    df["TotalObserved"] = df.apply(
        lambda r: extract_pipl_count(r[focal_col], r[num_col]), axis=1
    )
else:
    df["TotalObserved"] = df[focal_col].apply(extract_pipl_count)

extracted  = df["TotalObserved"].notna().sum()
failed     = df["TotalObserved"].isna().sum()
total_pipl = int(df["TotalObserved"].sum()) if extracted > 0 else 0

print(f"\n  PIPL count extracted: {extracted} rows, {failed} failed")
print(f"  Total PIPL count: {total_pipl}")

if failed > 0:
    print(f"\n  [WARNING] Could not extract PIPL count from {failed} rows:")
    for _, r in df[df["TotalObserved"].isna()].iterrows():
        print(f"    '{r[focal_col]}'")

# ── Normalize DS1 column names ─────────────────────────────────────────────────
rename_map = {}
group_col  = config.get("group_column")

for std_name, src_name in [
    ("SurveyDate",       cols["date"]),
    ("Route",            cols["location"]),
    ("Latitude",         cols["latitude"]),
    ("Longitude",        cols["longitude"]),
    ("Observer",         cols["observers"]),
    ("ObserverEmail",    cols.get("email")),
    ("FlagCode",         cols.get("flag_code")),
    ("BandCombo",        cols.get("band_combo")),
    ("Comments",         cols.get("comments")),
    ("GroupNumber",      group_col),
    ("SurveyTime",       cols.get("time")),
    ("WeatherCondition", cols.get("weather")),
]:
    if src_name and src_name in df.columns:
        rename_map[src_name] = std_name

df = df.rename(columns=rename_map)

# ── Normalize SurveyTime ───────────────────────────────────────────────────────
if "SurveyTime" in df.columns:
    df["SurveyTime"] = df["SurveyTime"].apply(normalize_survey_time)
    print(f"  SurveyTime normalized: {df['SurveyTime'].notna().sum()} values")


# ══════════════════════════════════════════════════════════════════════════════
# Read DS3 (band resights)
# ══════════════════════════════════════════════════════════════════════════════

band_cfg = config.get("band_sheet")

print(f"\n  Loading band resight sheet: '{band_cfg['sheet']}'")
bands = pd.read_excel(
    input_path,
    sheet_name=band_cfg["sheet"],
    header=band_cfg["header_row"],
)
bcols = band_cfg["columns"]

# Filter to PIPL only
species_col = bcols["species"]
bands = bands[bands[species_col].astype(str).str.contains("PIPL", case=False, na=False)].copy()
print(f"  PIPL band resight rows: {len(bands)}")

# ── ⑧ Fix DS3 Navarre Beach route name: "Soundside" → "Sound Side" ────────────
# DS1 and DS2 both use "Navarre Beach Sound Side" (two words); DS3 uses
# "Navarre Beach Soundside" (one word). normalize_route can't bridge this so
# correct DS3 to match DS1 before Option A matching.
navarre_mask = bands[bcols["location"]].astype(str).str.contains("Navarre Beach Soundside", case=False, na=False)
if navarre_mask.any():
    bands.loc[navarre_mask, bcols["location"]] = bands.loc[navarre_mask, bcols["location"]].str.replace(
        "Soundside", "Sound Side", case=False, regex=False
    )
    print(f"  [FIX ⑧] Corrected DS3 Navarre Beach route name: 'Soundside' → 'Sound Side' ({navarre_mask.sum()} row(s))")

# ── ⑨ Fix DS3 Bunche Beach group 1 → 2 ────────────────────────────────────────
# DS1 only has Bunche Beach group 2 (15 PIPL). DS3 recorded group 1 in error.
# Reassign before matching so the banded birds link to the correct DS1 flock.
bunche_mask = (
    bands[bcols["location"]].astype(str).str.contains("Bunche Beach", case=False, na=False)
    & (bands[bcols["group"]].astype(str) == "1")
)
if bunche_mask.any():
    bands.loc[bunche_mask, bcols["group"]] = 2
    print(f"  [FIX ⑨] Reassigned DS3 Bunche Beach group 1 → 2 ({bunche_mask.sum()} row(s))")

# ── Build BandCombo ────────────────────────────────────────────────────────────
def build_band_combo(row):
    parts = []
    orientation = str(row.get(bcols["code_orientation"], "")).strip() \
        if pd.notna(row.get(bcols["code_orientation"])) else ""
    if orientation and not is_blank_band_val(orientation):
        parts.append(f"Orientation:{orientation}")
    for label, col_name in [
        ("UL", bcols["upper_left"]),
        ("LL", bcols["lower_left"]),
        ("UR", bcols["upper_right"]),
        ("LR", bcols["lower_right"]),
    ]:
        val = str(row.get(col_name, "")).strip() if pd.notna(row.get(col_name)) else ""
        val = "-" if is_blank_band_val(val) else val
        parts.append(f"{label}:{val}")
    return "; ".join(parts)

bands["BandCombo"] = bands.apply(build_band_combo, axis=1)

# ── Rename DS3 columns to internal names ──────────────────────────────────────
band_rename = {}
for std, src in [
    ("_band_date",     bcols["date"]),
    ("_band_route",    bcols["location"]),
    ("_band_group",    bcols["group"]),
    ("_band_observer", bcols["observer"]),
    ("FlagCode",       bcols["flag_code"]),
    ("FlagColor",      bcols.get("flag_color")),
    ("_band_comments", bcols.get("comments")),
]:
    if src and src in bands.columns:
        band_rename[src] = std
bands = bands.rename(columns=band_rename)


# ══════════════════════════════════════════════════════════════════════════════
# Option A expansion: merge DS1 flocks with DS3 banded birds
# ══════════════════════════════════════════════════════════════════════════════

# Route names are normalized for matching (punctuation + whitespace insensitive).
# FIX ③ already corrected DS1 "Dunlaton" → "Dunlawton", so both sides now match.
# FIX ⑤ trailing newlines in DS3 Wild Goose Lagoon are collapsed by normalize_route.

matched_band_indices = set()
expanded_rows        = []

for _, flock_row in df.iterrows():
    flock_date  = flock_row.get("SurveyDate")
    flock_route = normalize_route(flock_row.get("Route", ""))
    flock_group = flock_row.get("GroupNumber")
    flock_pipl  = flock_row.get("TotalObserved")

    matched_bands = bands[
        (bands["_band_date"].astype(str).str[:10] == str(flock_date)[:10]) &
        (bands["_band_route"].astype(str).apply(normalize_route) == flock_route) &
        (bands["_band_group"].astype(str) == str(flock_group))
    ]

    n_banded = len(matched_bands)

    if n_banded == 0:
        expanded_rows.append(flock_row.to_dict())
    else:
        for idx, band_row in matched_bands.iterrows():
            matched_band_indices.add(idx)
            row_data = flock_row.to_dict()
            row_data["TotalObserved"] = 1
            row_data["FlagCode"]      = band_row.get("FlagCode", "")
            row_data["FlagColor"]     = band_row.get("FlagColor", "")
            row_data["BandCombo"]     = band_row.get("BandCombo", "")
            band_comment = band_row.get("_band_comments", "")
            if pd.notna(band_comment) and str(band_comment).strip():
                existing = str(row_data.get("Comments", "")) if pd.notna(row_data.get("Comments")) else ""
                row_data["Comments"] = f"{existing} | Band: {band_comment}" if existing else f"Band: {band_comment}"
            expanded_rows.append(row_data)

        unbanded_count = int(flock_pipl - n_banded) if pd.notna(flock_pipl) else 0
        if unbanded_count > 0:
            row_data = flock_row.to_dict()
            row_data["TotalObserved"] = unbanded_count
            row_data["FlagCode"]      = ""
            row_data["FlagColor"]     = ""
            row_data["BandCombo"]     = ""
            expanded_rows.append(row_data)
        elif unbanded_count < 0:
            print(f"    [WARNING] More banded ({n_banded}) than PIPL count ({flock_pipl}) "
                  f"at {flock_row.get('Route')} grp {flock_group} on {str(flock_date)[:10]}")

# ── ⑦ Handle unmatched DS3 rows ───────────────────────────────────────────────
# Rule: a DS3 row with no group-level DS1 match is only useful if we can give it
# a location. Strategy:
#   • If the route + date appears anywhere in DS1 → borrow GPS from that DS1 entry
#     and keep the row (with a comment noting the GPS is approximate).
#   • If the route has NO DS1 entry at all → drop the row entirely. Band info
#     without any location is not meaningful for analysis.

unmatched_bands   = bands[~bands.index.isin(matched_band_indices)]
standalone_kept   = 0
standalone_dropped = 0

for _, band_row in unmatched_bands.iterrows():
    band_route  = band_row.get("_band_route")
    band_date   = band_row.get("_band_date")
    band_comment = band_row.get("_band_comments", "")
    route_norm  = normalize_route(str(band_route))
    date_str    = str(band_date)[:10]

    # Look for any DS1 row at same route + date (group number may differ).
    # Use substring matching in both directions: DS3 "Pavilion Key" should match
    # DS1 "Everglades City to Pavilion Key", and vice versa.
    ds1_routes_norm = df["Route"].apply(normalize_route)
    route_match = (
        (ds1_routes_norm == route_norm) |
        ds1_routes_norm.str.contains(route_norm, regex=False) |
        pd.Series([route_norm in r for r in ds1_routes_norm], index=df.index)
    )
    ds1_match = df[
        route_match &
        (df["SurveyDate"].astype(str).str[:10] == date_str) &
        df["Latitude"].notna() & df["Longitude"].notna()
    ]

    if ds1_match.empty:
        # No DS1 entry for this route — drop, nothing to anchor it geographically
        standalone_dropped += 1
        print(f"    [DROP] Unmatched DS3 row dropped — route '{band_route}' has no DS1 entry "
              f"(no GPS available): grp {band_row.get('_band_group')}, {date_str}")
        continue

    # Borrow GPS from first DS1 row found at this route
    ref = ds1_match.iloc[0]
    lat = ref["Latitude"]
    lon = ref["Longitude"]

    comment_parts = []
    if pd.notna(band_comment) and str(band_comment).strip():
        comment_parts.append(str(band_comment).strip())
    comment_parts.append(
        f"GPS approx: no group match in DS1 — coordinates from DS1 group {ref['GroupNumber']} "
        f"at same route/date"
    )

    row_data = {
        "SurveyDate":    band_date,
        "Route":         band_route,
        "GroupNumber":   band_row.get("_band_group"),
        "Observer":      band_row.get("_band_observer", ""),
        "TotalObserved": 1,
        "FlagCode":      band_row.get("FlagCode", ""),
        "FlagColor":     band_row.get("FlagColor", ""),
        "BandCombo":     band_row.get("BandCombo", ""),
        "Comments":      " | ".join(comment_parts),
        "Latitude":      lat,
        "Longitude":     lon,
    }
    standalone_kept += 1
    expanded_rows.append(row_data)

if standalone_kept:
    print(f"  [INFO] {standalone_kept} unmatched DS3 row(s) kept with GPS borrowed from DS1 same-route entry")
if standalone_dropped:
    print(f"  [INFO] {standalone_dropped} unmatched DS3 row(s) dropped — no DS1 entry found for route")

df = pd.DataFrame(expanded_rows)
print(f"\n  After Option A expansion: {len(df)} total rows")

# ── Keep only the columns we need ──────────────────────────────────────────────
keep = [
    "SurveyDate", "SurveyTime", "WeatherCondition",
    "Route", "TotalObserved", "GroupNumber", "Latitude", "Longitude",
    "Observer", "ObserverEmail", "FlagCode", "FlagColor", "BandCombo", "Comments",
    "_removal_reason",
]
keep = [c for c in keep if c in df.columns]
df = df[keep]

# ── Save output ────────────────────────────────────────────────────────────────
df.to_excel(output_path, index=False)

print(f"\n[DONE] Step 0 complete")
print(f"  Year    : {year}")
print(f"  Rows    : {len(df)}")
print(f"  Columns : {list(df.columns)}")
print(f"  Output  : {output_path}")
