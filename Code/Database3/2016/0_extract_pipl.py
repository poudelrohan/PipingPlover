"""
Step 0: Extract PIPL Data from GPS Sheet — 2016
─────────────────────────────────────────────────
Reads DATA SHEET 1 (flock GPS) and DATA SHEET 3 (band resights), extracts
PIPL counts, normalizes columns, and merges via Option A expansion.

2016-specific handling (beyond standard 2015 logic):
  ① "same" in lat/lon → carry forward coordinates from the row directly above
  ② DS3 date typo: 2015-02-08 → corrected to 2016-02-08
  ③ DS3 Bunche Beach Preserve: Group or Point # NaN → assigned 2
  ④ Route names normalized for matching (punctuation-insensitive) but original
     name is kept in output (handles "Crandon Park Beach; ..." vs ", ~...")
  ⑤ Unmatched DS3 rows (banded birds with no matchable DS1 flock) → included
     as standalone rows with TotalObserved=1 and NaN GPS rather than dropped,
     EXCEPT routes in DS3_REMOVAL_ROUTES which are flagged for removal instead
  ⑥ "---" (any run of dashes) in leg positions treated as missing → "-"
  ⑦ "unreadable" flag codes kept as-is; leg position text kept if not parsable
  ⑧ DS1 filtered to year == 2016 — drops stray 2017 rows entered in wrong file
     (Long Key 2017-02-08 group 2, Mizell-Johnson 2017-02-07 group 1)
  ⑨ Three Rooker Bar north island unmatched DS3 rows (groups 2 & 3) removed —
     DS2 confirms north island total = 10 = DS1 groups 4+5; these banded birds
     are already counted in DS1; group numbers on DS3 are a recording error

Option A expansion (same as 2015):
  - One row per banded bird  (TotalObserved=1, with band details)
  - One row for unbanded remainder  (TotalObserved = flock_count − n_banded)
  - If no banded birds match: keep original DS1 row unchanged

Output: db3_2016_extracted.xlsx  (in Output/2016/)
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
    Handles human-entered variations like semicolons vs commas, tildes, etc.
    Example: 'Crandon Park Beach; lifeguard stand 1<->13'
          == 'Crandon Park Beach, ~lifeguard stand 1<->13'
    """
    s = str(s).lower().strip()
    s = re.sub(r'[^\w\s]', ' ', s)   # replace all punctuation with space
    s = re.sub(r'\s+', ' ', s).strip()  # collapse whitespace
    return s


def normalize_time_str(t):
    """Normalize a single time component to HH:MM."""
    import datetime as dt
    # Handle pandas-parsed datetime.time objects (from Excel time cells)
    if isinstance(t, dt.time):
        return f"{t.hour:02d}:{t.minute:02d}"
    t = str(t).strip()
    t = re.sub(r'\s+[A-Z]{2,3}$', '', t).strip()                      # strip timezone (EST, etc.)
    t = re.sub(r'\s*(am|pm|a|p)\s*$', '', t, flags=re.IGNORECASE).strip()  # strip am/pm/a/p
    t = t.replace(';', ':')                                             # fix semicolon typo

    if ':' in t:
        parts = t.split(':')
        try:
            h = int(parts[0])
            # Strip any trailing non-digit chars from minutes (e.g. "15a" → 15)
            m_str = re.sub(r'\D', '', parts[1]) if len(parts) > 1 else '0'
            m = int(m_str) if m_str else 0
            return f"{h:02d}:{m:02d}"
        except (ValueError, IndexError):
            return t  # fallback
    elif re.match(r'^\d{3,4}$', t):
        return f"{t[:-2].zfill(2)}:{t[-2:]}"
    return t  # fallback


def normalize_survey_time(raw):
    """
    Normalize a survey time range to 'HH:MM - HH:MM'.
    Handles: '1021-1220', '9:40 am - 12:10 pm', '09:15 to 12:30',
             '07:45 / 17:30', '07:30-14:30 EST', '10:15a- 1:45p',
             '10:30 AM-12:30 PM', and pandas datetime.time objects.
    """
    import datetime as dt
    if pd.isna(raw):
        return raw
    # Pandas may parse a lone Excel time cell as datetime.time
    if isinstance(raw, dt.time):
        return normalize_time_str(raw)
    raw = str(raw).strip()
    raw = re.sub(r'\s+to\s+', '-', raw, flags=re.IGNORECASE)
    raw = re.sub(r'\s*/\s*', '-', raw)
    raw = re.sub(r'\s*-\s*', '-', raw)
    parts = raw.split('-')
    if len(parts) == 2:
        return f"{normalize_time_str(parts[0])} - {normalize_time_str(parts[1])}"
    return raw  # fallback


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
    # Any string that is only dashes (-, --, ---) counts as blank
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

# ── ⑧ Filter DS1 to survey year only ──────────────────────────────────────────
# The 2016 sheet contains 2 stray 2017 rows entered in the wrong file:
#   Long Key 2017-02-08 (group 2, 3 PIPL) and Mizell-Johnson 2017-02-07 (1 grp).
date_col_raw = cols["date"]
if date_col_raw in df.columns:
    _orig_len = len(df)
    df = df[pd.to_datetime(df[date_col_raw], errors='coerce').dt.year == int(year)].copy().reset_index(drop=True)
    _dropped = _orig_len - len(df)
    if _dropped > 0:
        print(f"  [FIX ⑧] Dropped {_dropped} DS1 row(s) with year ≠ {year} (stray entries from another survey year)")

# ── ① Carry forward "same" coordinates before parsing ─────────────────────────
lat_src = cols["latitude"]
lon_src = cols["longitude"]

same_count = 0
for i in range(1, len(df)):
    idx     = df.index[i]
    idx_prev = df.index[i - 1]
    for col in [lat_src, lon_src]:
        if str(df.at[idx, col]).strip().lower() == 'same':
            prev_val = df.at[idx_prev, col]
            df.at[idx, col] = prev_val
            same_count += 1
            if same_count <= 5:   # only print first few to avoid noise
                print(f"    [FIX ①] Row {i} '{col}' = 'same' → carried forward: {prev_val}")

if same_count > 0:
    print(f"  {same_count} 'same' coordinate value(s) carried forward from row above")

# ── Parse coordinates ──────────────────────────────────────────────────────────
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
# Read DS3 (band resights) and apply 2016 fixes
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

# ── ② Fix DS3 date typo: 2015-02-08 → 2016-02-08 ─────────────────────────────
date_col = bcols["date"]
wrong_year_mask = bands[date_col].astype(str).str.startswith("2015")
if wrong_year_mask.any():
    bands.loc[wrong_year_mask, date_col] = pd.to_datetime(
        bands.loc[wrong_year_mask, date_col].astype(str).str.replace("2015", "2016", n=1)
    )
    print(f"  [FIX ②] Corrected {wrong_year_mask.sum()} DS3 date(s): 2015 → 2016 (typo)")

# ── ③ Fix Bunche Beach: assign Group or Point # = 2 where it is NaN ───────────
grp_col   = bcols["group"]
loc_col   = bcols["location"]
bunche_mask = (
    bands[loc_col].astype(str).str.contains("Bunche Beach", case=False, na=False)
    & bands[grp_col].isna()
)
if bunche_mask.any():
    bands.loc[bunche_mask, grp_col] = 2
    print(f"  [FIX ③] Assigned Group or Point # = 2 to {bunche_mask.sum()} Bunche Beach DS3 row(s)")

# ── ④ Fix Hutchinson Island DS3 zone typo: YM-JJ → T-Y ───────────────────────
# Confirmed by biologist: DS3 entered wrong zone name, should match DS1 T-Y zone
hutchinson_mask = (
    bands[loc_col].astype(str).str.contains("Hutchinson", case=False, na=False)
    & bands[loc_col].astype(str).str.contains("YM", case=False, na=False)
)
if hutchinson_mask.any():
    correct_route = "S. Hutchinson Island Zones T - Y (Normandy Beach Access through County Line)"
    bands.loc[hutchinson_mask, loc_col] = correct_route
    print(f"  [FIX ④] Corrected DS3 Hutchinson Island zone: YM-JJ → T-Y ({hutchinson_mask.sum()} row(s))")

# ── Build BandCombo ───────────────────────────────────────────────────────────
# Rules:
#   • All 4 leg positions (UL, LL, UR, LR) always included
#   • Empty / NaN / any run of dashes ("---") → "-"
#   • Descriptive or "unreadable" text kept as-is (not parsable, not lost)
#   • Orientation prefix only added when present and not blank/dashes
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

# ④ Route names are normalized for matching (punctuation-insensitive)
#    but the original DS1 Route name is preserved in output.

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
        # No banded birds for this flock — keep DS1 row as-is
        expanded_rows.append(flock_row.to_dict())
    else:
        # One row per banded bird
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

        # One row for unbanded remainder
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

# ── ⑤ Handle unmatched DS3 rows ───────────────────────────────────────────────
# Routes listed here are flagged for removal with a reason (not kept as standalone).
# All other unmatched DS3 rows are included as standalone banded-bird rows.
DS3_REMOVAL_ROUTES = {
    "outback key": (
        "DS3 band resight row could not be matched to DS1 flock — "
        "DS1 Outback Key entry has no Group or Point # recorded and "
        "coordinates entered as 'same' (unresolvable)"
    ),
    "three rooker bar north island": (
        "DS3 band resight row could not be matched to DS1 flock — "
        "DS3 recorded group 2 or 3 but DS1 north island only has groups 4 & 5; "
        "DS2 confirms total = 10 (groups 4+5 only), so these banded birds are "
        "already counted within DS1; group numbers on DS3 are a recording error"
    ),
}

unmatched_bands = bands[~bands.index.isin(matched_band_indices)]
standalone_kept    = 0
standalone_removed = 0

for _, band_row in unmatched_bands.iterrows():
    band_comment = band_row.get("_band_comments", "")
    route_norm   = normalize_route(band_row.get("_band_route", ""))

    row_data = {
        "SurveyDate":    band_row.get("_band_date"),
        "Route":         band_row.get("_band_route"),
        "GroupNumber":   band_row.get("_band_group"),
        "Observer":      band_row.get("_band_observer", ""),
        "TotalObserved": 1,
        "FlagCode":      band_row.get("FlagCode", ""),
        "FlagColor":     band_row.get("FlagColor", ""),
        "BandCombo":     band_row.get("BandCombo", ""),
        "Comments":      str(band_comment).strip() if pd.notna(band_comment) else "",
    }

    # Check if this route is flagged for removal
    removal_reason = next(
        (reason for key, reason in DS3_REMOVAL_ROUTES.items() if key in route_norm),
        None
    )

    if removal_reason:
        row_data["_removal_reason"] = removal_reason
        standalone_removed += 1
    else:
        standalone_kept += 1

    expanded_rows.append(row_data)

if standalone_removed:
    print(f"\n  [INFO] {standalone_removed} unmatched DS3 row(s) flagged for removal (see DS3_REMOVAL_ROUTES)")
if standalone_kept:
    print(f"  [INFO] {standalone_kept} unmatched DS3 row(s) kept as standalone banded-bird row(s) with NaN GPS")

df = pd.DataFrame(expanded_rows)
print(f"\n  After Option A expansion: {len(df)} total rows")

# ── Keep only the columns we need ──────────────────────────────────────────────
keep = [
    "SurveyDate", "SurveyTime", "WeatherCondition",
    "Route", "TotalObserved", "GroupNumber", "Latitude", "Longitude",
    "Observer", "ObserverEmail", "FlagCode", "FlagColor", "BandCombo", "Comments",
    "_removal_reason",   # carries pre-flagged rows forward to Step 2
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
