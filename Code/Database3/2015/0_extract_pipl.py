"""
Step 0: Extract PIPL Data from GPS Sheet
─────────────────────────────────────────
Reads the GPS sheet, extracts the PIPL count from the free-text focal
species column, normalizes column names to the standard pipeline format,
and outputs a flat file ready for Steps 1-6.

For years with a separate band resight sheet (band_sheet in config),
merges band data using Option A:
  - Each banded PIPL gets its own row (TotalObserved=1) with band details
  - Remaining unbanded birds get one row (TotalObserved=N-banded)
  - All rows share the same GPS coordinates from the flock/group

PIPL count extraction (priority order):
  1. "9 PIPL"                      → number before PIPL
  2. "PIPL (5)"                    → number in parens
  3. "PIPL (5 total, 1 banded)"    → first number in parens
  4. "PIPL - 8"                    → number after dash
  5. "PIPL 17"                     → number after space
  6. "Piping Plovers (12: ...)"    → full species name
  7. Bare "PIPL" / "Piping Plover" → fallback to num_individuals, then assume 1

Output: db3_<year>_extracted.xlsx  (in Output/<year>/)
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

def normalize_time_str(t):
    """Normalize a single time component (e.g. '1021', '9:40 am') to 'HH:MM'."""
    t = t.strip()
    # Remove timezone abbreviations (EST, CST, ...)
    t = re.sub(r'\s+[A-Z]{2,3}$', '', t).strip()
    # Remove am/pm suffix
    t = re.sub(r'\s*(am|pm)\s*$', '', t, flags=re.IGNORECASE).strip()
    # Fix semicolon typo used instead of colon (e.g. "2;45")
    t = t.replace(';', ':')

    if ':' in t:
        parts = t.split(':')
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return f"{h:02d}:{m:02d}"
    elif re.match(r'^\d{3,4}$', t):
        return f"{t[:-2].zfill(2)}:{t[-2:]}"
    return t  # fallback — return as-is


def normalize_survey_time(raw):
    """
    Normalize a survey time range to 'HH:MM - HH:MM' format.
    Handles formats like: '1021-1220', '9:40 am - 12:10 pm', '09:15 to 12:30',
    '07:45 / 17:30', '07:30-14:30 EST', '12:34-2;45', '10:30 AM-12:30 PM'.
    """
    if pd.isna(raw):
        return raw
    raw = str(raw).strip()

    # Normalize "to" and "/" separators to "-"
    raw = re.sub(r'\s+to\s+', '-', raw, flags=re.IGNORECASE)
    raw = re.sub(r'\s*/\s*', '-', raw)
    # Collapse any whitespace around "-" so we get a clean single delimiter
    raw = re.sub(r'\s*-\s*', '-', raw)

    parts = raw.split('-')
    if len(parts) == 2:
        start = normalize_time_str(parts[0])
        end   = normalize_time_str(parts[1])
        return f"{start} - {end}"

    return raw  # fallback — could not parse, return original


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
    if val == "":
        return None

    try:
        return float(val)
    except ValueError:
        pass

    # Pattern: "N 28.05048" or "W 082.81961" (letter prefix)
    m = re.match(r'^([NSEW])\s*([\d.]+)$', val, re.IGNORECASE)
    if m:
        direction = m.group(1).upper()
        num = float(m.group(2))
        if direction in ('S', 'W'):
            num = -num
        return num

    # Pattern: "27.42262N" or "-82.67205W" (letter suffix, no space)
    m = re.match(r'^(-?[\d.]+)\s*([NSEW])$', val, re.IGNORECASE)
    if m:
        num = float(m.group(1))
        direction = m.group(2).upper()
        if direction in ('S', 'W'):
            num = -abs(num)
        else:
            num = abs(num)
        return num

    # Pattern: "27 37.957'N" (degrees + decimal minutes)
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


# ══════════════════════════════════════════════════════════════════════════════
# Read GPS sheet
# ══════════════════════════════════════════════════════════════════════════════

df = pd.read_excel(
    input_path,
    sheet_name=config["sheet"],
    header=config["header_row"],
)
print(f"  Loaded {len(df)} rows from '{config['file']}' → sheet '{config['sheet']}'")

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
num_col   = cols.get("num_individuals")
has_num   = config.get("has_num_individuals", False)

if has_num and num_col and num_col in df.columns:
    df["TotalObserved"] = df.apply(
        lambda r: extract_pipl_count(r[focal_col], r[num_col]), axis=1
    )
else:
    df["TotalObserved"] = df[focal_col].apply(
        lambda text: extract_pipl_count(text)
    )

extracted  = df["TotalObserved"].notna().sum()
failed     = df["TotalObserved"].isna().sum()
total_pipl = int(df["TotalObserved"].sum()) if extracted > 0 else 0

print(f"\n  PIPL count extracted: {extracted} rows, {failed} failed")
print(f"  Total PIPL count: {total_pipl}")

if failed > 0:
    print(f"\n  [WARNING] Could not extract PIPL count from {failed} rows:")
    for _, r in df[df["TotalObserved"].isna()].iterrows():
        print(f"    '{r[focal_col]}'")

# ── Normalize column names ─────────────────────────────────────────────────────
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
# Band sheet merge (Option A) — only for years with a separate band sheet
# ══════════════════════════════════════════════════════════════════════════════

band_cfg = config.get("band_sheet")

if band_cfg:
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

    # Build BandCombo from leg positions.
    # All 4 positions (UL, LL, UR, LR) are always included; missing → "-".
    # Orientation prefix is only added when present.
    def build_band_combo(row):
        parts = []
        orientation = str(row.get(bcols["code_orientation"], "")).strip() if pd.notna(row.get(bcols["code_orientation"])) else ""
        if orientation and orientation != "-":
            parts.append(f"Orientation:{orientation}")
        for label, col_name in [("UL", bcols["upper_left"]), ("LL", bcols["lower_left"]),
                                 ("UR", bcols["upper_right"]), ("LR", bcols["lower_right"])]:
            val = str(row.get(col_name, "")).strip() if pd.notna(row.get(col_name)) else ""
            val = val if val and val != "-" else "-"
            parts.append(f"{label}:{val}")
        return "; ".join(parts)

    bands["BandCombo"] = bands.apply(build_band_combo, axis=1)

    # Rename band columns
    band_rename = {}
    for std, src in [
        ("_band_date",     bcols["date"]),
        ("_band_route",    bcols["location"]),
        ("_band_group",    bcols["group"]),
        ("FlagCode",       bcols["flag_code"]),
        ("FlagColor",      bcols.get("flag_color")),
        ("_band_comments", bcols.get("comments")),
    ]:
        if src and src in bands.columns:
            band_rename[src] = std
    bands = bands.rename(columns=band_rename)

    # ── Option A expansion ────────────────────────────────────────────────────
    expanded_rows = []

    for _, flock_row in df.iterrows():
        flock_date  = flock_row.get("SurveyDate")
        flock_route = str(flock_row.get("Route", "")).strip().lower()
        flock_group = flock_row.get("GroupNumber")
        flock_pipl  = flock_row.get("TotalObserved")

        matched_bands = bands[
            (bands["_band_date"].astype(str).str[:10] == str(flock_date)[:10]) &
            (bands["_band_route"].astype(str).str.strip().str.lower() == flock_route) &
            (bands["_band_group"].astype(str) == str(flock_group))
        ]

        n_banded = len(matched_bands)

        if n_banded == 0:
            expanded_rows.append(flock_row.to_dict())
        else:
            # One row per banded bird
            for _, band_row in matched_bands.iterrows():
                row_data = flock_row.to_dict()
                row_data["TotalObserved"] = 1
                row_data["FlagCode"]  = band_row.get("FlagCode", "")
                row_data["FlagColor"] = band_row.get("FlagColor", "")
                row_data["BandCombo"] = band_row.get("BandCombo", "")
                band_comment = band_row.get("_band_comments", "")
                if pd.notna(band_comment) and str(band_comment).strip():
                    existing = str(row_data.get("Comments", "")) if pd.notna(row_data.get("Comments")) else ""
                    row_data["Comments"] = f"{existing} | Band: {band_comment}" if existing else f"Band: {band_comment}"
                expanded_rows.append(row_data)

            # One row for unbanded remainder
            unbanded_count = (flock_pipl - n_banded) if pd.notna(flock_pipl) else 0
            if unbanded_count > 0:
                row_data = flock_row.to_dict()
                row_data["TotalObserved"] = unbanded_count
                row_data["FlagCode"]  = ""
                row_data["FlagColor"] = ""
                row_data["BandCombo"] = ""
                expanded_rows.append(row_data)
            elif unbanded_count < 0:
                print(f"    [WARNING] More banded ({n_banded}) than PIPL count ({flock_pipl}) "
                      f"at {flock_route} grp {flock_group} on {str(flock_date)[:10]}")

    df = pd.DataFrame(expanded_rows)
    print(f"\n  After Option A expansion: {len(df)} rows "
          f"(was {extracted} flocks)")


# ── Keep only the columns we need ──────────────────────────────────────────────
keep = ["SurveyDate", "SurveyTime", "WeatherCondition",
        "Route", "TotalObserved", "GroupNumber", "Latitude", "Longitude",
        "Observer", "ObserverEmail", "FlagCode", "FlagColor", "BandCombo", "Comments"]
keep = [c for c in keep if c in df.columns]
df = df[keep]

# ── Save output ────────────────────────────────────────────────────────────────
df.to_excel(output_path, index=False)

print(f"\n[DONE] Step 0 complete")
print(f"  Year    : {year}")
print(f"  Rows    : {len(df)}")
print(f"  Columns : {list(df.columns)}")
print(f"  Output  : {output_path}")
