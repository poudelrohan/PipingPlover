"""
Database 3 Configuration — Winter Bird Survey 2024
────────────────────────────────────────────────────
Input : Databases/Database3BandReview/Winter Birds 2024 Band Review.xlsx
Output: Databases/Database3/Output/2024/

2024-specific notes:
  - Same Band Review structure as 2020/2021/2022 (78 cols, 19 points).
  - All Species sheet has a SHIFTED column layout:
      • "Transect" header expanded to
        "Transect (if yours is missing, email Beth Forys: forysea@eckerd.edu)"
      • New col "If you did not survey - let us know why" inserted at col 2.
      • Date/observer/time/weather columns shifted right by 2 positions.
    The metadata column NAMES are still the same except for Transect, so
    we just update the route header here. Code lookups by name still work.
  - Biologist review file: Biologist_Band_Review_Completed_2024.xlsx, sheet "2024".
  - All 28 Band Review routes join cleanly to All Species (no aliases needed).
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2024",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input (Band Review file) ───────────────────────────────────────────────
    "input_folder": "../../../Databases/Database3BandReview",
    "file":         "Winter Birds 2024 Band Review.xlsx",
    "sheet":        "Focal Observations",
    "header_row":   0,

    # ── Metadata source (All Species sheet of raw file) ───────────────────────
    "metadata_folder": "../../../Databases/Database3",
    "metadata_file":   "Winter Birds 2024.xlsx",
    "metadata_sheet":  "All Species",

    # ── Metadata column names (exact match against All Species headers) ───────
    # Note: 2024's Transect column has a long instructional suffix in the
    # header — we still match by exact column NAME, not index.
    "columns": {
        "route":    "Transect (if yours is missing, email Beth Forys: forysea@eckerd.edu)",
        "date":     "Date you did your survey",
        "time":     "Route Start Time",
        "temp":     "Weather:  temperature (optional)",   # double space preserved
        "wind":     "Wind (optional)",
        "rain":     "Rain (optional)",
        "observer": "Names of observers",
        "email":    "Email of primary observer",
        "comments": None,
    },

    # ── Route aliases (none needed for 2024) ──────────────────────────────────
    "route_aliases": {},

    # ── Points (all 19 present) ────────────────────────────────────────────────
    "points": list(range(1, 20)),

    # ── Output ─────────────────────────────────────────────────────────────────
    "output": {
        "base_folder": "../../../Databases/Database3/Output",
    },

    # ── Geography Validation ──────────────────────────────────────────────────
    "geography": {
        "lat_column": "Latitude",
        "lon_column":  "Longitude",
        "lat_min": 24.0,
        "lat_max": 31.5,
        "lon_min": -88.0,
        "lon_max": -79.5,
    },

    # ── Location Fields Rule ──────────────────────────────────────────────────
    "location_fields": {
        "fields":          ["Route", "Latitude", "Longitude"],
        "removal_reason":  "Missing all location fields (Route, Latitude, Longitude)",
        "warn_if_partial": True,
    },

    # ── Required Fields ───────────────────────────────────────────────────────
    "required_fields": ["SurveyDate"],

    # ── Final Column Order ────────────────────────────────────────────────────
    "columns_to_keep": [
        "unique_id",
        "SurveyDate",
        "SurveyTime",
        "WeatherCondition",
        "Route",
        "Latitude",
        "Longitude",
        "GroupNumber",
        "TotalObserved",
        "Observer",
        "ObserverEmail",
        "FlagCode",
        "FlagColor",
        "BandCombo",
        "Comments",
        "source_database",
        "source_file",
        "source_sheet",
    ],

    # ── Column Renames ────────────────────────────────────────────────────────
    "column_rename": {},

    # ── Duplicate Criteria ────────────────────────────────────────────────────
    "duplicate_criteria": [
        "Route", "Latitude", "Longitude", "SurveyDate",
        "TotalObserved", "FlagCode", "FlagColor", "BandCombo",
    ],
}


def get_output_folder(script_dir):
    import os
    year   = config["year"]
    base   = os.path.normpath(os.path.join(script_dir, config["output"]["base_folder"]))
    folder = os.path.join(base, year)
    os.makedirs(folder, exist_ok=True)
    return folder


def get_filename(step_name):
    return f"db3_{config['year']}_{step_name}.xlsx"
