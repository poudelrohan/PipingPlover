"""
Database 3 Configuration — Winter Bird Survey 2018
────────────────────────────────────────────────────
Input : Databases/Database3BandReview/Winter Birds 2018 Band Review.xlsx
Output: Databases/Database3/Output/2018/

2018-specific notes:
  - Single sheet "Sheet1", wide format (one row per route, groups spread as columns)
  - No DS3 band resight sheet — band info already parsed into {N}PIPLbands columns
    in the Band Review file by parse_bands_2018.py
  - Points present: 1–15, 17, 19  (points 16 and 18 absent from column structure)
  - Column typo: '1PIPL ands' instead of '1PIPLbands' — handled in Step 0
  - GPS: use {N}Lat / {N}Long (point-level), NOT Start Lat / Start Long
  - Option A expansion:  k banded-bird rows (TotalObserved=1)  +
                         1 unbanded-remainder row (TotalObserved = PIPL − k)
  - FlagCode / FlagColor: left null — not parsed from raw band text
  - Known GPS issue: Huguenot Memorial Park point 3 has no GPS → borrowed from
    2024 Focal Observations data automatically in Step 0
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2018",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input (Band Review file produced by parse_bands_2018.py) ──────────────
    "input_folder": "../../../Databases/Database3BandReview",
    "file":         "Winter Birds 2018 Band Review.xlsx",
    "sheet":        "Sheet1",
    "header_row":   0,

    # ── Column names in the Band Review file ──────────────────────────────────
    "columns": {
        "route":       "Transect",
        "date":        "Date you did your survey",
        "time":        "Route Start Time",
        "temp":        "Weather: temperature (optional)",
        "wind":        "Wind (optional)",
        "rain":        "Rain (optional)",
        "observers":   "Names of observers",
        "email":       "Lead observer's email",
        "comments":    "Any additional survey comments?",
    },

    # ── Points present (16 and 18 are absent from column structure) ───────────
    "points": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 19],

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
