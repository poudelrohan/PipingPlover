"""
Database 3 Configuration — Winter Bird Survey 2013
────────────────────────────────────────────────────
Input : Databases/Database3Clean/Winter Birds '13 Clean.xlsx
Output: Databases/Database3/Output/2013/

2013-specific notes:
  - Source format: single 'Species GPS' sheet (multi-species per row).
      Row 1: human annotations (skip)
      Row 2: real headers
      Row 3+: data
  - PIPL count must be extracted from free-text species column
    (e.g. "PIPL- 8, SNPL- 4, WIPL- 3").
  - Band info lives in BOTH col 9 (Flag color/code) and col 10 (Color band combo).
  - No Group/Point # column — every row is one location (GroupNumber stays null).
  - Biologist review file:
      Biologist_Band_Review_2013-2017.xlsx, sheet "2013"
"""

config = {
    "year": "2013",
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input (source clean file) ──────────────────────────────────────────────
    "input_folder": "../../../Databases/Database3Clean",
    "file":         "Winter Birds '13 Clean.xlsx",
    "sheet":        "Species GPS",
    "header_row":   2,    # 1-indexed; data starts at row 3

    # ── Biologist review file ──────────────────────────────────────────────────
    "biologist_review_file":  "Biologist_Band_Review_2013-2017.xlsx",
    "biologist_review_sheet": "2013",

    # ── Source column indices (0-based) ────────────────────────────────────────
    "columns": {
        "date":     0,
        "route":    1,
        "species":  2,   # free text — parse PIPL count from this
        "lat":      5,
        "lon":      6,
        "observer": 7,
        "email":    8,
        "flag":     9,   # Flag color/code column (sometimes "0" placeholder)
        "combo":    10,  # Color band combo column
        "comments": 11,
    },

    # ── Output ─────────────────────────────────────────────────────────────────
    "output": {"base_folder": "../../../Databases/Database3/Output"},

    "geography": {
        "lat_column": "Latitude", "lon_column": "Longitude",
        "lat_min": 24.0, "lat_max": 31.5,
        "lon_min": -88.0, "lon_max": -79.5,
    },

    "location_fields": {
        "fields": ["Route", "Latitude", "Longitude"],
        "removal_reason": "Missing all location fields (Route, Latitude, Longitude)",
        "warn_if_partial": True,
    },

    "required_fields": ["SurveyDate"],

    "columns_to_keep": [
        "unique_id", "SurveyDate", "SurveyTime", "WeatherCondition",
        "Route", "Latitude", "Longitude",
        "GroupNumber", "TotalObserved",
        "Observer", "ObserverEmail",
        "FlagCode", "FlagColor", "BandCombo",
        "Comments",
        "source_database", "source_file", "source_sheet",
    ],

    "column_rename": {},

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
