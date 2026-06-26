"""
Database 3 Configuration — Winter Bird Survey 2014
────────────────────────────────────────────────────
Input : Databases/Database3Clean/Winter Birds '14 Clean.xlsx
Output: Databases/Database3/Output/2014/

2014-specific notes:
  - Source format: 'Indiv Flock GPS & bands' sheet (multi-species per row).
      Row 1: top-level instruction text (skip)
      Row 2: empty / continuation (skip)
      Row 3: real headers
      Row 4+: data
  - PIPL count must be extracted from free-text species column
    (e.g. "PIPL (5 TOTAL, 3 banded)").
  - Band info lives in a SINGLE column (col 7) and was historically misrouted
    into the FlagCode output column. This pipeline correctly writes it into
    BandCombo. FlagCode / FlagColor stay null for 2014.
  - GroupNumber is sourced from col 3 ("Group or Point #").
  - Biologist review file:
      Biologist_Band_Review_2013-2017.xlsx, sheet "2014"
"""

config = {
    "year": "2014",
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    "input_folder": "../../../Databases/Database3Clean",
    "file":         "Winter Birds '14 Clean.xlsx",
    "sheet":        "Indiv Flock GPS & bands",
    "header_row":   3,    # 1-indexed; data starts at row 4

    "biologist_review_file":  "Biologist_Band_Review_2013-2017.xlsx",
    "biologist_review_sheet": "2014",

    "columns": {
        "date":     0,
        "route":    1,
        "observer": 2,
        "point":    3,
        "lat":      4,
        "lon":      5,
        "species":  6,   # free text — parse PIPL count from this
        "bands":    7,   # single column with all band info (→ BandCombo)
        "photo":    8,
        "notes":    9,
    },

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
