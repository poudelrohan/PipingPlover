"""
Database 3 Configuration — Winter Bird Survey 2015
────────────────────────────────────────────────────
Self-contained config for the 2015 year file only.
Run all 7 pipeline scripts from this folder.

Input : Databases/Database3Clean/Winter Birds '15 Clean.xlsx
Output: Databases/Database3/Output/2015/

Notes:
  - Two sheets: DATA SHEET 1 (flock GPS) + DATA SHEET 3 (band resights)
  - DS1 and DS3 are merged via Option A expansion (Date + Route + Group #)
  - Option A: one row per banded bird (TotalObserved=1) + one row for unbanded remainder
  - 128 total rows after expansion: 69 DS1 flocks + 74 DS3 band resights → 128 rows
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2015",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input File (DATA SHEET 1 — flock GPS) ─────────────────────────────────
    "input_folder": "../../../Databases/Database3Clean",
    "file":         "Winter Birds '15 Clean.xlsx",
    "sheet":        "DATA SHEET 1",
    "header_row":   2,      # row 0 = biologist annotations, row 1 = instructions, row 2 = column names

    # ── Column Mapping — DATA SHEET 1 (raw → standard) ────────────────────────
    "columns": {
        "date":          "Date",
        "location":      "Route Name/ Description",
        "focal_species": "Species and number of individuals",
        "latitude":      "Group/Point Latitude",
        "longitude":     "Group/Point Longitude",
        "observers":     "Observer(s)",
        "email":         None,      # drop phone & email (not in DS1)
        "flag_code":     None,      # band info is in DS3
        "band_combo":    None,      # band info is in DS3
        "comments":      "Comments",
        "time":          "Route Start & End Times",
        "weather":       "Weather Condition",
    },

    "has_num_individuals": False,
    "group_column":        "Group or Point #",

    # ── Band Resight Sheet (DATA SHEET 3) ─────────────────────────────────────
    # Merged into DS1 via Option A: Date + Route + Group #
    "band_sheet": {
        "sheet":      "DATA SHEET 3",
        "header_row": 2,
        "columns": {
            "date":             "Date",
            "location":         "Route Name/ Description",
            "group":            "Group or Point #",
            "species":          "Species",
            "flag_code":        "Band/Flag Code",
            "flag_color":       "Band/Flag Color",
            "code_orientation": "Code Orientation",
            "upper_left":       "Upper Left",
            "lower_left":       "Lower Left",
            "upper_right":      "Upper Right",
            "lower_right":      "Lower Right",
            "comments":         "Comments",
        },
    },

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
    """Return the year-specific output folder: .../Output/2015/"""
    import os
    year = config["year"]
    base = os.path.normpath(os.path.join(script_dir, config["output"]["base_folder"]))
    folder = os.path.join(base, year)
    os.makedirs(folder, exist_ok=True)
    return folder


def get_filename(step_name):
    """Return a year-tagged filename, e.g. 'db3_2015_extracted.xlsx'"""
    return f"db3_{config['year']}_{step_name}.xlsx"
