"""
Database 3 Configuration — Winter Bird Survey 2016
────────────────────────────────────────────────────
Self-contained config for the 2016 year file only.
Run all 7 pipeline scripts from this folder.

Input : Databases/Database3Clean/Winter Birds '16 Clean.xlsx
Output: Databases/Database3/Output/2016/

Notes vs 2015:
  - DS1 header_row=1  (2015 was 2)
  - DS3 header_row=1  (2015 was 2)
  - Column names are otherwise identical to 2015
  - DS1 has extra Route Start/End lat-lon columns — ignored, we use Group/Point lat-lon
  - Special fixes handled in 0_extract_pipl.py:
      * "same" in lat/lon → carry forward from row above
      * DS3 date typo 2015-02-08 → corrected to 2016-02-08
      * DS3 Bunche Beach Preserve: Group or Point # NaN → assigned 2
      * Route name normalization for DS1↔DS3 matching (punctuation-insensitive)
      * Unmatched DS3 rows included as standalone banded-bird rows (NaN GPS)
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2016",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input File (DATA SHEET 1 — flock GPS) ─────────────────────────────────
    "input_folder": "../../../Databases/Database3Clean",
    "file":         "Winter Birds '16 Clean.xlsx",
    "sheet":        "DATA SHEET 1",
    "header_row":   1,      # row 0 = instructions, row 1 = column names

    # ── Column Mapping — DATA SHEET 1 (raw → standard) ────────────────────────
    "columns": {
        "date":          "Date",
        "location":      "Route Name/ Description",
        "focal_species": "Species and number of individuals",
        "latitude":      "Group/Point Latitude",
        "longitude":     "Group/Point Longitude",
        "observers":     "Observer(s)",
        "email":         None,          # phone & email combined in one column, not separated
        "flag_code":     None,          # band info is in DS3
        "band_combo":    None,          # band info is in DS3
        "comments":      "Comments",
        "time":          "Route Start & End Times",
        "weather":       "Weather Condition",
    },

    "has_num_individuals": False,
    "group_column":        "Group or Point #",

    # ── Band Resight Sheet (DATA SHEET 3) ─────────────────────────────────────
    "band_sheet": {
        "sheet":      "DATA SHEET 3",
        "header_row": 1,        # row 0 = data, row 1 = column names (same sheet structure as 2016 DS1)
        "columns": {
            "date":             "Date",
            "location":         "Route Name/ Description",
            "group":            "Group or Point #",
            "observer":         "Observer(s)",
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
    """Return the year-specific output folder: .../Output/2016/"""
    import os
    year = config["year"]
    base = os.path.normpath(os.path.join(script_dir, config["output"]["base_folder"]))
    folder = os.path.join(base, year)
    os.makedirs(folder, exist_ok=True)
    return folder


def get_filename(step_name):
    """Return a year-tagged filename, e.g. 'db3_2016_extracted.xlsx'"""
    return f"db3_{config['year']}_{step_name}.xlsx"
