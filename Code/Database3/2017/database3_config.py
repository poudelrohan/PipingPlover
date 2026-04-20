"""
Database 3 Configuration — Winter Bird Survey 2017
────────────────────────────────────────────────────
Self-contained config for the 2017 year file only.
Run all 7 pipeline scripts from this folder.

Input : Databases/Database3Clean/Winter Birds '17 Clean.xlsx
Output: Databases/Database3/Output/2017/

Notes vs 2016:
  - DS1 header_row=0  (2016 was 1) — first row is column names, no instructions row
  - DS3 header_row=0  (2016 was 1)
  - Column names are otherwise identical to 2016
  - 2017-specific fixes handled in 0_extract_pipl.py:
      * Lanark Reef date has trailing backtick ('02/08/2017`') → stripped before parsing
      * Outback Key group 2 longitude entered as -82751911 (missing decimal) → -82.751911
      * DS1 route "Dunlaton Bridge to Ponce Inlet" is a typo → corrected for DS1↔DS3 matching
      * Wild Goose Lagoon DS3 trailing newline handled automatically by normalize_route
      * Highland Beach and New Smyrna Beach positive longitudes auto-corrected in Step 2
      * Pavilion Key DS3 rows have no DS1 match → included as standalone banded-bird rows
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2017",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input File (DATA SHEET 1 — flock GPS) ─────────────────────────────────
    "input_folder": "../../../Databases/Database3Clean",
    "file":         "Winter Birds '17 Clean.xlsx",
    "sheet":        "DATA SHEET 1",
    "header_row":   0,      # row 0 = column names (no instructions row in 2017)

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
        "header_row": 0,        # row 0 = column names (no instructions row in 2017)
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
    """Return the year-specific output folder: .../Output/2017/"""
    import os
    year = config["year"]
    base = os.path.normpath(os.path.join(script_dir, config["output"]["base_folder"]))
    folder = os.path.join(base, year)
    os.makedirs(folder, exist_ok=True)
    return folder


def get_filename(step_name):
    """Return a year-tagged filename, e.g. 'db3_2017_extracted.xlsx'"""
    return f"db3_{config['year']}_{step_name}.xlsx"
