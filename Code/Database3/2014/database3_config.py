"""
Database 3 Configuration — Winter Bird Survey 2014
────────────────────────────────────────────────────
Self-contained config for the 2014 year file only.
Run all 7 pipeline scripts from this folder.

Input : Databases/Database3Clean/Winter Birds '14 Clean.xlsx
Output: Databases/Database3/Output/2014/

Notes:
  - header_row=2: row 0 = biologist annotations, row 1 = instructions, row 2 = column names
  - No email column in the GPS sheet
  - Band info is in the same sheet (no separate DS3)
  - 2 rows removed for missing SurveyDate
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2014",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input File ─────────────────────────────────────────────────────────────
    "input_folder": "../../../Databases/Database3Clean",
    "file":         "Winter Birds '14 Clean.xlsx",
    "sheet":        "Indiv Flock GPS & bands",
    "header_row":   2,      # row 0 = annotations, row 1 = instructions, row 2 = column names

    # ── Column Mapping (raw → standard) ───────────────────────────────────────
    "columns": {
        "date":          "Date",
        "location":      "Route Name",
        "focal_species": "Species and number of individuals",
        "latitude":      "(Individual bird or flock) Latitude in decimal degrees, e.g. 27.123456",
        "longitude":     "(Individual bird or flock) Longitude in decimal degrees e.g. -81.123456",
        "observers":     "Observer(s)",
        "email":         None,      # no email column in 2014 GPS sheet
        "flag_code":     "band information (list species and describe band colors for upper left, lower left, upper right and lower right)",
        "band_combo":    None,      # combined into flag_code in 2014
        "comments":      "NOTES",
    },

    "has_num_individuals": False,               # no fallback num_individuals column
    "group_column":        "Group or Point #",  # biologist annotation: "yess" (keep)
    "band_sheet":          None,                # band info is in the main GPS sheet

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
    # Step 5 silently skips any column not present in this year's data
    "duplicate_criteria": [
        "Route", "Latitude", "Longitude", "SurveyDate",
        "TotalObserved", "FlagCode", "FlagColor", "BandCombo",
    ],

}


def get_output_folder(script_dir):
    """Return the year-specific output folder: .../Output/2014/"""
    import os
    year = config["year"]
    base = os.path.normpath(os.path.join(script_dir, config["output"]["base_folder"]))
    folder = os.path.join(base, year)
    os.makedirs(folder, exist_ok=True)
    return folder


def get_filename(step_name):
    """Return a year-tagged filename, e.g. 'db3_2014_extracted.xlsx'"""
    return f"db3_{config['year']}_{step_name}.xlsx"
