"""
Database 3 Configuration — Winter Bird Survey 2013
────────────────────────────────────────────────────
Self-contained config for the 2013 year file only.
Run all 7 pipeline scripts from this folder.

Input : Databases/Database3Clean/Winter Birds '13 Clean.xlsx
Output: Databases/Database3/Output/2013/
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2013",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input File ─────────────────────────────────────────────────────────────
    "input_folder": "../../../Databases/Database3Clean",
    "file":         "Winter Birds '13 Clean.xlsx",
    "sheet":        "Species GPS",
    "header_row":   1,      # row 0 = biologist annotations, row 1 = column names

    # ── Column Mapping (raw → standard) ───────────────────────────────────────
    "columns": {
        "date":            "Date ",                  # note trailing space
        "location":        "Location (all lines of data need a location)",
        "focal_species":   "Focal species (PIPL, SNPL, REKN, WIPL, other banded birds)",
        "num_individuals": "No. of individuals of same species or flock size",
        "latitude":        "(Individual bird or flock) Latitude",
        "longitude":       "(Individual bird or flock) Longitude",
        "observers":       "Observers",
        "email":           "Email address (of primary observer)",
        "flag_code":       "Flag color and code (REKN, AMOY, other)",
        "band_combo":      "Color band combo (Left leg top to bottom, Right leg top to bottom)",
        "comments":        "Comments",
    },

    "has_num_individuals": True,    # fallback column for PIPL count
    "group_column":        None,    # no group/point column in 2013
    "band_sheet":          None,    # no separate band sheet in 2013

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
        "fields":           ["Route", "Latitude", "Longitude"],
        "removal_reason":   "Missing all location fields (Route, Latitude, Longitude)",
        "warn_if_partial":  True,
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
    """Return the year-specific output folder: .../Output/2013/"""
    import os
    year = config["year"]
    base = os.path.normpath(os.path.join(script_dir, config["output"]["base_folder"]))
    folder = os.path.join(base, year)
    os.makedirs(folder, exist_ok=True)
    return folder


def get_filename(step_name):
    """Return a year-tagged filename, e.g. 'db3_2013_extracted.xlsx'"""
    return f"db3_{config['year']}_{step_name}.xlsx"
