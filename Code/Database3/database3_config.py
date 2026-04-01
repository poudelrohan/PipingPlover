"""
Database 3 Configuration — Winter Bird Surveys (per-year processing)
────────────────────────────────────────────────────────────────────
Each year file is processed independently through the same 6-step pipeline.
Sheet names and column names vary by year — this config maps them.

Currently configured for: 2015
"""

config = {

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name": "database3",

    # ── Year Being Processed ──────────────────────────────────────────────────
    # Change this to switch which year file the pipeline processes
    "active_year": "2014",

    # ── Output ─────────────────────────────────────────────────────────────────
    # Each year gets its own subfolder: Output/2013/, Output/2014/, etc.
    "output": {
        "base_folder": "../../Databases/Database3/Output",
    },

    # ── Source Database Label ─────────────────────────────────────────────────
    "source_database": "WinterBirdSurvey",

    # ── Per-Year File & Sheet Mapping ─────────────────────────────────────────
    # Each year has different sheet names and column names.
    # The pipeline reads from the CLEAN files (Database3Clean/).
    "years": {

        "2013": {
            "input_folder": "../../Databases/Database3Clean",
            "file": "Winter Birds '13 Clean.xlsx",
            "sheet": "Species GPS",
            "header_row": 1,               # row 0 = biologist annotations, row 1 = column names
            "columns": {
                "date":          "Date ",                  # note trailing space
                "location":      "Location (all lines of data need a location)",
                "focal_species": "Focal species (PIPL, SNPL, REKN, WIPL, other banded birds)",
                "num_individuals": "No. of individuals of same species or flock size",
                "latitude":      "(Individual bird or flock) Latitude",
                "longitude":     "(Individual bird or flock) Longitude",
                "observers":     "Observers",
                "email":         "Email address (of primary observer)",
                "flag_code":     "Flag color and code (REKN, AMOY, other)",
                "band_combo":    "Color band combo (Left leg top to bottom, Right leg top to bottom)",
                "comments":      "Comments",
            },
            "has_num_individuals": True,    # fallback column for PIPL count
            "group_column": None,          # no group/point column in 2013
        },

        "2014": {
            "input_folder": "../../Databases/Database3Clean",
            "file": "Winter Birds '14 Clean.xlsx",
            "sheet": "Indiv Flock GPS & bands",
            "header_row": 2,               # row 0 = biologist annotations, row 1 = instructions, row 2 = column names
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
            "has_num_individuals": False,   # no fallback column in 2014
            "group_column": "Group or Point #",  # biologist annotation: "yess" (keep)
            "band_sheet": None,              # band info is in same sheet in 2014
        },

        "2015": {
            "input_folder": "../../Databases/Database3Clean",
            "file": "Winter Birds '15 Clean.xlsx",
            "sheet": "DATA SHEET 1",           # flock GPS data
            "header_row": 2,                   # row 0 = biologist annotations, row 1 = instructions, row 2 = column names
            "columns": {
                "date":          "Date",
                "location":      "Route Name/ Description",
                "focal_species": "Species and number of individuals",
                "latitude":      "Group/Point Latitude",
                "longitude":     "Group/Point Longitude",
                "observers":     "Observer(s)",
                "email":         None,         # drop phone & email
                "flag_code":     None,         # band info is in DS3
                "band_combo":    None,         # band info is in DS3
                "comments":      "Comments",
                "time":          "Route Start & End Times",
                "weather":       "Weather Condition",
            },
            "has_num_individuals": False,
            "group_column": "Group or Point #",
            # ── Band resight sheet (DS3) — merged via Date + Route + Group # ──
            "band_sheet": {
                "sheet": "DATA SHEET 3",
                "header_row": 1,
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
        },
    },

    # ── Geography Validation ──────────────────────────────────────────────────
    # Bounding box for Florida (same as Database 2)
    "geography": {
        "lat_column": "Latitude",
        "lon_column": "Longitude",
        "lat_min": 24.0,
        "lat_max": 31.5,
        "lon_min": -88.0,
        "lon_max": -79.5,
    },

    # ── Location Fields Rule ──────────────────────────────────────────────────
    "location_fields": {
        "fields": ["Route", "Latitude", "Longitude"],
        "removal_reason": "Missing all location fields (Route, Latitude, Longitude)",
        "warn_if_partial": True,
    },

    # ── Required Fields ───────────────────────────────────────────────────────
    "required_fields": [
        "SurveyDate",
    ],

    # ── Columns to Keep & Final Column Order ──────────────────────────────────
    # After Step 0 normalizes column names, these are the standardized names
    "columns_to_keep": [
        # ── Row Identifier ────────────────
        "unique_id",

        # ── Survey Event ──────────────────
        "SurveyDate",
        "SurveyTime",
        "WeatherCondition",

        # ── Location ─────────────────────
        "Route",
        "Latitude",
        "Longitude",

        # ── Observation ──────────────────
        "GroupNumber",
        "TotalObserved",

        # ── Observer ─────────────────────
        "Observer",
        "ObserverEmail",

        # ── Banding ──────────────────────
        "FlagCode",
        "FlagColor",
        "BandCombo",

        # ── Notes ────────────────────────
        "Comments",

        # ── Tracking ─────────────────────
        "source_database",
        "source_file",
        "source_sheet",
    ],

    # ── Column Renames ────────────────────────────────────────────────────────
    # Step 0 already normalizes to these names, so no renames needed in Step 4
    "column_rename": {},

    # ── Duplicate Criteria ────────────────────────────────────────────────────
    "duplicate_criteria": ["Route", "Latitude", "Longitude", "SurveyDate", "TotalObserved", "FlagCode"],

}


def get_output_folder(script_dir):
    """Return the year-specific output folder path, e.g. .../Output/2013/"""
    import os
    year = config["active_year"]
    base = os.path.normpath(os.path.join(script_dir, config["output"]["base_folder"]))
    folder = os.path.join(base, year)
    os.makedirs(folder, exist_ok=True)
    return folder


def get_filename(step_name):
    """Return a year-tagged filename, e.g. 'db3_2013_extracted.xlsx'"""
    year = config["active_year"]
    return f"db3_{year}_{step_name}.xlsx"
