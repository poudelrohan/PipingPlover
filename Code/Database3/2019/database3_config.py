"""
Database 3 Configuration — Winter Bird Survey 2019
────────────────────────────────────────────────────
Input : Databases/Database3BandReview/Winter Birds 2019 Band Review.xlsx
Output: Databases/Database3/Output/2019/

2019-specific notes vs 2018:
  - Sheet: "Form Responses 1" (Google Form export)
  - Route names have ", County=X" suffix → stripped in Step 0
  - Point column names are long strings (matched by regex in Step 0):
      Lat  : "Point N Latitude (in decimal degrees)..."
      Long : "Point N Longitude (in decimal degrees)..."
      PIPL : "N Number of Piping Plovers (at this location)"
      Band : "N Band/Flag Codes for Piping Plovers..."
  - Point 1 has NO longitude column in the Band Review or Clean file
    (the Google Form column was named "1 Longitude..." without "Point",
    so it was excluded during the clean step). Step 0 reads Point 1
    longitude directly from the raw file: Winter Birds 2019.xlsx col 63.
  - All 19 points present (unlike 2018 which was missing 16 and 18)
  - No observer email column in the 2019 form
  - New Smyrna Beach has two rows (same observer, same date, same bird U9):
      Row 1  time=10:30, no GPS → dropped as incomplete duplicate
      Row 2  time=09:15, has GPS → kept
  - FlagCode / FlagColor: left null (not parsed from raw band text)
  - Option A expansion: same as 2018
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2019",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input (Band Review file) ───────────────────────────────────────────────
    "input_folder": "../../../Databases/Database3BandReview",
    "file":         "Winter Birds 2019 Band Review.xlsx",
    "sheet":        "Form Responses 1",
    "header_row":   0,

    # ── Raw file — used only to read Point 1 longitude (missing from clean) ───
    "raw_folder": "../../../Databases/Database3",
    "raw_file":   "Winter Birds 2019.xlsx",
    "raw_sheet":  "Form Responses 1",

    # ── Metadata column names (exact match) ───────────────────────────────────
    "columns": {
        "route":    0,          # column index — header is very long
        "date":     "Date you did your survey",
        "time":     "Route Start Time",
        "temp":     "Weather:  temperature (optional)",
        "wind":     "Wind (optional)",
        "rain":     "Rain (optional)",
        "observer": "Names of observers",
        "email":    None,       # not present in 2019 form
        "comments": "Any additional survey comments (if you are doing a new survey, please include information about the habitat and general area of the points)",
    },

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
