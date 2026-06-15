"""
Database 3 Configuration — Winter Bird Survey 2021
────────────────────────────────────────────────────
Input : Databases/Database3BandReview/Winter Birds 2021 Band Review.xlsx
Output: Databases/Database3/Output/2021/

2021-specific notes (same shape as 2020):
  - Band Review file has clean short column names:
      Latitude (point N), Longitude (point N),
      Number of Piping Plovers (point N), PIPL Band/Flag Codes (point N)
  - Metadata (date, observer, time, weather, email) comes from the
    All Species sheet of Winter Birds 2021.xlsx, joined on route name.
  - All Species sheet has a 'Totals' row at the top — skipped.
  - All 19 points present, all GPS columns present.
  - Observer email present.
  - No survey comments column.
  - Biologist review file: Biologist_Band_Review_Completed_2021.xlsx, sheet "2021".
  - Some route names include parenthetical county suffix (e.g.
    "Anclote Key North (Pasco)") — preserved as-is per biologist guidance.
"""

config = {

    # ── Year ───────────────────────────────────────────────────────────────────
    "year": "2021",

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name":   "database3",
    "source_database": "WinterBirdSurvey",

    # ── Input (Band Review file) ───────────────────────────────────────────────
    "input_folder": "../../../Databases/Database3BandReview",
    "file":         "Winter Birds 2021 Band Review.xlsx",
    "sheet":        "Focal Observations",
    "header_row":   0,

    # ── Metadata source (All Species sheet of raw file) ───────────────────────
    "metadata_folder": "../../../Databases/Database3",
    "metadata_file":   "Winter Birds 2021.xlsx",
    "metadata_sheet":  "All Species",

    # ── Metadata column names (exact match against All Species headers) ───────
    "columns": {
        "route":    "Transect ",
        "date":     "Date you did your survey",
        "time":     "Route Start Time",
        "temp":     "Weather:  temperature (optional)",   # NOTE: double space
        "wind":     "Wind (optional)",
        "rain":     "Rain (optional)",
        "observer": "Names of observers",
        "email":    "Email of primary observer",
        "comments": None,
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
