config = {

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name": "database2",

    # ── Input ──────────────────────────────────────────────────────────────────
    "input": {
        "folder": "../../Databases/Database2",
        "files": ["NonBreedingPIPL_FL2011-2023.xlsx"],
        "header_row": 1,       # Real column names are in row 1, not row 0 (row 0 = biologist notes)
        "sheet": "PIPL Location Data",  # Only this sheet; "Locations" sheet not needed
    },

    # ── Output ─────────────────────────────────────────────────────────────────
    "output": {
        "folder": "../../Databases/Database2/Output",
    },

    # ── Source Database Label ─────────────────────────────────────────────────
    # Added as a column to every row so data origin is clear after merging databases
    "source_database": "NonBreedingPIPL",

    # ── Geography Validation ───────────────────────────────────────────────────
    # Uses a lat/lon bounding box for Florida instead of a shapefile.
    # The box is generous enough to cover the Gulf coast, Keys, and barrier islands.
    "geography": {
        "lat_column": "Latitude",
        "lon_column": "Longitude",
        "lat_min": 24.0,    # south of Key West
        "lat_max": 31.5,    # north of FL panhandle
        "lon_min": -88.0,   # west of FL panhandle
        "lon_max": -79.5,   # east coast
    },

    # ── Location Fields Rule ───────────────────────────────────────────────────
    # "if no location no lat no longitude" → if ALL three missing → remove row
    # If at least ONE is present → keep row, flag a warning in Summary_Report
    "location_fields": {
        "fields": ["Route", "Latitude", "Longitude"],
        "removal_reason": "Missing all location fields (Route, Latitude, Longitude)",
        "warn_if_partial": True,
    },

    # ── Required Fields ────────────────────────────────────────────────────────
    # "if empty then remove it"
    # Removal reason logged as: "Missing required field: <field_name>"
    "required_fields": [
        "Date",
    ],

    # ── Columns to Keep & Final Column Order ───────────────────────────────────
    # Based on biologist row 0 annotations:
    #   Keep: Route, Latitude, Longitude, Date, Time Sited, Tide,
    #         Foraging, Roosting, Habitat Type, Total Observed, Total Banded,
    #         Observer, Notes
    #   Drop: Species ("no"), Age ("no")
    # Order: ID → Survey Event → Location → Observation → Observer → Notes → Tracking
    "columns_to_keep": [
        # ── Row Identifier (always first) ───────
        "unique_id",

        # ── Survey Event ───────────────────────
        "Date",
        "Time Sited",

        # ── Location ──────────────────────────
        "Route",
        "Latitude",
        "Longitude",

        # ── Observation ────────────────────────
        "Group Number",
        "Tide",
        "Foraging",
        "Roosting",
        "Habitat Type",
        "Total Observed",
        "Total Banded",

        # ── Observer ───────────────────────────
        "Observer",

        # ── Notes ──────────────────────────────
        "Notes",

        # ── Tracking (always last) ──────────────
        "source_database",
        "source_file",
        "source_sheet",
    ],

    # ── Column Renames ─────────────────────────────────────────────────────────
    # Clean up column names with spaces → PascalCase
    "column_rename": {
        "Date":           "SurveyDate",
        "Time Sited":     "TimeSited",
        "Group Number":   "GroupNumber",
        "Habitat Type":   "HabitatType",
        "Total Observed": "TotalObserved",
        "Total Banded":   "TotalBanded",
    },

    # ── Duplicate Criteria ─────────────────────────────────────────────────────
    # Two rows are duplicates if ALL of these match:
    #   Route + Latitude + Longitude → same exact survey location
    #   Date                         → same day
    #   Group Number                 → same group (different groups = distinct observations)
    #   Total Observed               → same flock count
    # Note: Notes column excluded — rows with different notes can still be duplicates
    "duplicate_criteria": ["Route", "Latitude", "Longitude", "Date", "Group Number", "Total Observed"],

}
