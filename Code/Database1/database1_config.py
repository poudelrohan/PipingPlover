config = {

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name": "database1",

    # ── Input ──────────────────────────────────────────────────────────────────
    "input": {
        "folder": "../../Databases/Database1",
        "files": ["Ebird PIPL Data.xlsx"],
        "header_row": 1,       # Real column names are in row 1, not row 0 (row 0 = biologist notes)
        "sheet": "ebd_US-FL_pipplo_201001_202508_",   # Only this sheet has data; Sheet1 is empty
    },

    # ── Output ─────────────────────────────────────────────────────────────────
    "output": {
        "folder": "../../Databases/Database1/Output",
    },

    # ── Source Database Label ───────────────────────────────────────────────────
    # Added as a column to every row so data origin is clear after merging databases
    "source_database": "eBird",

    # ── Geography Validation ───────────────────────────────────────────────────
    # Uses actual Florida state boundary (shapefile) + 1 km coastal buffer
    # so beach / barrier-island points are not falsely rejected
    "geography": {
        "lat_column": "LATITUDE",
        "lon_column": "LONGITUDE",
        "shapefile": "../../Databases/Shapefiles/us_states/tl_2023_us_state.shp",
        "state_fips": "12",            # Florida FIPS code
        "buffer_meters": 1000,         # 1 km beyond the coastline
    },

    # ── Location Fields Rule ───────────────────────────────────────────────────
    # If ALL three are missing → remove row, log reason
    # If at least ONE is present → keep row, flag a warning in Summary_Report
    "location_fields": {
        "fields": ["LOCALITY", "LATITUDE", "LONGITUDE"],
        "removal_reason": "Missing all location fields (LOCALITY, LATITUDE, LONGITUDE)",
        "warn_if_partial": True,
    },

    # ── Required Fields ────────────────────────────────────────────────────────
    # Standard required fields — missing → remove row
    # Removal reason logged as: "Missing required field: <field_name>"
    # Note: OBSERVATION DATE is handled here
    "required_fields": [
        "OBSERVATION DATE",
    ],

    # ── Observation Count Special Rule ─────────────────────────────────────────
    # OBSERVATION COUNT = "x" → unknown count, keep the row (do NOT remove)
    # OBSERVATION COUNT = blank/null → remove row
    # Removal reason: "Missing required field: OBSERVATION COUNT"
    "observation_count": {
        "column":          "OBSERVATION COUNT",
        "keep_value":      "x",       # eBird convention: x = species present, count unknown
        "removal_reason":  "Missing required field: OBSERVATION COUNT",
    },

    # ── Columns to Keep & Final Column Order ───────────────────────────────────
    # Only columns explicitly annotated by biologist (y/Y) are kept
    # Unannotated columns = not needed (drop)
    # Order: ID → Sighting Event → Observation → Notes → Tracking
    "columns_to_keep": [
        # ── Row Identifier (always first) ───────
        "unique_id",

        # ── Sighting Event ──────────────────────
        "OBSERVATION DATE",
        "LOCALITY",
        "LATITUDE",
        "LONGITUDE",

        # ── Observation ─────────────────────────
        "COMMON NAME",
        "OBSERVATION COUNT",
        "TIME OBSERVATIONS STARTED",

        # ── Notes ───────────────────────────────
        "CHECKLIST COMMENTS",
        "SPECIES COMMENTS",

        # ── Tracking (always last) ───────────────
        "source_database",
        "source_file",
        "source_sheet",
    ],

    # ── Column Renames ─────────────────────────────────────────────────────────
    # Clean up ALL CAPS eBird column names to readable titles
    "column_rename": {
        "OBSERVATION DATE":           "ObservationDate",
        "LOCALITY":                   "LocationName",
        "LATITUDE":                   "Latitude",
        "LONGITUDE":                  "Longitude",
        "COMMON NAME":                "CommonName",
        "OBSERVATION COUNT":          "ObservationCount",
        "TIME OBSERVATIONS STARTED":  "TimeStarted",
        "CHECKLIST COMMENTS":         "ChecklistComments",
        "SPECIES COMMENTS":           "SpeciesComments",
    },

    # ── Duplicate Criteria ─────────────────────────────────────────────────────
    # Two rows are duplicates if ALL of these match (comments excluded intentionally)
    # Two rows can have different comments and still be the same sighting
    # Null values in TIME OBSERVATIONS STARTED are treated as equal
    "duplicate_criteria": [
        "LATITUDE",
        "LONGITUDE",
        "OBSERVATION DATE",
        "TIME OBSERVATIONS STARTED",
        "COMMON NAME",
        "OBSERVATION COUNT",
    ],

}
