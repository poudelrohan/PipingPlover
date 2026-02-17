config = {

    # ── Basic Info ─────────────────────────────────────────────────────────────
    "database_name": "database4",

    # ── Input ──────────────────────────────────────────────────────────────────
    "input": {
        "folder": "../../Databases/Database4",
        "files": ["Banded Birds.xlsx"],
        "header_row": 1,       # Real column names are in row 1, not row 0 (row 0 = biologist notes)
    },

    # ── Output ─────────────────────────────────────────────────────────────────
    "output": {
        "folder": "../../Databases/Database4/Output",
    },

    # ── Geography Validation ───────────────────────────────────────────────────
    # Rows outside Florida bounding box → removed, reason logged
    "geography": {
        "lat_column": "Latitude",
        "lon_column": "Longitude",
        "lat_min": 24.396308,
        "lat_max": 31.000968,
        "lon_min": -87.634938,
        "lon_max": -80.031362,
    },

    # ── Location Fields Rule ───────────────────────────────────────────────────
    # If ALL three are missing → remove row, log reason
    # If at least ONE is present → keep row, flag a warning in Summary_Report
    "location_fields": {
        "fields": ["LocationName", "Latitude", "Longitude"],
        "removal_reason": "Missing all location fields (LocationName, Latitude, Longitude)",
        "warn_if_partial": True,
    },

    # ── Required Fields ────────────────────────────────────────────────────────
    # If any of these are missing → remove row
    # Removal reason logged as: "Missing required field: <field_name>"
    "required_fields": [
        "ResightDate",
    ],

    # ── Columns to Keep ────────────────────────────────────────────────────────
    # Biologist-approved columns (from row 0 annotations: yes/Yes)
    # Dropped: ProjectID, StateProvince, ObserverEmail, dbo_Flocks.SpeciesID
    # Pending decision: ResightingMasterID, ResightingID, LocationID
    "columns_to_keep": [
        "ResightDate",
        "LocationName",
        "Latitude",
        "Longitude",
        "ObserverFirst",
        "ObserverLast",
        "dbo_ResightingMasters.Comments",
        "SpeciesID",
        "FlagID",
        "FlagCode",
        "UpperLeft",
        "LowerLeft",
        "UpperRight",
        "LowerRight",
        "dbo_Resightings.Comments",
        "FlockSize",
        "FlockActivityID",
    ],

    # ── Column Renames ─────────────────────────────────────────────────────────
    # Clean up ugly dbo_ prefixed column names
    "column_rename": {
        "dbo_ResightingMasters.Comments": "MasterComments",
        "dbo_Resightings.Comments":       "ResightingComments",
    },

    # ── Duplicate Criteria ─────────────────────────────────────────────────────
    # TODO: Define what makes two rows duplicates
    # Pending decision from user
    "duplicate_criteria": [],

}
