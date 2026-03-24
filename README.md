# USFWS Piping Plover Database Cleaning Pipeline

Data cleaning pipeline for Florida Piping Plover (PIPL) observation databases. Each database goes through a 6-step process that produces a final Excel report with clean data, removed rows, and summary statistics.

## Where Are the Final Clean Files?

| Database | Description | Final File | Clean Rows | Removed Rows |
|----------|-------------|------------|------------|--------------|
| Database 1 | eBird observations | `Databases/Database1/Output/database1_FINAL.xlsx` | 25,392 | 6,798 |
| Database 2 | Non-breeding survey (2011-2023) | `Databases/Database2/Output/database2_FINAL.xlsx` | 5,937 | 82 |
| Database 3 | Winter Bird Survey (2013-2024) | *In progress* | — | — |
| Database 4 | Banded bird resights | `Databases/Database4/Output/database4_FINAL.xlsx` | 192 | 3 |

Each final file (`_FINAL.xlsx`) has 3 sheets:

- **Clean_Data** — All valid observations after cleaning
- **Removed_Rows** — Rows that were removed, with a `_removal_reason` column explaining why
- **Summary_Report** — Statistics about the cleaning process (row counts, removal breakdown)

## Folder Structure

```
Databases/
  Database1/                  ← eBird PIPL observations
    Ebird PIPL Data.xlsx         Source file
    Output/                      Pipeline output files
      database1_FINAL.xlsx       ** FINAL CLEAN FILE **

  Database2/                  ← Non-breeding PIPL survey data
    NonBreedingPIPL_FL2011-2023.xlsx   Source file
    Output/
      database2_FINAL.xlsx       ** FINAL CLEAN FILE **

  Database3/                  ← Winter Bird Survey (raw yearly files)
    Winter Birds '13.xlsx        2013 raw data
    Winter Birds '14.xlsx        2014 raw data
    ...                          (12 files total, 2013-2024)

  Database3Clean/             ← Winter Bird Survey (PIPL-only filtered)
    Winter Birds '13 Clean.xlsx  2013 filtered to PIPL rows only
    Winter Birds '14 Clean.xlsx  2014 filtered
    Winter Birds '15 Clean.xlsx  2015 filtered
    Winter Birds '16 Clean.xlsx  2016 filtered

  Database4/                  ← Banded bird resight data
    Banded Birds.xlsx            Source file
    Output/
      database4_FINAL.xlsx       ** FINAL CLEAN FILE **

  OriginalData/               ← Backup copies of all original source files

Code/
  Database1/                  ← Pipeline scripts for Database 1
  Database2/                  ← Pipeline scripts for Database 2
  Database3/                  ← Exploration scripts for Database 3 (in progress)
  Database4/                  ← Pipeline scripts for Database 4
```

## What the Pipeline Does (6 Steps)

Each database runs through the same 6-step cleaning process:

| Step | Script | What It Does |
|------|--------|-------------|
| 1 | `1_add_ids.py` | Adds tracking columns: unique_id, source_database, source_file, source_sheet |
| 2 | `2_validate_geography.py` | Checks coordinates are within Florida boundaries, auto-corrects longitude sign errors |
| 3 | `3_check_required_fields.py` | Removes rows missing required fields (latitude, longitude, location name, date, etc.) |
| 4 | `4_select_columns.py` | Keeps only the columns we need, renames and reorders them |
| 5 | `5_remove_duplicates.py` | Removes duplicate observations, logs which row each duplicate matches |
| 6 | `6_create_final_report.py` | Produces the 3-sheet Excel file (Clean_Data, Removed_Rows, Summary_Report) |

## Database 3 Status

Database 3 (Winter Bird Survey) is more complex than the others because each yearly file has two related sheets:

- **Species GPS sheet** — Individual flock sightings with precise lat/lon coordinates
- **Counts sheet** — Survey route totals with route-level start/end coordinates

These need to be matched and merged before the standard pipeline can run. Exploration and pattern analysis is in progress. See `Code/Database3/` for current work.
