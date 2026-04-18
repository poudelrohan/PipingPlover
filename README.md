# USFWS Piping Plover Database Cleaning Pipeline

Data cleaning pipeline for Florida Piping Plover (PIPL) observation databases. Each database goes through a structured cleaning process that produces a final Excel report with clean data, removed rows, and summary statistics.

## Where Are the Final Clean Files?

| Database | Description | Final File(s) | Status |
|----------|-------------|---------------|--------|
| Database 1 | eBird observations | `Databases/Database1/Output/database1_FINAL.xlsx` | ✅ Complete (25,392 clean rows) |
| Database 2 | Non-breeding survey (2011–2023) | `Databases/Database2/Output/database2_FINAL.xlsx` | ✅ Complete (5,937 clean rows) |
| Database 3 | Winter Bird Survey (2013–2024) | `Databases/Database3/Output/<year>/db3_<year>_FINAL.xlsx` | 🔄 In progress (2013–2016 done) |
| Database 4 | Banded bird resights | `Databases/Database4/Output/database4_FINAL.xlsx` | ✅ Complete (192 clean rows) |

Each final file (`_FINAL.xlsx`) has 3 sheets:

- **Clean_Data** — All valid observations after cleaning
- **Removed_Rows** — Rows that were removed, with a `removal_reason` column explaining why
- **Summary_Report** — Statistics about the cleaning process (row counts, removal breakdown)

---

## Database 3 — Winter Bird Survey (Per-Year Pipeline)

Database 3 is more complex than the others. Each yearly Excel file contains up to three sheets:

| Sheet | Contents |
|-------|----------|
| **DATA SHEET 1** | Individual flock sightings — one row per group, with GPS coordinates and PIPL count |
| **DATA SHEET 2** | Route-level observer summary — total PIPL per route (used for QA/verification) |
| **DATA SHEET 3** | Banded bird resights — one row per banded bird with flag/band detail |

### Merge Strategy (Option A Expansion)

DS1 and DS3 are merged by matching on **Date + Route + Group or Point #**:

- Each matched banded bird gets its own row (`TotalObserved = 1`) with full band details
- The remaining unbanded birds from that group get one row (`TotalObserved = flock_count − n_banded`)
- DS1 groups with no DS3 match are kept as-is
- Unmatched DS3 rows that cannot be resolved are flagged and moved to Removed_Rows

DS2 totals are used **after processing** as a cross-check to catch inflated counts from wrong-year rows or DS3 group number recording errors.

### Completed Years

| Year | Clean Rows | Removed | PIPL Total | Notes |
|------|-----------|---------|------------|-------|
| 2013 | — | — | — | Pipeline complete; outputs in `Output/2013/` |
| 2014 | — | — | — | Pipeline complete; outputs in `Output/2014/` |
| 2015 | — | — | — | Pipeline complete; SurveyTime normalized, BandCombo standardized |
| 2016 | 84 | 4 | 251 | DS2 audit verified; 2 wrong-year DS1 rows removed; Three Rooker Bar DS3 group mismatch resolved |

### 2016 Data Fixes Applied

The following issues were found and corrected during the 2016 build:

1. **"same" coordinates** — Rows where lat/lon was written as "same" carry forward the value from the row above before parsing
2. **DS3 date typo** — One DS3 row had date 2015-02-08; corrected to 2016-02-08
3. **Bunche Beach NaN group number** — DS3 row with blank Group # assigned group 2 to enable matching
4. **Hutchinson Island zone typo** — DS3 recorded zone "YM-JJ"; corrected to "T-Y" to match DS1 (confirmed biologist error)
5. **Wrong-year DS1 rows** — Two 2017 rows (Long Key, Mizell-Johnson) found in the 2016 sheet; removed by year filter
6. **Three Rooker Bar north island DS3 mismatch** — DS3 recorded banded birds at groups 2 & 3, but DS1 only has groups 4 & 5; DS2 confirms total = 10 (groups 4+5 only); unmatched DS3 rows removed as recording errors

### Per-Year Pipeline Structure

Every year has its own self-contained folder under `Code/Database3/<year>/`:

```
Code/Database3/
  2013/
    database3_config.py     ← Year-specific config (file name, header rows, column mapping)
    0_extract_pipl.py       ← Read DS1 + DS3, extract PIPL counts, merge via Option A
    1_add_ids.py            ← Assign unique_id and source tracking columns
    2_validate_geography.py ← Check/fix coordinates against Florida bounding box
    3_check_required_fields.py  ← Remove rows missing required fields
    4_select_columns.py     ← Rename, reorder, and keep only needed columns
    5_remove_duplicates.py  ← Deduplicate observations
    6_create_final_report.py    ← Build 3-sheet FINAL.xlsx
  2014/  (same structure)
  2015/  (same structure)
  2016/  (same structure)
```

Each year's config captures the year-specific differences (header row positions, column name variations, sheet structures, known data quirks) so the scripts can be run independently without affecting other years.

---

## Other Databases — Folder Structure

```
Databases/
  Database1/                     ← eBird PIPL observations
    Ebird PIPL Data.xlsx
    Output/
      database1_FINAL.xlsx       ← FINAL CLEAN FILE

  Database2/                     ← Non-breeding PIPL survey data
    NonBreedingPIPL_FL2011-2023.xlsx
    Output/
      database2_FINAL.xlsx       ← FINAL CLEAN FILE

  Database3/                     ← Winter Bird Survey raw yearly files
    Winter Birds '13.xlsx  ...   (raw survey files, 2013–2024)
    Output/
      2013/  2014/  2015/  2016/ ← Per-year pipeline outputs
        db3_<year>_FINAL.xlsx    ← FINAL CLEAN FILE (per year)

  Database3Clean/                ← Winter Bird Survey pre-filtered to PIPL rows
    Winter Birds '13 Clean.xlsx  ...

  Database4/                     ← Banded bird resight data
    Banded Birds.xlsx
    Output/
      database4_FINAL.xlsx       ← FINAL CLEAN FILE

Code/
  Database1/                     ← Pipeline scripts for Database 1
  Database2/                     ← Pipeline scripts for Database 2
  Database3/
    2013/ 2014/ 2015/ 2016/      ← Per-year self-contained pipelines (see above)
  Database4/                     ← Pipeline scripts for Database 4
```

## Standard Pipeline Steps

All databases run through the same core steps:

| Step | Script | What It Does |
|------|--------|--------------|
| 0 | `0_extract_pipl.py` | *(DB3 only)* Read DS1 + DS3, extract PIPL counts, merge banded/unbanded via Option A |
| 1 | `1_add_ids.py` | Add `unique_id`, `source_database`, `source_file`, `source_sheet` tracking columns |
| 2 | `2_validate_geography.py` | Verify coordinates are within Florida bounding box; auto-correct longitude sign errors |
| 3 | `3_check_required_fields.py` | Remove rows missing required fields (date, location) |
| 4 | `4_select_columns.py` | Keep only needed columns, rename and reorder to standard format |
| 5 | `5_remove_duplicates.py` | Remove exact duplicates, log which row each duplicate matched |
| 6 | `6_create_final_report.py` | Produce 3-sheet Excel: Clean_Data, Removed_Rows, Summary_Report |
