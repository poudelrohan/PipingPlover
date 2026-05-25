# USFWS Piping Plover Database Cleaning Pipeline

Data cleaning pipeline for Florida Piping Plover (PIPL) observation databases. Each database goes through a structured cleaning process that produces a final Excel report with clean data, removed rows, and summary statistics.

## Where Are the Final Clean Files?

| Database | Description | Final File(s) | Status |
|----------|-------------|---------------|--------|
| Database 1 | eBird observations | `Databases/Database1/Output/database1_FINAL.xlsx` | ✅ Complete (25,392 clean rows) |
| Database 2 | Non-breeding survey (2011–2023) | `Databases/Database2/Output/database2_FINAL.xlsx` | ✅ Complete (5,937 clean rows) |
| Database 3 | Winter Bird Survey (2013–2024) | `Databases/Database3/Output/<year>/db3_<year>_FINAL.xlsx` | 🔄 In progress (2013–2019 done) |
| Database 4 | Banded bird resights | `Databases/Database4/Output/database4_FINAL.xlsx` | ✅ Complete (192 clean rows) |

Each final file (`_FINAL.xlsx`) has 3 sheets:

- **Clean_Data** — All valid observations after cleaning
- **Removed_Rows** — Rows that were removed, with a `removal_reason` column explaining why
- **Summary_Report** — Statistics about the cleaning process (row counts, removal breakdown)

---

## Database 3 — Winter Bird Survey (Per-Year Pipeline)

Database 3 is the most complex. Each year has its own self-contained pipeline under `Code/Database3/<year>/`.

### Completed Years

| Year | Clean Rows | PIPL | Banded Rows | Removed | Notes |
|------|-----------|------|-------------|---------|-------|
| 2013 | 15 | — | — | 0 | Sheet: "Species GPS", header_row=1 |
| 2014 | 65 | — | — | 2 | 2 rows removed for missing date |
| 2015 | 128 | — | — | 0 | DS1 + DS3, Option A expansion |
| 2016 | 84 | 251 | — | 4 | DS2 audit; wrong-year DS1 rows removed |
| 2017 | 103 | 276 | — | 0 | All rows have GPS |
| 2018 | 118 | 306 | 74 | 0 | Wide format; biologist-corrected band data |
| 2019 | 109 | 314 | 44 | 1 | Wide format; pt1 lon from raw file; biologist-corrected |
| 2020–2024 | — | — | — | — | Pipelines pending biologist band review |

### Format Changes by Year

| Years | Format | Sheet | Band Data |
|-------|--------|-------|-----------|
| 2013–2017 | DS1 (flocks) + DS3 (resights) | Multiple sheets | Structured per-bird |
| 2018–2019 | Wide format — one row per route, groups as columns | Single sheet | Free text in `{N}PIPLbands` |
| 2020–2024 | Wide format — Focal Observations sheet | Two sheets | Free text in `PIPL Band/Flag Codes (point N)` |

### Merge Strategy (Option A Expansion)

For all years, the final output has **one row per individual bird**:

- Each banded bird gets its own row (`TotalObserved = 1`, `BandCombo` = band string)
- The remaining unbanded birds at that point get one row (`TotalObserved = PIPL − n_banded`)
- Points with no banded birds get one row (`TotalObserved = PIPL`, `BandCombo` = null)

### Band Data Processing (2018+)

Raw band text in the survey data is free-form. Processing happens in three stages:

**Stage 1 — Band Review** (`Code/Database3/parse_bands_2018.py`, `parse_bands.py`)
Reads the PIPL-filtered clean files, splits each band cell into numbered individual bird entries (`1) ... \n2) ... \n3) ...`). Cells that cannot be confidently split are highlighted pink for biologist review. Output: `Databases/Database3BandReview/`.

**Stage 2 — Biologist Review** (`Code/Database3/generate_biologist_review.py`)
Generates a single Excel file (`Databases/Database3BiologistReview/`) with one sheet per year. Biologists review all band entries, leave the Correction column blank if correct, or write the corrected `1) 2) 3)` format if not. Logic rules are documented in `Code/Database3/BAND_LOGIC.md`.

**Stage 3 — Pipeline** (`Code/Database3/<year>/0_extract_pipl.py`)
Reads biologist corrections from the completed review file. Priority: (1) biologist correction → (2) our interpretation → (3) raw text. Corrections saying "unreadable / make unbanded" or "remove this row" are handled automatically.

---

## Running a Year's Pipeline

```bash
cd Code/Database3/2018        # or 2019, 2017, etc.
python3 0_extract_pipl.py
python3 1_add_ids.py
python3 2_validate_geography.py
python3 3_check_required_fields.py
python3 4_select_columns.py
python3 5_remove_duplicates.py
python3 6_create_final_report.py
```

Output lands in `Databases/Database3/Output/<year>/db3_<year>_FINAL.xlsx`.

---

## Folder Structure

```
Databases/
  Database1/                     ← eBird PIPL observations
  Database2/                     ← Non-breeding PIPL survey data
  Database3/                     ← Raw yearly survey files (2013–2024)
    Output/
      2013/ … 2019/              ← Per-year pipeline outputs
        db3_<year>_FINAL.xlsx    ← FINAL CLEAN FILE
  Database3Clean/                ← PIPL-filtered versions of raw files
  Database3BandReview/           ← AI-structured band cells (Stage 1 output)
  Database3AIBandReview/         ← AI-reviewed band cells (work in progress)
  Database3BiologistReview/      ← Biologist review workbooks (Stage 2)
  Database4/                     ← Banded bird resight data

Code/
  Database1/                     ← Pipeline scripts
  Database2/
  Database3/
    BAND_LOGIC.md                ← Rules for interpreting band notation
    parse_bands_2018.py          ← Band Review generator for 2018
    parse_bands.py               ← Band Review generator for 2019–2024
    generate_biologist_review.py ← Biologist review Excel generator
    2013/ … 2019/                ← Per-year self-contained pipelines
      database3_config.py        ← Year-specific config
      0_extract_pipl.py          ← Extract + merge + Option A expansion
      1_add_ids.py               ← Assign unique_id and source columns
      2_validate_geography.py    ← Florida bounding box check / coord fixes
      3_check_required_fields.py ← Remove rows missing required fields
      4_select_columns.py        ← Rename, reorder, keep standard columns
      5_remove_duplicates.py     ← Deduplicate observations
      6_create_final_report.py   ← Build 3-sheet FINAL.xlsx
  Database4/
```

## Standard Output Columns

All Database 3 final outputs share these columns:

| Column | Description |
|--------|-------------|
| `unique_id` | Sequential integer per year |
| `SurveyDate` | Date of survey |
| `SurveyTime` | Route start time |
| `WeatherCondition` | Temperature, wind, rain combined |
| `Route` | Survey route name |
| `Latitude` | Point/group latitude |
| `Longitude` | Point/group longitude |
| `GroupNumber` | Point or group number (1–19) |
| `TotalObserved` | Birds in this row (1 for banded, N for unbanded remainder) |
| `Observer` | Observer name(s) |
| `ObserverEmail` | Lead observer email (where available) |
| `FlagCode` | Flag code (2013–2017 only) |
| `FlagColor` | Flag color (2013–2017 only) |
| `BandCombo` | Full band combination string (null for unbanded rows) |
| `Comments` | Original observer comments only — no pipeline notes |
| `source_database` | Always `WinterBirdSurvey` |
| `source_file` | Source Excel filename |
| `source_sheet` | Source sheet name |
