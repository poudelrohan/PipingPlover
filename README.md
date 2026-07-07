# USFWS Piping Plover Database Cleaning Pipeline

Data cleaning pipeline for Florida Piping Plover (PIPL) observation databases. Each source database is cleaned by its own pipeline and produces a per-database FINAL Excel file. A cross-database combiner then merges all four into one workbook for biologist review.

## Where Are the Final Clean Files?

| Database | Description | Final File(s) | Status |
|----------|-------------|---------------|--------|
| Database 1 | eBird observations | `Databases/Database1/Output/database1_FINAL.xlsx` | ✅ Complete (25,392 clean rows) |
| Database 2 | Non-breeding survey (2011–2023) | `Databases/Database2/Output/database2_FINAL.xlsx` | ✅ Complete (5,937 clean rows) |
| Database 3 | Winter Bird Survey (2013–2024) | `Databases/Database3/Output/<year>/db3_<year>_FINAL.xlsx` | ✅ Complete (all 12 years 2013–2024) |
| Database 3 — combined | All 12 years merged into one workbook | `Databases/Database3/Output/AllYears/db3_ALL_YEARS_FINAL.xlsx` | ✅ Complete (1,130 rows) |
| Database 4 | Banded bird resights | `Databases/Database4/Output/database4_FINAL.xlsx` | ✅ Complete (192 clean rows) |
| **All 4 databases combined** | **Harmonized cross-DB workbook** | `Databases/AllDatabasesCombined/db_ALL_COMBINED_FINAL.xlsx` | ✅ Complete (32,651 rows) |

Each per-database final file (`_FINAL.xlsx`) has 3 sheets:

- **Clean_Data** — All valid observations after cleaning
- **Removed_Rows** — Rows that were removed, with a `removal_reason` column explaining why
- **Summary_Report** — Statistics about the cleaning process (row counts, removal breakdown)

The cross-database combined file has 5 sheets: `AllDBCombined` (harmonized 35-column view) + one untouched copy of each source DB's main data sheet (`DB1`, `DB2`, `DB3`, `DB4`).

---

## Database 3 — Winter Bird Survey (Per-Year Pipeline)

Database 3 is the most complex. Each year has its own self-contained pipeline under `Code/Database3/<year>/`.

### Completed Years

| Year | Clean Rows | PIPL | Banded | Removed | Notes |
|------|-----------|------|--------|---------|-------|
| 2013 | 17 | 95 | 2 | 0 | Species GPS sheet; PIPL count parsed from free-text species column; biologist-corrected |
| 2014 | 100 | 371 | 46 | 3 | Indiv Flock GPS & bands sheet; band info in BandCombo (fixed from prior FlagCode misroute); biologist-corrected |
| 2015 | 128 | 374 | 67 | 0 | DS1 + DS3; Option A expansion |
| 2016 | 84 | 251 | 40 | 4 | DS2 audit; wrong-year DS1 rows removed |
| 2017 | 103 | 276 | 47 | 0 | All rows have GPS |
| 2018 | 118 | 306 | 73 | 0 | Wide format; biologist-corrected band data |
| 2019 | 109 | 314 | 44 | 1 | Wide format; pt1 lon from raw file; biologist-corrected |
| 2020 | 72 | 307 | 25 | 0 | Metadata joined from All Species sheet |
| 2021 | 106 | 254 | 39 | 0 | 1 longitude sign auto-corrected |
| 2022 | 100 | 278 | 35 | 0 | Biscayne route alias; degree-symbol GPS parse fix |
| 2023 | 85 | 354 | 37 | 1 | All Species column layout shifted (Transect renamed) |
| 2024 | 108 | 322 | 36 | 0 | 12 forced-unbanded directives applied |
| **Total** | **1,130** | **3,502** | **491** | **9** | 180 unique routes across 12 years |

### Format Changes by Year

| Years | Format | Sheet | Band Data |
|-------|--------|-------|-----------|
| 2013 | Species GPS (single sheet, multi-species per row) | Row 2 headers | Prose in Flag + Combo columns |
| 2014 | Indiv Flock GPS & bands (multi-species per row) | Row 3 headers | Prose in single band column |
| 2015–2017 | DS1 (flocks) + DS3 (per-bird resights) | Multiple sheets | Structured per-bird (UL/LL/UR/LR columns) |
| 2018–2019 | Wide format — one row per route, groups as columns | Single sheet | Free text in `{N}PIPLbands` |
| 2020–2024 | Wide format — Focal Observations sheet + All Species metadata | Two sheets | Free text in `PIPL Band/Flag Codes (point N)` |

### Merge Strategy (Option A Expansion)

For all years, the final output has **one row per individual bird**:

- Each banded bird gets its own row (`TotalObserved = 1`, `BandCombo` = band string)
- The remaining unbanded birds at that point get one row (`TotalObserved = PIPL − n_banded`)
- Points with no banded birds get one row (`TotalObserved = PIPL`, `BandCombo` = null)

### Band Data Processing

Raw band text is free-form. Processing happens in three stages:

**Stage 1 — Band Review** (`Code/Database3/parse_bands.py`)
Reads the PIPL-filtered clean files, splits each band cell into numbered individual bird entries (`1) ... \n2) ... \n3) ...`). Output: `Databases/Database3BandReview/`.

**Stage 2 — Biologist Review** (`Code/Database3/generate_biologist_review.py` and `generate_biologist_review_2013_2017.py`)
Generates Excel workbooks with one sheet per year. Biologists review all band entries, leave the Correction column blank if correct, or write the corrected `1) 2) 3)` format if not. Two separate workbooks exist:
- `Biologist_Band_Review_2013-2017.xlsx` (2013 and 2014 corrections applied; 2015–2017 available for future review)
- `Biologist_Band_Review_Completed_<year>.xlsx` for 2018–2024

Logic rules are documented in `Code/Database3/BAND_LOGIC.md`.

**Stage 3 — Pipeline** (`Code/Database3/<year>/0_extract_pipl.py`)
Reads biologist corrections from the completed review file. Priority: (1) biologist correction → (2) our interpretation → (3) raw text.

The biologist-correction resolver recognises these directives:
- `treat as unbanded` (in Notes or Correction) → force 0 banded, remainder unbanded
- `remove` / `delete` → skip that (route, point) entirely
- `good` / `confirm` / `correct` / `right` → confirmation, use our interpretation
- `unbanded` / `unreadable` → force 0 banded
- Any numbered `1) ... 2) ...` → use as split band list

---

## Running a Year's Pipeline

```bash
cd Code/Database3/2020        # or any year 2013–2024
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

## Combining All 12 Years of Database 3

After all per-year pipelines have run, one script merges every year into a single workbook:

```bash
cd Code/Database3
python3 combine_all_years.py
```

This populates `Databases/Database3/Output/AllYears/`:

- Copies of every `db3_<year>_FINAL.xlsx` (2013–2024)
- `db3_ALL_YEARS_FINAL.xlsx` — combined workbook with 14 sheets:
  - `All_Years_Combined` — 1,130 rows stacked in year order
  - `Summary` — one row per year of stats (routes, PIPL, banded, unbanded, % banded) + a TOTAL row
  - `2013`, `2014`, …, `2024` — per-year slices for easy click-through

The combined view adds four columns at the front:

| Column | Description |
|--------|-------------|
| `unique_id` | New global sequential int (1..N). Regenerated each run. |
| `database` | `3` (placeholder for future merges with DB1/2/4). |
| `year_id` | `<year>_<4-digit-original-id>` — stable citable key that survives pipeline re-runs. |
| `SurveyYear` | Int year, derived from `SurveyDate`. |

No rows are removed in this step — cross-year duplicates (same route/point/band combo in different years) are valid distinct annual observations.

---

## Combining All 4 Databases

The top-level combiner produces one workbook containing every clean row from all four databases, plus a harmonized view for cross-database analysis:

```bash
cd Code
python3 combine_all_databases.py
```

Output: `Databases/AllDatabasesCombined/db_ALL_COMBINED_FINAL.xlsx` (5 sheets):

| Sheet | Rows | Cols | Contents |
|---|---|---|---|
| `AllDBCombined` | 32,651 | 35 | All 4 DBs harmonized into a single column schema |
| `DB1` | 25,392 | 13 | Exact copy of `database1_FINAL.xlsx` → `Clean_Data` |
| `DB2` | 5,937 | 18 | Exact copy of `database2_FINAL.xlsx` → `Clean_Data` |
| `DB3` | 1,130 | 21 | Exact copy of `db3_ALL_YEARS_FINAL.xlsx` → `All_Years_Combined` |
| `DB4` | 192 | 23 | Exact copy of `database4_FINAL.xlsx` → `Clean_Data` |

The per-DB sheets are untouched copies of the source data (same columns, same values). The `AllDBCombined` sheet uses a unified 35-column schema with **column names that tell you which databases contribute to each field**:

- **No suffix** — all 4 DBs contribute (e.g. `Date`, `Location`, `Latitude`, `Longitude`, `Species`, `TotalObserved`, `PrimaryComments`)
- **`(DB2, DB3)`** — `GroupNumber`
- **`(DB2, DB3, DB4)`** — `TotalBanded`, `Observer`
- **`(DB3, DB4)`** — `FlagCode`
- **`(DB1, DB4)`** — `SecondaryComments`
- **`(DB3)`** — `ObserverEmail`, `FlagColor`, `BandCombo`, `WeatherCondition`
- **`(DB4)`** — `EndTime`, `FlagID`, `UpperLeft`, `LowerLeft`, `UpperRight`, `LowerRight`, `FlockActivity`
- **`(DB2)`** — `HabitatType`, `Tide`, `Foraging`, `Roosting`

The suffix is computed automatically from the data — if a column's contributing databases change in the future, the suffix updates on the next run.

Identifiers added to `AllDBCombined`:

| Column | Description |
|--------|-------------|
| `unique_id` | New global sequential int (1..32,651) |
| `database` | int 1/2/3/4 |
| `db_id` | Citable per-DB key: `db1_00042`, `db2_1234`, `db3_2020_0042`, `db4_017` |
| `Year` | Derived from Date (or DB3's `SurveyYear`) |

Species values are normalised to the 4-letter code (`PIPL`). The combiner warns in console if any non-PIPL species is found (currently 0 — all four DBs are 100% PIPL).

---

## Folder Structure

```
Databases/
  Database1/                     ← eBird PIPL observations
  Database2/                     ← Non-breeding PIPL survey data
  Database3/                     ← Raw yearly survey files (2013–2024)
    Output/
      2013/ … 2024/              ← Per-year pipeline outputs
        db3_<year>_FINAL.xlsx    ← FINAL CLEAN FILE (per year)
      AllYears/                  ← Copies of all 12 year FINAL files +
        db3_ALL_YEARS_FINAL.xlsx ← combined 14-sheet workbook
  Database3Clean/                ← PIPL-filtered versions of raw files
  Database3BandReview/           ← AI-structured band cells (Stage 1 output)
  Database3AIBandReview/         ← AI-reviewed band cells (work in progress)
  Database3BiologistReview/      ← Biologist review workbooks:
                                     Biologist_Band_Review_2013-2017.xlsx
                                     Biologist_Band_Review_Completed_<year>.xlsx (2018–2024)
  Database4/                     ← Banded bird resight data
  AllDatabasesCombined/          ← Cross-database final workbook
    db_ALL_COMBINED_FINAL.xlsx   ← 5-sheet harmonized combined view

Code/
  combine_all_databases.py       ← Cross-database combiner (top-level)
  Database1/                     ← DB1 pipeline scripts
  Database2/                     ← DB2 pipeline scripts
  Database3/
    BAND_LOGIC.md                ← Rules for interpreting band notation
    parse_bands.py               ← Band Review generator for 2019–2024
    parse_bands_2018.py          ← Band Review generator for 2018
    generate_biologist_review.py ← Biologist review generator (2018–2024)
    generate_biologist_review_2013_2017.py ← Biologist review generator (2013–2017)
    combine_all_years.py         ← Merge all 12 years into AllYears/
    2013/ … 2024/                ← Per-year self-contained pipelines
      database3_config.py        ← Year-specific config
      0_extract_pipl.py          ← Extract + biologist corrections + Option A expansion
      1_add_ids.py               ← Assign unique_id and source columns
      2_validate_geography.py    ← Florida bounding box check / coord fixes
      3_check_required_fields.py ← Remove rows missing required fields
      4_select_columns.py        ← Rename, reorder, keep standard columns
      5_remove_duplicates.py     ← Deduplicate observations
      6_create_final_report.py   ← Build 3-sheet FINAL.xlsx
  Database4/                     ← DB4 pipeline scripts
```

---

## Standard Output Columns (Database 3 per-year FINAL files)

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
| `FlagCode` | Flag code (populated in 2015–2017) |
| `FlagColor` | Flag color (populated in 2015–2017) |
| `BandCombo` | Full band combination string (null for unbanded rows) |
| `Comments` | Original observer comments only — no pipeline notes |
| `source_database` | Always `WinterBirdSurvey` |
| `source_file` | Source Excel filename |
| `source_sheet` | Source sheet name |

Not all years populate every column — earlier years (2013–2014) have fewer fields available in the source data. Where a column doesn't apply, the cell is left blank rather than fabricated.
