# Database 3 — Data Issues Log

Problems found during the "light clean" stage (row/column filtering).
These are NOT fixed in the clean files — they are flagged here so the
7-step pipeline can handle them year by year.

---

## 2016

| Route | Issue | Type |
|-------|-------|------|
| Long Key State Park | 2 stray rows dated 2017-02-08 and 2017-02-07 found in DS1 | Wrong year in file |
| Mizell-Johnson State Park | Row dated 2017-02-07 in DS1 | Wrong year in file |
| Outback Key | DS1 entry has no Group # and coordinates written as "same" (unresolvable) | Missing GPS |
| Outback Key | 2 DS3 band rows could not be matched to DS1 — removed | DS3 mismatch |
| Three Rooker Bar, north island | DS3 recorded groups 2 & 3 but DS1 only has groups 4 & 5; DS2 confirms total = 10 | DS3 group # recording error |
| S. Hutchinson Island | DS3 zone written as "YM-JJ" instead of "T-Y" — confirmed typo | Route name typo |
| Bunche Beach Preserve | DS3 row has blank Group # | Missing group number |
| Lanark Reef / Bunche Beach | DS3 date entered as 2015-02-08 instead of 2016-02-08 | Date typo |
| North Rockhouse Creek Shoals | DS1 sums to 13 PIPL but DS2 says 12 — raw data discrepancy | Count mismatch (unfixable) |
| Fort De Soto North Beach Lagoon | DS2 shows 1 PIPL but no DS1/DS3 rows submitted | Missing detail data |

---

## 2017

| Route | Issue | Type |
|-------|-------|------|
| Lanark Reef | Date entered as `02/08/2017\`` (backtick in cell) | Date parse error |
| Outback Key (group 2) | Longitude entered as `-82751911` — missing decimal, should be `-82.751911` | Coordinate typo |
| Highland Beach (group 2) | Longitude positive (`81.20972`) — should be negative | Coordinate sign error |
| New Smyrna Beach (group 6) | Longitude positive (`80.911805`) — should be negative | Coordinate sign error |
| Wild Goose Lagoon | DS3 route name has trailing newline (`"Wild Goose Lagoon\n"`) | Route name formatting |
| Dunlawton Bridge | DS3 spells it "Dunlawton" but DS1 spells it "Dunlaton" | Route name spelling mismatch |
| Borth Biscayne Bay Islands | Coordinates place it in Pinellas County (28.085, -82.837), not Biscayne Bay — likely route name typo | Route name typo |
| DS2 | One 2016-dated row (St. Andrews State Park, 2016-02-04) in the 2017 file | Wrong year in file |
| DS2 | Several routes appear twice with identical counts (Anclote Key, Three Rooker, etc.) | Duplicate DS2 submissions |
| Martin County IRL | DS2 PIPL column has `PIPLbands` as free-text band description instead of a count | Wrong data type in column |
| Pavilion Key | In DS3 but not in DS1 — unmatched band resight rows | DS3 mismatch |
| Navarre Beach | DS3 spells route "Navarre Beach Soundside" (one word) but DS1/DS2 spell it "Navarre Beach Sound Side" (two words) — spacing mismatch prevents matching | Route name formatting |
| Bunche Beach | DS3 records group 1 but DS1 only has group 2 (15 PIPL) — group number recording error; DS3 group reassigned to 2 for matching | DS3 group # recording error |

---

## 2018

**Format change:** Single sheet (Sheet1), one row per transect (wide format).
All groups/points are columns on the same row instead of separate rows.
No DS3 equivalent — band info is a count per group, not individual bird detail.
Pipeline will need to melt wide → long format before processing.

| Route | Issue | Type |
|-------|-------|------|
| Hobe Sound NWR | Route total = 2 PIPL but no group-level GPS points entered | Missing group detail |
| St. Joseph Peninsula State Park | Route total = 2 PIPL but all group columns show 0 | Missing group detail |
| Martin County IRL | Route total = 3, group sum = 2 — off by 1; also band info entered as free text in band count column | Count mismatch + wrong data type |
| Shell Key Preserve | Route total = 15 PIPL but group sum = 17 — route total is wrong (undercounted) | Count mismatch |
| Three Rooker South Island | Route total = 37, group sum = 29 — 8 PIPL seen but not assigned to any group | Missing group detail |
| General (multiple routes) | `PIPLbands` column used as free-text band combo description instead of integer count | Wrong data type in column |
| Group 16 | Column block exists but has NO PIPL column (only AMOY/REKN/SNPL/WIPL/BLSK) | Missing PIPL column for group |
| Group 18 | Entirely missing from the column structure | Missing group |
| Group 20 | Only has lat/long/nonfocals — no PIPL column | Missing PIPL column for group |

---

## 2019

**Format:** Google Form export (sheet: "Form Responses 1"), same wide structure as 2018.
19 groups (1-19, all present). Band info is free text per group (richer than 2018).
Route names have ", County=X" appended from Google Form dropdown — needs stripping in pipeline.

| Route | Issue | Type |
|-------|-------|------|
| Cedar Key West | Route total = 3 PIPL but no group-level GPS points entered | Missing group detail |
| Hillsboro Inlet to Lauderdale by the Sea | Group sum = 10 PIPL but route total left blank (0) | Missing route total |
| New Smyrna Beach jetty at Ponce Inlet | Two identical submissions — same date (2019-02-01), same count (1 PIPL) | Duplicate submission |
| General | Route names have ", County=X" suffix from Google Form dropdown | Route name formatting |

---

## 2020

**Format:** Two useful sheets — "All Species" (route totals, like DS2) and "Focal Observations"
(per-group detail, like DS1). Both use the same transect names so they can be joined.
19 groups per transect. PIPL band info is free text per group.

| Route | Issue | Type |
|-------|-------|------|
| Anastasia State Park Beach | Route total = 1 PIPL but no focal group GPS detail entered | Missing group detail |
| Ponce Inlet Parks shorelines and Disappearing Islands | Route total = 3 but focal group sum = 5 (+2) | Count mismatch |
| "Z My transect is not listed" | Catch-all Google Form entry; route total = 1 PIPL but no GPS detail | Missing group detail + unnamed route |

---

## 2021

**Format:** Two sheets — "All Species" (route totals) + "Focal Observations" (per-group GPS, MOST IMPORTANT).
19 groups per transect. Same structure as 2020.

| Route | Issue | Type |
|-------|-------|------|
| All Species | "Totals" row present with PIPL=312 — summary row, not a transect; dropped in cleaning | Summary row |
| Little Talbot Island State Park | AllSp=13, Focal sum=15 (+2) — focal records more PIPL than route total | Count mismatch |
| Flag Island | AllSp=5, Focal sum=6 (+1) — focal records more PIPL than route total | Count mismatch |
| Crandon Park Beach | AllSp=26, Focal=0 — route total filled but no GPS group detail submitted | Missing group detail |
| Palm Beach (Monceau Court) to Breakers Resort | AllSp=7, Focal sum=2 — two focal rows (one empty, one with 2 PIPL); route total does not match focal | Count mismatch |
| Outback Key | AllSp=31, Focal=0 — route total filled but no GPS group detail submitted | Missing group detail |

---

## 2022

**Format:** Two sheets — "All Species" (route totals) + "Focal Observations" (per-group GPS, MOST IMPORTANT).
19 groups per transect. Same structure as 2020–2021.

| Route | Issue | Type |
|-------|-------|------|
| All Species | "Totals" row present with PIPL=288 — summary row, not a transect; dropped in cleaning | Summary row |
| Dr. Von D Mizell-Eula Johnson State Park | AllSp=3, Focal=0 — route total filled but no GPS group detail submitted | Missing group detail |
| Huguenot Memorial Park | AllSp=1, Focal=0 — route total filled but no GPS group detail submitted | Missing group detail |
| Biscayne National Park / Convoy Point to Elliott Key Harbor... | AllSp=1, Focal=0 — no GPS group detail | Missing group detail |
| Long Key State Park | AllSp=1, Focal=0 — route total filled but no GPS group detail submitted | Missing group detail |
| Fred Howard Park | AllSp=4, Focal=0 — route total filled but no GPS group detail submitted | Missing group detail |

---

## 2023

**Format:** Two sheets — "All Species" (route totals) + "Focal Observations" (per-group GPS, MOST IMPORTANT).
19 groups per transect. Key change from 2020–2022: All Species col 0 is now the long transect name
"Transect (if yours is missing, email Beth Forys: forysea@eckerd.edu)"; 11 metadata cols instead of 10
(extra col: "If you did not survey - let us know why"). Focal Observations still uses "Transect " (short).

| Route | Issue | Type |
|-------|-------|------|
| All Species | "Totals" row present — summary row, not a transect; dropped in cleaning | Summary row |
| Bunche Beach - main? | AllSp=1, Focal=0 — route total filled but no GPS group detail submitted | Missing group detail |
| Lake Worth Lagoon - Ibis Island to Peanut Island | AllSp=2, Focal=0 — no GPS group detail | Missing group detail |

---

## 2024

**Format:** Two sheets — "All Species" (route totals) + "Focal Observations" (per-group GPS, MOST IMPORTANT).
19 groups per transect. Same structure as 2023 (long transect col name, 11 metadata cols).

| Route | Issue | Type |
|-------|-------|------|
| All Species | "Totals" row present — summary row, not a transect; dropped in cleaning | Summary row |
| Alligator Point Developed | AllSp=1, Focal=0 — route total filled but no GPS group detail submitted | Missing group detail |
| Cedar Key - West | AllSp=1, Focal=0 — no GPS group detail | Missing group detail |
| Grassy Key | AllSp=1, Focal=0 — no GPS group detail | Missing group detail |
| Ohio Key | AllSp=1, Focal=0 — no GPS group detail | Missing group detail |
| Atlantic Dunes Park to Palm Beach (Monceau Court) | AllSp=6, Focal=0 — no GPS group detail | Missing group detail |
| John D. MacArthur Beach State Park | AllSp=4, Focal=0 — no GPS group detail | Missing group detail |
| Fort Matanzas National Monument | AllSp=1, Focal=0 — no GPS group detail | Missing group detail |
