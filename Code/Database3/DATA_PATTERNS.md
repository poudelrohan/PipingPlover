# Database 3 — Data Patterns & Notes

## Overview
- **Source:** Winter Bird Survey data, 12 yearly Excel files (2013–2024)
- **Each file has 2 key sheets:**
  - **Species GPS** — individual flock/bird sightings with precise lat/lon
  - **Counts** — one row per survey route, with total PIPL count and route start/end coordinates

---

## Common Patterns (confirmed in 2013 + 2014)

### Two sheets, one-to-many relationship
- One Counts row (survey route) can have multiple Species GPS rows (individual flocks seen along that route)
- Example: Honeymoon Island north route → 1 Counts row (PIPL=14), 5 GPS rows (2+2+8+1+1=14)

### PIPL count extraction (Species GPS sheet)
The focal species column contains free text in many formats. Extraction rule (in priority order):

| Priority | Pattern | Example | Extracts |
|-|-|-|-|
| 1 | `N PIPL` (number before PIPL) | `9 PIPL (2 banded)` | 9 |
| 2 | `PIPL (N)` (simple number in parens) | `PIPL (5)` | 5 |
| 3 | `PIPL (N total/comma...)` (number then extra text) | `PIPL (5 total, 1 banded)` | 5 |
| 4 | `PIPL - N` or `PIPL-N` (dash separated) | `PIPL- 8, SNPL- 4` | 8 |
| 5 | `PIPL N` (space separated, no parens) | `PIPL 17, SPPL 67` | 17 |
| 6 | `Piping Plovers (N: ...)` (full name) | `Piping Plovers (12: 10 unbanded, 2 banded)` | 12 |
| 7 | Bare `PIPL` or `Piping Plover` (no number) | `PIPL` | Fallback to `No. of individuals` column (2013), or assume 1 if no such column |

**Priority order matters!** `9 PIPL (2 banded)` must match rule 1 (=9), not rule 3 (=2).

Using these rules, all matched GPS↔Counts rows produce **perfect count matches** in both 2013 and 2014.

### Location matching
- Location names vary slightly between sheets (dashes, spaces, typos)
- **Strict matching works best:** normalize punctuation but keep route numbers/directions
  - "Cayo Costa Route 2" must NOT match "Cayo Costa Route 3"
  - "Three Rooker Island - north" must NOT match "Three Rooker Island - south"
- Loose prefix matching (first 12-15 chars) causes false cross-matches on numbered/directional routes
- **Solution:** Normalize commas/spaces, then check: exact match OR one name fully contains the other

### Coverage gaps (consistent pattern)
- Not all survey routes have GPS data — some routes only appear in Counts
- 2013: 7 of 17 routes have GPS (41%)
- 2014: 24 of 35 routes have GPS (69%)
- Routes without GPS still have: route name, county, date, observers, starting/ending lat/lon

### Unmatched rows (both directions)
**GPS rows with no Counts match:**
- Typos: "North Anclote Bar" (GPS) vs "Noth Anclote Bar" (Counts) — fixable
- Off-cycle observations: GPS sighting on a date with no survey count logged
- Bad dates: 2014 has a GPS row dated `2001-02-16` (likely typo for 2014)
- Missing dates: Some GPS rows have no date at all (NaT)

**Counts rows with no GPS match:**
- Valid PIPL observations, just nobody recorded individual flock GPS points

### Data quality issues
- **Coordinate sign errors:** Some GPS rows have positive longitude (e.g. 82.73 instead of -82.73). Same pattern we auto-correct in other databases.
- **Date errors:** Wrong year (2001 instead of 2014), missing dates
- **Typos in location names:** "Noth" vs "North"

---

## 2013 Specifics

- **Sheet names:** "Species GPS", "counts"
- **GPS header row:** 2 (row 0 = biologist annotations, row 1 = instructions)
- **Species column:** "Focal species (PIPL, SNPL, REKN, WIPL, other banded birds)"
- **Has `No. of individuals` column:** YES — used as fallback when text just says "PIPL"
- **Location column (GPS):** "Location (all lines of data need a location)"
- **Location column (Counts):** "Location " (with trailing space)
- **Results:** 17 counts rows, 15 GPS rows. 7 perfect matches, 10 no GPS, 2 unmatched GPS

### Column inventory (2013)

**Species GPS sheet:**
- Date, Location, Focal species (free text), No. of individuals, List species in flock
- **(Individual bird or flock) Latitude**, **(Individual bird or flock) Longitude** ← the coordinates we want
- Observers, Email, Flag color and code, Color band combo, Comments

**Counts sheet:**
- Location, County, Starting Lat/Lon, Ending Lat/Lon ← route boundaries, NOT flock locations
- Observers, Email, Date, Starting Time
- Piping Plover (count), GRAND TOTAL, Comments

---

## 2014 Specifics

- **Sheet names:** "Indiv Flock GPS & bands", "Total Survey Counts", "Sheet1" (empty)
- **GPS header row:** 2 (row 0 = biologist annotations, row 1 = instructions)
- **Species column:** "Species and number of individuals"
- **Has `No. of individuals` column:** NO — bare "PIPL" defaults to 1
- **Location column (GPS):** "Route Name"
- **Location column (Counts):** "Route Name"
- **Results:** 35 counts rows, 67 GPS rows. 24 perfect matches, 9 no GPS, 2 genuine mismatches

### Biologist annotations (row 0 of GPS sheet, 2014)
- Lat column: "match this with total survey count"
- Lon column: "match it with total survey count"
- Species column: "extract the count of pipl total it and match it with total survey counts"
- Band info: "yes keep it"
- Photo available: "No"

### New text patterns in 2014 (not seen in 2013)
- `PIPL (5 total, 1 banded)` — number with "total" qualifier
- `PIPL (14, 2 banded)` — number then comma with banding info
- `PIPL 17, SPPL 67` — space-separated (no dash or parens)
- `Piping Plovers (12: 10 unbanded, 2 banded)` — full species name
- `9 PIPL (2 banded)` — number before PIPL with extra parens info
- `PIPL, AMOY(2)` — PIPL with no number at all, next to another species

### Genuine data discrepancies (2014, need biologist review)
1. **AMI (Alamanda Rd to Longboat):** Counts=2, GPS=92. GPS row has `PIPL (92)` in a massive mixed flock — likely a data entry error (92 may be total shorebirds, not PIPL).
2. **SEVAS VII (Canaveral NS):** Counts=14, GPS has only 1 row with PIPL=5. 9 birds not GPS-logged.

### Unmatched GPS rows (2014)
1. "dania to hollywood beach" — no date, no count extracted
2. "Hobe Sound NWR" dated 2001-02-16 — likely date typo
3. Two "SEVAS VII" rows with no date (NaT)

---

## Questions for Biologist (pending)
1. Routes that only appear in Counts (no GPS data) — keep with route Starting Lat/Lon as fallback, or leave lat/lon blank, or exclude?
2. GPS sightings with no matching survey count (off-cycle observations) — keep or exclude?
3. Which columns to keep in final output? Are banding columns (flag color, band combo) needed?
4. Genuine data discrepancies (AMI 2 vs 92, SEVAS 14 vs 5) — which sheet to trust?
5. GPS rows with bad/missing dates — fixable or exclude?

---

## Other Years
*(To be filled in as we explore each year)*

### 2015
- Sheet names: "DATA SHEET 1", "DATA SHEET 2", "DATA SHEET 3" (extra sheet)
- Species column: "Species and number"
- TODO: explore and document patterns

### 2016
- Sheet names: "DATA SHEET 1", "DATA SHEET 2", "DATA SHEET 3" (extra sheet)
- Species column: "Species and number"
- TODO: explore and document patterns

### 2017–2024
- Not yet explored — sheet names/structure unknown
- TODO: explore and document patterns
