# Database 3 (Winter Bird Surveys) — Summary & Questions for Biologist

## What We Have

- 12 yearly Excel files (2013-2024)
- Each file has 2 sheets with PIPL data:
  - **Species GPS sheet** — individual flock sightings with precise lat/lon coordinates
  - **Counts sheet** — one row per survey route, with total PIPL count and route-level start/end coordinates (not flock-level)
- The two sheets are related: each survey route (1 Counts row) can have multiple individual flock GPS points recorded along it

## What We've Found (explored 2013 and 2014 so far)

### 1. The two sheets are complementary but incomplete in different ways

The Species GPS sheet has the precise coordinates we want, but not every survey route has GPS data recorded.
The Counts sheet covers every route surveyed but only has route start/end coordinates, not where the actual birds were.

| Year | Total routes surveyed (Counts) | Routes with GPS flock data | Coverage |
|------|-------------------------------|---------------------------|----------|
| 2013 | 17 | 7 | 41% |
| 2014 | 35 | 24 | 69% |

**Examples of routes with NO GPS data (2013):**
- Shell Key (23 PIPL counted, no flock GPS recorded)
- Three Rooker Bar Mid-North Route (75 PIPL counted, no flock GPS)
- Huguenot Memorial Park (6 PIPL counted, no flock GPS)
- Caladesi Island State Park (9 PIPL, no flock GPS)

### 2. Where GPS and Counts both exist, the PIPL numbers match

We can extract the PIPL count from the free-text species column in the GPS sheet and match it to the Counts total. When both sheets have data for the same route and date, the counts add up correctly.

**Example (2013, Honeymoon Island north route):**
- Counts sheet: 14 PIPL total
- GPS sheet: 5 flock entries → 2 + 2 + 8 + 1 + 1 = 14. Perfect match.

### 3. Some GPS rows don't match any Counts row

**Typos in location names:**
- GPS says "North Anclote Bar" but Counts says "Noth Anclote Bar" (2013) — we can fix these ourselves
- GPS says "Crooked Island West- Tyndall AFB" but Counts says "Crooked Island West- Tyndall Air Force Base" (2014)
- GPS says "dania to hollywood beach" but Counts says "Hollywood to Dania Beach" (2014)
- GPS says "South Beach Park, Vero Beach, Indian River County" but Counts says "South Beach Park, Vero Beach" (2014)

**Off-cycle observations:**
- 2013: GPS sighting at Mid-Fort De Soto on April 17, but the survey for that route was logged on Feb 6. Someone recorded a PIPL with GPS outside of the survey window.

**Date errors:**
- 2014: GPS row for "Hobe Sound NWR" is dated 2001-02-16 — likely should be 2014
- 2014: Some GPS rows have no date at all

### 4. Some Counts rows have no matching GPS data at all

These routes were surveyed and PIPL were counted, but nobody recorded individual flock GPS points.

**Examples from 2014 (8 unmatched routes):**
- Shell Island- Tyndall Air Force Base: 11 PIPL
- Ponce Inlet North Rockhouse Creek Shoals: 5 PIPL
- Cedar Key (west side): 3 PIPL
- Fort Desoto East: 2 PIPL
- North end of Little Estero CWA to Carlos Point Condos: 2 PIPL (also missing date)

### 5. A few genuine data discrepancies where both sheets have data but counts don't match

These are cases where the GPS and Counts sheets both have data for the same route and date, but the PIPL numbers don't agree:

| Route | Year | Counts total | GPS sum | Difference | Possible explanation |
|-------|------|-------------|---------|------------|---------------------|
| AMI - Alamanda Rd to Longboat Pass | 2014 | 2 | 92 | -90 | GPS row says "PIPL (92)" in a mixed flock of 100+ — likely 92 is total shorebirds, not PIPL |
| Anclote Key | 2014 | 64 | 4 | +60 | Only 4 of 64 birds had GPS flock points recorded |
| Gulf Islands NS Santa Rosa North | 2014 | 15 | 2 | +13 | Only 2 of 15 birds had GPS points |
| SEVAS VII Canaveral NS | 2014 | 14 | 5 | +9 | Only 1 GPS row out of 14 birds |

---

## Questions

### Q1: What is the primary data source?

Since the GPS sheet is incomplete (not all routes have flock GPS), how should we handle routes that only exist in the Counts sheet?

Options:
- **(a)** Only keep rows that have precise flock GPS coordinates (we lose the routes with no GPS data)
- **(b)** Keep all rows — use precise GPS where available, use route Starting Lat/Lon as a fallback coordinate for routes without GPS
- **(c)** Keep all rows — use precise GPS where available, leave lat/lon blank for routes without GPS

### Q2: GPS sightings with no matching survey count — keep or drop?

Example: 2013 Mid-Fort De Soto on April 17 — someone recorded a PIPL with GPS but no survey was logged that day. Should we keep these observations or exclude them?

### Q3: Data discrepancies — which sheet do we trust?

When the GPS count and the Counts total disagree (like AMI: Counts=2 vs GPS=92), which number should we use? Should we flag these for manual review?

### Q4: What columns should we keep in the final output?

The GPS sheet has banding information:
- Flag color and code
- Color band combo
- Band information (upper left, lower left, upper right, lower right)

Are these needed in the final clean dataset?

### Q5: GPS rows with bad or missing dates — fix or exclude?

- 2014 "Hobe Sound NWR" dated 2001-02-16 (likely 2014) — should we correct this?
- Some GPS rows have no date at all — drop them or keep with blank date?

### Q6: Column annotations

Same as we did for Database 1 — can you mark up a copy of one of the files with which columns to keep and which to drop?

### Q7: Is the structure consistent across all 12 years?

We've only explored 2013 and 2014 so far. The sheet names and column names already differ between these two years. Should we expect more variation in later years? Any years that were set up differently?
