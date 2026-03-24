"""Temporary exploration script for Database 3 - 2013 file"""
import pandas as pd
import re

# Load both sheets
counts = pd.read_excel("Databases/Database3/Winter Birds '13.xlsx", sheet_name='counts', header=1)
gps = pd.read_excel("Databases/Database3/Winter Birds '13.xlsx", sheet_name='Species GPS', header=1)

# Clean up counts - remove blank/instruction rows
loc_col = 'Location '
mask = counts[loc_col].notna() & ~counts[loc_col].astype(str).isin(['Leave blank'])
counts = counts[mask]

# Filter counts to PIPL rows
pipl_counts = counts[counts['Piping Plover'].notna() & (counts['Piping Plover'] > 0)].copy()
print(f'=== Counts sheet: {len(pipl_counts)} rows with PIPL > 0 ===')

# Filter Species GPS to PIPL rows
focal_col = 'Focal species (PIPL, SNPL, REKN, WIPL, other banded birds)'
pipl_gps = gps[gps[focal_col].astype(str).str.contains('PIPL', case=False, na=False)].copy()
print(f'=== Species GPS: {len(pipl_gps)} rows mentioning PIPL ===')

# Extract PIPL number from focal species text
def extract_pipl_count(text):
    if pd.isna(text):
        return None
    text = str(text)
    # Pattern 1: 'PIPL- 8' or 'PIPL -8' or 'PIPL-8'
    m = re.search(r'PIPL\s*[-:]\s*(\d+)', text, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # Pattern 2: '46 PIPL'
    m = re.search(r'(\d+)\s*PIPL', text, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # Pattern 3: Just 'PIPL' alone
    if re.search(r'\bPIPL\b', text, re.IGNORECASE):
        return None
    return None

pipl_gps['pipl_count_from_focal'] = pipl_gps[focal_col].apply(extract_pipl_count)

print()
print('=== Comparison: Counts PIPL rows vs Species GPS PIPL rows ===')
print()
for _, cr in pipl_counts.iterrows():
    date = cr['Date ']
    loc = str(cr['Location ']).strip()
    count = int(cr['Piping Plover'])
    observer = str(cr['Observers']).strip()[:35] if pd.notna(cr['Observers']) else 'N/A'

    # Find matching GPS rows by date
    matches = pipl_gps[pipl_gps['Date '] == date]

    print(f'COUNTS: {loc[:45]:45s} | {str(date)[:10]} | PIPL={count:3d} | {observer}')

    if len(matches) > 0:
        for _, gm in matches.iterrows():
            gloc = str(gm['Location (all lines of data need a location)']).strip()[:45]
            gfocal = str(gm[focal_col])[:55]
            gcount = gm['pipl_count_from_focal']
            glat = gm['(Individual bird or flock) Latitude']
            glon = gm['(Individual bird or flock) Longitude']
            gobs = str(gm['Observers']).strip()[:35] if pd.notna(gm['Observers']) else 'N/A'
            match_loc = 'LOC-MATCH' if loc.lower().startswith(gloc[:15].lower()) or gloc.lower().startswith(loc[:15].lower()) else ''
            match_count = f'COUNT-MATCH' if gcount is not None and gcount == count else ''
            print(f'  GPS:  {gloc:45s} | {gfocal:55s} | PIPL#={gcount} | ({glat}, {glon}) | {gobs} | {match_loc} {match_count}')
    else:
        print(f'  GPS:  ** NO MATCHING GPS ROWS ON THIS DATE **')
    print()

# Also check: are there PIPL GPS rows with no matching counts row?
print('=== GPS PIPL rows with NO matching counts date ===')
counts_dates = set(pipl_counts['Date '].dropna())
for _, gm in pipl_gps.iterrows():
    gdate = gm['Date ']
    if gdate not in counts_dates:
        gloc = str(gm['Location (all lines of data need a location)']).strip()[:45]
        print(f'  GPS orphan: {gloc} | Date: {gdate} | Focal: {gm[focal_col]}')
