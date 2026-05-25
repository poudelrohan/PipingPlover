# Band Data Logic Rules
Reference used when interpreting raw band text cells from Winter Bird Survey data.

---

## 1. Count Constraint (hard rule)
- `banded_entries ≤ PIPL_count` at that point — always.
- If `banded_entries > PIPL_count` → **flag immediately**, do not write to output.
- If `PIPL_count = 0` → any band text is noise, **skip entirely**.
- `remainder = PIPL_count − banded_entries` → pipeline creates 1 unbanded row automatically.

---

## 2. "No banded birds" synonyms (skip — not shown to biologist)
Treat any of these as "no banded birds" and skip:
`none`, `n/a`, `na`, `0`, `no bands`, `no bands seen`, `none banded`,
`not banded`, `both were unbanded`, `all unbanded`, `all xx`, `xx`, `x:x`,
`no bands or flags`, `none visible`, `n.a.`, blank/null

---

## 3. Unbanded markers within a cell
- `X:X`, `X//X`, `XX` within a multi-bird cell = that specific bird is **unbanded**.
- Do **not** create a banded row for it — it contributes to the unbanded remainder count instead.
- Example: `1) Yf(A41)//WK:S//BO  2) X:X` with PIPL=2 → 1 banded row + 1 unbanded row.

---

## 4. "and N unbanded" / "N unbanded" suffix
Phrases like `"and 2 unbanded"` or `"4 Unbanded, 1 Banded"` are metadata.
Strip from BandCombo — the pipeline calculates the remainder automatically.

---

## 5. Notation formats (all describe 1 bird each)

### Slash notation
`UL//LL:UR//LR` — `//` separates upper/lower on same leg, `:` separates left/right legs.
Example: `Of//GG:S//K` = UL=Orange flag, LL=Green/Green, UR=Silver, LR=Black → **1 bird**

### Dot-dash notation
`UL.LL-UR.LR` — `.` separates upper/lower, `-` separates left/right legs.
Example: `FO.YY-S.G` = UL=Orange Flag, LL=Yellow, UR=Silver, LR=Green → **1 bird**

### Comma notation
`UL,LL:UR,LR` — comma separates upper/lower, colon separates legs.
Example: `S,_:GF(E00),_` = UL=Silver, LL=none, UR=Green flag (E00), LR=none → **1 bird**

### Prose notation
`"Left leg: Orange flag (YX6), upper; Silver, lower. Right leg: ..."` → **1 bird**

---

## 6. Code identifiers (part of one bird's combo — not separators)
Parenthetical codes like `(KT)`, `(T14)`, `(82A)`, `[MA]`, `[76E]`, `(E00)` are
flag/band identifiers. They are part of one bird's entry, never a separator between birds.

---

## 7. Multi-bird separators (split on these)
In order of reliability:
1. Newlines (`\n`) with each line starting `N) `
2. Inline numbering: `1) bird1 2) bird2` (all on one line)
3. ` and ` (word "and" with spaces)
4. ` & ` (ampersand with spaces)
5. 3+ spaces between entries (when each piece looks like a band combo)
6. Semicolons (`;`) when each piece looks like a band combo
7. `PP1:`, `PP2:` style observer numbering
8. `PIPL A:`, `PIPL B:` style observer labeling

---

## 8. Trailing metadata noise (strip from BandCombo)
These phrases at the end of an entry are reporter notes — strip them:
- "reported to banders" / "reported to great lakes researchers"
- "will report to plover@umn.edu"
- "photos taken" / "photo available"
- "banded in N. Nova Scotia"
- "continuing bird, will be batch reported at end of season"

---

## 9. Unreadable bands → flag for biologist
Phrases meaning birds were seen banded but codes couldn't be read:
`"none readable"`, `"unable to read"`, `"too distant"`, `"not 100%"`,
`"partial resighting"`, `"incomplete band combo recorded"`
→ Flag yellow in review sheet. Count is uncertain — biologist must confirm.

---

## 10. Count prefix (strip before parsing)
Phrases like `"5 banded PIPL seen, ..."` or `"2 banded: ..."` → strip the prefix,
then parse the remaining text. The number in the prefix is a hint for expected entry count.

---

## 11. Special codes (keep as-is, note in review)
- `FTP` — Flag code used in certain banding programs (not a color)
- `NLP` — Similar project code
- `HY` — Hatch Year
- `m` or `M` — Metal band (USFWS standard silver band)
- `_` or `-` — No band at that position

---

## 12. Status classification for review sheet
| Status | Color | Meaning |
|--------|-------|---------|
| ✓ OK | White | Confident interpretation, no biologist input needed |
| ⚠ Review | Yellow | Ambiguous — biologist should confirm |
| ✗ Conflict | Orange | Count conflict or data error — biologist must fix |
