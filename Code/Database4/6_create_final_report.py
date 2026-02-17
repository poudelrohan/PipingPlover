"""
Step 6: Final Report Generation
─────────────────────────────────
Combines all pipeline outputs into a single professional Excel file
with 3 sheets:

  Sheet 1 — Clean_Data
      The fully processed, validated rows ready for biologist use.

  Sheet 2 — Removed_Rows
      Every row that was removed at any step, with columns:
        unique_id, source_file, source_sheet, removal_reason,
        + all original data columns

  Sheet 3 — Summary_Report
      Processing statistics:
        - Total rows input
        - Rows removed per reason (counts + percentages)
        - Rows in final clean dataset
        - Partial location warnings (rows kept but missing some location data)
        - Duplicate criteria used
        - Date pipeline was run

Output: database4_FINAL.xlsx
"""

import pandas as pd
import os
import sys
from datetime import datetime

# ── Load config ────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from database4_config import config

# ── Resolve paths ──────────────────────────────────────────────────────────────
script_dir    = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.normpath(os.path.join(script_dir, config["output"]["folder"]))

clean_path   = os.path.join(output_folder, "database4_clean.xlsx")
removed_path = os.path.join(output_folder, "database4_removed.xlsx")
ids_path     = os.path.join(output_folder, "database4_with_ids.xlsx")
final_path   = os.path.join(output_folder, "database4_FINAL.xlsx")

# ── Load data ──────────────────────────────────────────────────────────────────
for path, label in [(clean_path, "Clean data (Step 5)"), (removed_path, "Removed rows"), (ids_path, "Original with IDs (Step 1)")]:
    if not os.path.exists(path):
        print(f"[ERROR] Missing file: {path} ({label})")
        print("        Run all previous steps first.")
        sys.exit(1)

clean_df   = pd.read_excel(clean_path)
removed_df = pd.read_excel(removed_path)
original   = pd.read_excel(ids_path)

print(f"  Clean rows   : {len(clean_df)}")
print(f"  Removed rows : {len(removed_df)}")
print(f"  Total input  : {len(original)}")

# ── Sheet 2: Removed_Rows ──────────────────────────────────────────────────────
# Reorder so key info is first
removed_sheet = removed_df.copy()
priority_cols = ["unique_id", "source_file", "source_sheet", "_removal_reason"]
other_cols    = [c for c in removed_sheet.columns if c not in priority_cols]
removed_sheet = removed_sheet[
    [c for c in priority_cols if c in removed_sheet.columns] + other_cols
]
removed_sheet = removed_sheet.rename(columns={"_removal_reason": "removal_reason"})

# ── Sheet 3: Summary_Report ────────────────────────────────────────────────────
total_input   = len(original)
total_clean   = len(clean_df)
total_removed = len(removed_df)

# Breakdown by removal reason
reason_counts = removed_df["_removal_reason"].value_counts().reset_index()
reason_counts.columns = ["removal_reason", "count"]
reason_counts["percentage"] = (reason_counts["count"] / total_input * 100).round(2).astype(str) + "%"

# Partial location warnings (rows kept but missing some location fields)
loc_fields = config["location_fields"]["fields"]
partial_warnings = []
for _, row in clean_df.iterrows():
    missing = [f for f in loc_fields if f in clean_df.columns and (pd.isna(row.get(f)) or str(row.get(f)).strip() == "")]
    if 0 < len(missing) < len(loc_fields):
        partial_warnings.append({
            "unique_id":      row.get("unique_id"),
            "missing_fields": ", ".join(missing)
        })

# Build summary rows
summary_rows = [
    ("Run date",                    datetime.now().strftime("%Y-%m-%d %H:%M")),
    ("Database",                    config["database_name"]),
    ("",                            ""),
    ("── Input ──",                 ""),
    ("Total rows input",            total_input),
    ("Source files",                ", ".join(config["input"]["files"])),
    ("",                            ""),
    ("── Output ──",                ""),
    ("Total rows in Clean_Data",    total_clean),
    ("Total rows removed",          total_removed),
    ("Pct rows kept",               f"{round(total_clean / total_input * 100, 2)}%"),
    ("",                            ""),
    ("── Removal Breakdown ──",     ""),
]

for _, r in reason_counts.iterrows():
    summary_rows.append((f"  {r['removal_reason']}", f"{r['count']} rows ({r['percentage']})"))

summary_rows += [
    ("",                            ""),
    ("── Location Warnings ──",     ""),
    ("Rows with partial location data (kept but flagged)", len(partial_warnings)),
]

if partial_warnings:
    for w in partial_warnings:
        summary_rows.append((f"  unique_id {w['unique_id']}", f"Missing: {w['missing_fields']}"))

summary_rows += [
    ("",                            ""),
    ("── Duplicate Criteria ──",    ""),
    ("Columns used",                ", ".join(config["duplicate_criteria"])),
]

summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

# ── Write final Excel with 3 sheets ───────────────────────────────────────────
with pd.ExcelWriter(final_path, engine="openpyxl") as writer:
    clean_df.to_excel(writer,    sheet_name="Clean_Data",     index=False)
    removed_sheet.to_excel(writer, sheet_name="Removed_Rows", index=False)
    summary_df.to_excel(writer,  sheet_name="Summary_Report", index=False)

    # ── Basic formatting ───────────────────────────────────────────────────────
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = writer.book

    header_font  = Font(bold=True, color="FFFFFF")
    header_fills = {
        "Clean_Data":     PatternFill("solid", fgColor="2E7D32"),  # dark green
        "Removed_Rows":   PatternFill("solid", fgColor="C62828"),  # dark red
        "Summary_Report": PatternFill("solid", fgColor="1565C0"),  # dark blue
    }

    for sheet_name in ["Clean_Data", "Removed_Rows", "Summary_Report"]:
        ws   = wb[sheet_name]
        fill = header_fills[sheet_name]

        # Style header row
        for cell in ws[1]:
            cell.font      = header_font
            cell.fill      = fill
            cell.alignment = Alignment(horizontal="center")

        # Auto-fit column widths
        for col in ws.columns:
            max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)

print(f"\n[DONE] Step 6 complete")
print(f"  Final report : {final_path}")
print(f"  Sheets       : Clean_Data ({total_clean} rows) | Removed_Rows ({total_removed} rows) | Summary_Report")
