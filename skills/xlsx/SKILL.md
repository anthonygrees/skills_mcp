---
name: xlsx
description: "Create, read, edit, clean, or convert spreadsheets (.xlsx, .xlsm, .csv, .tsv). Use whenever a spreadsheet file is the input or the deliverable: building a workbook with formulas, formatting, charts, or validation; fixing or extending an existing workbook; cleaning messy tabular data into a proper sheet; converting between CSV and Excel; or building a financial or analytical model. Not for Word documents, HTML reports, standalone scripts, or Google Sheets API work."
license: Apache-2.0
---

# XLSX

Work with spreadsheets using `openpyxl` (formulas, formatting, charts) and `pandas` (reading, analysis, cleaning). This skill is self-contained: everything it refers to is in this folder.

## Setup

```bash
pip install openpyxl pandas
```

Optional: LibreOffice, to recalculate formulas (see [Recalculation](#recalculation)).

## Choose a path

| Task | Use | Go to |
|------|-----|-------|
| Read, summarize, analyze | pandas | [Reading](#reading) |
| Clean messy data | pandas | `references/cleaning.md` |
| Build a workbook with formulas, formatting, charts | openpyxl | `references/recipes.md` |
| Edit an existing workbook | openpyxl | [Editing](#editing) |
| Financial or scenario model | openpyxl | `references/modeling.md` |
| Check a finished workbook | script | [Checking](#checking-required) |

## Reading

```bash
python scripts/check_xlsx.py book.xlsx        # sheets, sizes, formulas, problems
```

```python
import pandas as pd
sheets = pd.read_excel("book.xlsx", sheet_name=None)       # dict of DataFrames, one per sheet
df = pd.read_excel("book.xlsx", dtype={"id": str})         # stop IDs and codes losing leading zeros
```

Read cached results with `load_workbook(path, data_only=True)`. Report what the data contains; do not modify the file when the task is only to read it. Legacy `.xls` files must be converted to `.xlsx` first (for example `soffice --headless --convert-to xlsx file.xls`, or open and re-save in Excel); tell the user if no converter is available.

## Core rule: formulas, not pasted results

When a cell should be a calculation, write a formula, not a number computed in Python.

```python
ws["B10"] = "=SUM(B2:B9)"            # good: updates when inputs change
ws["B10"] = df["Sales"].sum()        # bad: frozen value, silently goes stale
```

This applies to totals, percentages, ratios, growth, and differences. Put inputs such as rates and thresholds in their own labelled cells and reference them (`=B5*(1+$B$6)`, not `=B5*1.05`).

Exceptions: raw data you were given, and values the user explicitly wants frozen.

## Creating

1. **Clarify the output.** Which sheets, columns, and calculations; who will use it; whether a template applies.
2. **Lay out the data** with one header row, one record per row, one value per cell. No blank rows or columns inside a table, no merged cells in data ranges.
3. **Write data, then formulas, then formatting** (see `references/recipes.md`).
4. **Document assumptions** in labelled cells or cell comments, and the source of any hardcoded figure.
5. **Save, recalculate if you can, then run the checks** below.

Defaults unless told otherwise: one consistent font (Arial or Calibri, 10-11pt); bold header row with a fill; frozen header row; sensible column widths; number formats that match the data (never leave dates or currency as General); gridlines off only on presentation sheets.

## Editing

1. **Work on a copy.** Save under a new filename unless the user asks to overwrite.
2. **Study the existing conventions first** (fonts, fills, number formats, layout, naming) and match them exactly. Existing template conventions override the defaults above.
3. Load with `load_workbook(path)`. Use `keep_vba=True` for `.xlsm` and save with the `.xlsm` extension.
4. **Never load with `data_only=True` and then save.** That replaces every formula with its last cached value, permanently.
5. Make only the requested changes. Inserting or deleting rows and columns with `insert_rows` / `delete_cols` does **not** update formulas, named ranges, charts, or conditional formatting that refer to the shifted cells; fix those references yourself and verify.
6. `openpyxl` can drop features it does not support when it saves: pivot tables, slicers, form controls, some chart formatting, and macros unless `keep_vba=True`. Run the check script on the original and on the result and compare sheets, charts, and images. Warn the user if the file has features at risk.

## Recalculation

`openpyxl` stores formulas as text; it does not calculate them. A file it writes has **no cached values**, so a viewer that does not calculate (a preview pane, a script reading with `data_only=True`) shows blanks, and errors in formulas cannot be seen.

Excel calculates when the file is opened. To get values in the file now, and to find errors, recalculate with LibreOffice if it is installed:

```bash
soffice --headless --convert-to xlsx --outdir recalculated/ output.xlsx
python scripts/check_xlsx.py recalculated/output.xlsx
```

The converted copy has cached values and is the one to check. Keep the converted copy as the deliverable only if the user is happy with LibreOffice's rewrite (it can change styles slightly); otherwise deliver the original and say that values appear when opened in Excel.

If LibreOffice is not available, say so plainly. Then verify the numbers another way: compute the expected totals independently in pandas and compare them with what the formulas should produce, and review every formula by reading it. Do not claim the workbook is error-free when it has not been calculated.

## Checking (required)

```bash
python scripts/check_xlsx.py output.xlsx
```

It reports:

- **Errors:** cells holding `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`, and similar; formulas pointing at sheets that do not exist.
- **Warnings:** formulas with no cached values (never recalculated); numbers stored as text.
- **Notes:** hardcoded numbers inside formulas, which usually belong in an assumption cell.

Fix every error and warning, or explain why it stands, and rerun. "Never recalculated" is expected for a fresh `openpyxl` file until you recalculate it; it means the error check has not actually run.

Then verify by hand:

- Spot-check two or three formulas against the source data.
- Excel rows and columns are 1-based. DataFrame row `n` is Excel row `n + 2` after a header. Confirm column letters with `get_column_letter` instead of counting.
- Check ranges cover every row (no off-by-one) and totals do not include their own cell (circular reference).
- Guard divisions that can hit zero: `=IF(B2=0,0,A2/B2)` or `=IFERROR(A2/B2,"")` where hiding the error is acceptable and obvious.

## Rules of thumb

- One header row; one table per region; no data hidden in merged cells.
- Dates are real dates with a date format, not text. Numbers are numbers, not text.
- Never type totals; use `SUM` over the range.
- Label units in headers (`Revenue ($000)`).
- Use named ranges or an Assumptions sheet for inputs a user will change.
- Do not invent data. Mark missing values clearly and tell the user.
- Do not put credentials, internal URLs, or personal data into a workbook unless the user supplied them for that purpose.
