# openpyxl recipes

All snippets were run against `openpyxl` 3.1. They assume:

```python
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName
```

Rows and columns are 1-based. `ws.cell(row=1, column=1)` is `A1`. Use `get_column_letter(64)` (`BL`) and `column_index_from_string("BL")` (64) rather than counting.

## New workbook, data, formulas

```python
wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["Region", "Q1", "Q2", "Q3", "Q4", "Total"])
for row in [("North", 120, 135, 150, 160), ("South", 90, 95, 110, 105), ("East", 70, 80, 85, 99)]:
    ws.append(row)
last = ws.max_row                                    # last data row

for r in range(2, last + 1):
    ws[f"F{r}"] = f"=SUM(B{r}:E{r})"
ws[f"A{last + 1}"] = "Total"
for col in "BCDEF":
    ws[f"{col}{last + 1}"] = f"=SUM({col}2:{col}{last})"

wb.save("output.xlsx")
```

Formulas are strings starting with `=`, in US English function names with comma separators. Quote sheet names that contain spaces: `='Q1 Data'!B2`.

## Fonts, fills, alignment, borders

```python
for c in ws[1]:                                      # header row
    c.font = Font(name="Arial", bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for c in ws[last + 1]:                               # total row
    c.font = Font(name="Arial", bold=True)
    c.border = Border(top=Side(style="thin"))
```

Colours are hex `RRGGBB` without `#`. Styles are per cell; for big ranges loop over `ws.iter_rows(min_row=..., max_row=..., min_col=..., max_col=...)`.

## Number formats

```python
c.number_format = "#,##0"                            # thousands separator
c.number_format = '#,##0;(#,##0);"-"'                # negatives in brackets, zeros as a dash
c.number_format = "$#,##0.00"
c.number_format = "0.0%"                             # value must be a fraction: 0.125 -> 12.5%
c.number_format = "0.0x"                             # multiples
c.number_format = "yyyy-mm-dd"                       # dates
```

Write dates as `datetime.date` or `datetime.datetime` values, not strings. Format years that are labels (2024) as text so they do not show as 2,024.

## Column widths, freeze panes, filters

```python
ws.column_dimensions["A"].width = 18
ws.row_dimensions[1].height = 24
ws.freeze_panes = "B2"                               # freeze header row and first column
ws.auto_filter.ref = f"A1:F{last}"
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = "1F3864"
```

There is no auto-fit. Estimate the width from the longest value (`max(len(str(v))) + 2`) and cap it.

## Multiple sheets and cross-sheet formulas

```python
a = wb.create_sheet("Assumptions")
a["A1"] = "Growth rate"
a["B1"] = 0.05
a["B1"].number_format = "0.0%"

ws["H2"] = "=F5*(1+Assumptions!$B$1)"                # cross-sheet reference

wb.defined_names["GrowthRate"] = DefinedName("GrowthRate", attr_text="Assumptions!$B$1")
ws["H3"] = "=F5*(1+GrowthRate)"                      # named range
```

Reorder with `wb.move_sheet("Assumptions", offset=-1)`; hide with `ws.sheet_state = "hidden"`.

## Conditional formatting

```python
rng = f"B2:E{last}"
ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["140"],
                              fill=PatternFill("solid", start_color="C6EFCE")))
ws.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color="FFFFFF",
                              end_type="max", end_color="63BE7B"))
ws.conditional_formatting.add(f"A2:A{last}", FormulaRule(formula=["$F2>400"], font=Font(bold=True)))
```

## Data validation (drop-downs, limits)

```python
dv = DataValidation(type="list", formula1='"North,South,East,West"', allow_blank=False)
ws.add_data_validation(dv)
dv.add("A8:A20")

nums = DataValidation(type="decimal", operator="between", formula1="0", formula2="1")
nums.error = "Enter a value between 0 and 1"
ws.add_data_validation(nums)
nums.add("C2:C20")
```

## Excel tables

```python
tab = Table(displayName="Sales", ref=f"A1:F{last}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tab)
```

Table headers must be unique non-empty strings, and the range must not overlap another table or merged cells.

## Charts

```python
ch = BarChart()
ch.type = "col"
ch.title = "Sales by region"
ch.y_axis.title = "Units"
ch.add_data(Reference(ws, min_col=2, max_col=5, min_row=1, max_row=last), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
ch.height, ch.width = 8, 16                          # centimetres
ws.add_chart(ch, "H6")
```

`LineChart` and `PieChart` work the same way. One message per chart; title it with the takeaway.

## Comments, hyperlinks, merged cells

```python
ws["J2"].comment = Comment("Source: ledger export, 2025-01-31", "Analyst")
ws["J3"] = "Docs"
ws["J3"].hyperlink = "https://example.com"
ws["J3"].style = "Hyperlink"
ws.merge_cells("J5:K5")                              # titles only; never inside data ranges
```

## Print setup

```python
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:1"
```

## Editing an existing workbook

```python
wb = load_workbook("existing.xlsx")                  # xlsm: load_workbook(p, keep_vba=True)
for name in wb.sheetnames:
    print(name, wb[name].dimensions)
ws = wb["Sheet1"]
ws["A1"] = "New value"
wb.save("modified.xlsx")                             # new filename
```

Large files: `load_workbook(p, read_only=True)` to read; `Workbook(write_only=True)` to write (append-only, no random cell access).

## Gotchas

- `ws.append` writes to the first empty row after the last used row, including rows that only hold formatting.
- `ws.max_row` counts rows that only contain styling; for the last row with data, scan upward for a non-empty cell.
- Percent cells hold fractions: write `0.05`, format as `0.0%`.
- A value that starts with `=` is stored as a formula. To store literal text beginning with `=`, set `cell.data_type = "s"` after assigning the value.
- openpyxl does not evaluate formulas, so `ws["B10"].value` on a new file is the formula text.
- Merged cells keep their value in the top-left cell only.
- Sheet names: 31 characters maximum, and none of `\ / ? * [ ] :`.
