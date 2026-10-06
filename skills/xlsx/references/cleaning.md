# Reading and cleaning tabular data

Patterns for turning messy spreadsheets and CSVs into a clean table. All run against `pandas` 2.x. Work on a copy and keep the raw file unchanged.

## Read without guessing

```python
import pandas as pd

raw = pd.read_excel("messy.xlsx", header=None, dtype=str)     # nothing inferred, nothing dropped
raw = pd.read_excel("messy.xlsx", sheet_name=None, header=None, dtype=str)   # every sheet
df  = pd.read_csv("data.csv", dtype=str, encoding="utf-8-sig")  # utf-8-sig strips an Excel BOM
```

Reading everything as text first means leading zeros, long IDs, and odd values survive. Convert types deliberately afterwards. For other CSV dialects pass `sep=";"` or `sep="\t"`; try `encoding="latin-1"` if UTF-8 fails, and say which you used.

Useful options for large files: `usecols=["A", "C"]`, `nrows=1000` to preview, `skiprows=3` when you already know the header position.

## Find the real header

```python
hdr = raw.index[raw.iloc[:, 0].eq("Name")][0]          # row whose first cell is a known header
df = raw.iloc[hdr + 1:].copy()
df.columns = raw.iloc[hdr]
df.columns.name = None
df = df.reset_index(drop=True)
```

When the header text is not known, pick the first row with several non-empty cells: `raw.notna().sum(axis=1).gt(2).idxmax()`.

## Remove junk

```python
df = df.dropna(how="all")                              # blank rows
df = df.dropna(axis=1, how="all")                      # blank columns
df = df[df["Name"] != "Name"]                          # header repeated mid-file (page breaks)
df = df.drop_duplicates()
df.columns = df.columns.str.strip()
df["Name"] = df["Name"].str.strip()
```

Filling merged-cell gaps (value only in the first row of a group): `df["Region"] = df["Region"].ffill()`.

## Convert types

```python
df["Amount"] = (
    df["Amount"].str.strip()
    .str.replace(r"[$,\s]", "", regex=True)            # currency symbols, thousands separators
    .str.replace(r"^\((.*)\)$", r"-\1", regex=True)    # (30) -> -30
    .pipe(pd.to_numeric, errors="coerce")
)
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")           # add dayfirst=True for DD/MM/YYYY
df["Active"] = df["Active"].str.lower().map({"yes": True, "no": False, "y": True, "n": False})
```

`errors="coerce"` turns unparseable values into `NaN`/`NaT`. **Count them before moving on**, and report what you did with them instead of hiding them:

```python
bad = df[df["Amount"].isna() & raw_amount.notna()]     # raw_amount: the pre-conversion column
print(len(bad), "rows could not be parsed")
```

Percent strings: `"12.5%"` -> `pd.to_numeric(s.str.rstrip("%")) / 100`. Ambiguous dates (`03/04/2025`) need the user's locale; ask rather than guess.

## Reshape

```python
long = df.melt(id_vars=["Region"], var_name="Quarter", value_name="Sales")           # wide -> long
wide = long.pivot_table(index="Region", columns="Quarter", values="Sales", aggfunc="sum")  # long -> wide
summary = df.groupby("Region", as_index=False)["Sales"].sum()
merged = df.merge(lookup, on="id", how="left", validate="m:1")                      # validate catches duplicate keys
```

## Write the clean result

```python
from openpyxl.utils import get_column_letter

with pd.ExcelWriter("clean.xlsx", engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
    df.to_excel(xw, sheet_name="Clean", index=False)
    ws = xw.sheets["Clean"]                            # openpyxl worksheet: format it here
    ws.freeze_panes = "A2"
    for i, col in enumerate(df.columns, 1):
        width = max(len(str(col)), *(len(str(v)) for v in df[col])) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 50)
```

Always `index=False` unless the index carries data. CSV to Excel: `pd.read_csv(p, dtype=str).to_excel("out.xlsx", index=False)`; Excel to CSV: `df.to_csv("out.csv", index=False, encoding="utf-8-sig")` (the BOM keeps Excel from mangling non-ASCII text).

`to_excel` writes values, not formulas. If the cleaned sheet needs totals or derived columns, add them afterwards as formulas with openpyxl (`references/recipes.md`).

## Report what changed

After cleaning, tell the user: rows in and out, rows dropped and why, values that could not be parsed, columns renamed, and any assumptions (date format, currency, encoding). Offer to keep a "Rejected" sheet with the rows that were removed, so nothing disappears silently.

## Reading values from a workbook that has formulas

- `pd.read_excel` returns **cached values**. For a workbook that has never been calculated those are empty.
- To see formulas, use `load_workbook(path)` and read `cell.value`.
- To see values, use `load_workbook(path, data_only=True)`. Never save a workbook opened this way: it permanently replaces formulas with values.
