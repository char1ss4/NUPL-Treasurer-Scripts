# NUPL Treasurer Scripts

Small command-line tools for processing Northeastern University (women's)
powerlifting treasurer transaction exports — filtering a ledger by date and
splitting it out by spend/revenue category.

## Setup

```
pip install pandas openpyxl
```

Drop your `.xlsx` transaction exports into the `data/` folder. Both scripts
default to using every `.xlsx` file in `data/` when no filenames are given,
and will also accept bare filenames (looked up in `data/`) or full paths.

Output files are written to `output/`, with today's date appended to the
filename (e.g. `-o ledger.xlsx` -> `output/ledger_2026-09-28.xlsx`).

## Scripts

### `filter_ledger.py`

Filters one or more transaction exports to rows whose `Accounting Date` is
after (or on, with `--inclusive`) a cutoff date, keeps a chosen set of
columns, and combines everything into a single ledger. Also flags duplicate
rows in case input files overlap.

```
python filter_ledger.py                      # every .xlsx in data/
python filter_ledger.py file1.xlsx file2.xlsx
python filter_ledger.py file1.xlsx file2.xlsx --inclusive
python filter_ledger.py --cutoff 2026-04-01 -o ledger_apr_onward.xlsx
python filter_ledger.py --inclusive -o ledger.csv   # CSV output
python filter_ledger.py file1.xlsx --columns "Accounting Date" Debit Credit "Journal Source" Description
```

Key options:
- `--cutoff YYYY-MM-DD` — date to filter from (default `2026-04-01`)
- `--inclusive` — include transactions on the cutoff date
- `--columns ...` — columns to keep (default: Accounting Date, Debit, Credit,
  Journal Source, Spend/Revenue Category)
- `--sheet NAME` — restrict to one sheet (default: search all sheets)
- `--no-source` — omit the "Source File" column
- `-o/--output NAME` — output filename (`.xlsx` or `.csv`)

### `filter_by_category.py`

Loads transactions the same way as `filter_ledger.py`, strips the leading
code from `Spend/Revenue Category` values (e.g. `"RC051 Dues"` -> `"Dues"`),
lists the unique categories found, and prompts you to pick one to export.

```
python filter_by_category.py
python filter_by_category.py file1.xlsx file2.xlsx
python filter_by_category.py --columns "Accounting Date" Debit Credit
python filter_by_category.py -o dues.xlsx
```

Shares the `--columns`, `--sheet`, `--no-source`, and `-o/--output` options
with `filter_ledger.py`.

## Notes

- Both scripts scan the first 30 rows of each sheet to find the header row
  (exports often have title rows above the real headers), and match
  requested column names case-insensitively.
- Debit/Credit totals and the matched date range are printed to the console
  after each run.
