"""
Filter one or more transactions exports to rows whose "Accounting Date" is
after a cutoff, keep only selected columns, and combine them into one ledger.

Input files are read from the data/ subfolder next to this script. Pass bare
filenames (or nothing at all, to use every .xlsx in data/); absolute paths and
paths that already resolve from the current directory still work.

Output is written to the output/ subfolder, with today's date appended to the
filename (e.g. -o ledger.xlsx -> output/ledger_2026-09-28.xlsx).

Usage:
    python filter_ledger.py                      (every .xlsx in data/)
    python filter_ledger.py file1.xlsx file2.xlsx
    python filter_ledger.py file1.xlsx file2.xlsx --inclusive
    python filter_ledger.py --cutoff 2026-04-01 -o ledger_apr_onward.xlsx
    python filter_ledger.py --inclusive -o ledger.csv   (CSV output)
    python filter_ledger.py file1.xlsx --columns "Accounting Date" Debit Credit "Journal Source" Description

Requires: pip install pandas openpyxl
"""

import argparse
import sys
from datetime import date
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).resolve().parent / "data"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"
DATE_COL = "Accounting Date"
DEFAULT_COLUMNS = ["Accounting Date", "Debit", "Credit", "Journal Source", "Spend/Revenue Category"]


def data_files():
    """Every .xlsx in data/, used when no files are named on the command line."""
    if not DATA_DIR.is_dir():
        sys.exit(f"No data folder at {DATA_DIR} - create it and put the exports there.")
    found = sorted(f for f in DATA_DIR.glob("*.xlsx") if not f.name.startswith("~$"))
    if not found:
        sys.exit(f"No .xlsx files in {DATA_DIR}")
    return found


def resolve_path(name):
    """Look for the file in data/ first, then fall back to the path as given
    (so absolute paths and files elsewhere still work)."""
    given = Path(name)
    in_data = DATA_DIR / given.name
    if not given.is_absolute() and in_data.is_file():
        return in_data
    if given.is_file():
        return given
    sys.exit(f"{name}: not found in {DATA_DIR} or as a path from here.")


def resolve_output_path(name):
    """Route output into output/ and stamp the filename with today's date,
    e.g. 'ledger_filtered.xlsx' -> 'output/ledger_filtered_2026-09-28.xlsx'."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    given = Path(name)
    stamped = f"{given.stem}_{date.today():%Y-%m-%d}{given.suffix}"
    return OUTPUT_DIR / stamped


def find_header_row(path, sheet):
    """Exports often have title rows above the real headers.
    Scan the first 30 rows for the one containing the date column name."""
    preview = pd.read_excel(path, sheet_name=sheet, header=None, nrows=30)
    for i, row in preview.iterrows():
        cells = [str(c).strip().lower() for c in row.values]
        if DATE_COL.lower() in cells:
            return i
    return None


def load_file(path, sheet_arg):
    sheets = [sheet_arg] if sheet_arg else pd.ExcelFile(path).sheet_names
    for sheet in sheets:
        header_row = find_header_row(path, sheet)
        if header_row is not None:
            df = pd.read_excel(path, sheet_name=sheet, header=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            return df, sheet, header_row
    sys.exit(f'{path}: could not find a "{DATE_COL}" column in: {", ".join(sheets)}')


def match_columns(df, wanted, path):
    """Match requested columns case-insensitively; stop if any are missing."""
    lookup = {c.lower(): c for c in df.columns}
    missing = [w for w in wanted if w.lower() not in lookup]
    if missing:
        sys.exit(f"{path}: column(s) not found: {', '.join(missing)}\n"
                 f"Available columns: {', '.join(df.columns)}")
    return [lookup[w.lower()] for w in wanted]


def main():
    p = argparse.ArgumentParser()
    p.add_argument("files", nargs="*",
                   help="Transactions .xlsx files in data/ (default: every .xlsx in data/)")
    p.add_argument("--cutoff", default="2026-04-01", help="YYYY-MM-DD (default 2026-04-01)")
    p.add_argument("--inclusive", action="store_true",
                   help="Include transactions ON the cutoff date too")
    p.add_argument("--columns", nargs="+", default=DEFAULT_COLUMNS,
                   help="Columns to keep (default: Accounting Date, Debit, Credit, Journal Source, "
                        "Spend/Revenue Category)")
    p.add_argument("--sheet", default=None, help="Sheet name (default: search all sheets)")
    p.add_argument("--no-source", action="store_true",
                   help="Don't add a column showing which file each row came from")
    p.add_argument("-o", "--output", default="ledger_filtered.xlsx",
                   help="Output filename; saved into output/ with today's date appended")
    args = p.parse_args()

    wanted = list(args.columns)
    if DATE_COL.lower() not in [w.lower() for w in wanted]:
        wanted.insert(0, DATE_COL)

    cutoff = pd.Timestamp(args.cutoff)
    paths = [resolve_path(f) for f in args.files] if args.files else data_files()
    out_path = resolve_output_path(args.output)
    frames = []

    for path in paths:
        df, sheet, header_row = load_file(path, args.sheet)
        cols = match_columns(df, wanted, path)
        df = df[cols]
        df.columns = wanted  # consistent names across files
        df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")

        unparsed = df[DATE_COL].isna() & df.notna().any(axis=1)
        mask = df[DATE_COL] >= cutoff if args.inclusive else df[DATE_COL] > cutoff
        kept = df[mask].copy()
        if not args.no_source:
            kept["Source File"] = path.name
        frames.append(kept)

        print(f'{path.name}: sheet "{sheet}", headers on row {header_row + 1}, '
              f"kept {len(kept)} of {len(df)} rows")
        if unparsed.any():
            print(f"  Note: {unparsed.sum()} non-empty rows had no readable date and were "
                  f"skipped (often subtotal lines) - worth a quick look in the original.")

    result = pd.concat(frames, ignore_index=True).sort_values(DATE_COL, kind="stable")

    # Flag exact duplicates, in case the two files overlap
    dupes = result.drop(columns=["Source File"], errors="ignore").duplicated(keep=False)
    if dupes.any():
        print(f"\nWarning: {dupes.sum()} rows are identical across/within files. "
              f"If the exports overlap, these may be double-counted - check them.")

    if out_path.suffix.lower() == ".csv":
        result.to_csv(out_path, index=False, date_format="%Y-%m-%d")
    else:
        with pd.ExcelWriter(out_path, engine="openpyxl", datetime_format="YYYY-MM-DD") as w:
            result.to_excel(w, index=False, sheet_name="Ledger")
            ws = w.sheets["Ledger"]
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 50)

    op = ">=" if args.inclusive else ">"
    print(f"\nTotal: {len(result)} rows where {DATE_COL} {op} {cutoff.date()}")
    if len(result):
        print(f"Date range: {result[DATE_COL].min().date()} to {result[DATE_COL].max().date()}")
    for c in ("Debit", "Credit"):
        if c in result.columns:
            total = pd.to_numeric(result[c], errors="coerce").sum()
            print(f"{c} total: {total:,.2f}")
    print(f"Saved to {out_path}")


if __name__ == "__main__":
    main()