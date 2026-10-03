"""
Extract ledger rows by "Spend/Revenue Category".

Values in that column look like "RC051 Dues" - the leading code (RC051) is
stripped, leaving just "Dues", and the script lists the unique category
names it finds so you can pick one from the command line.

Input files are read from the data/ subfolder next to this script, the same
way filter_ledger.py does. Pass bare filenames (or nothing, to use every
.xlsx in data/); absolute paths and paths that already resolve from the
current directory still work.

Output is written to the output/ subfolder, with today's date appended to
the filename (e.g. -o dues.xlsx -> output/dues_2026-09-28.xlsx).

Usage:
    python filter_by_category.py
    python filter_by_category.py file1.xlsx file2.xlsx
    python filter_by_category.py --columns "Accounting Date" Debit Credit
    python filter_by_category.py -o dues.xlsx

Requires: pip install pandas openpyxl
"""

import argparse
import re
import sys

import pandas as pd

from filter_ledger import (
    DATE_COL,
    DEFAULT_COLUMNS,
    data_files,
    load_file,
    match_columns,
    resolve_output_path,
    resolve_path,
)

CATEGORY_COL = "Spend/Revenue Category"


def strip_category_code(value):
    """'RC051 Dues' -> 'Dues'. Leaves values with no leading code untouched."""
    text = str(value).strip()
    return re.sub(r"^\S+\s+", "", text).strip()


def choose_category(categories):
    """Prompt on the CLI until the user picks a valid category."""
    print("\nAvailable Spend/Revenue Categories:")
    for i, cat in enumerate(categories, 1):
        print(f"  {i}. {cat}")
    while True:
        choice = input("\nSelect a category (number or exact name): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(categories):
            return categories[int(choice) - 1]
        matches = [c for c in categories if c.lower() == choice.lower()]
        if matches:
            return matches[0]
        print("Not a valid choice - try again.")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("files", nargs="*",
                   help="Transactions .xlsx files in data/ (default: every .xlsx in data/)")
    p.add_argument("--columns", nargs="+", default=DEFAULT_COLUMNS,
                   help="Columns to keep in the output (default: Accounting Date, Debit, Credit, "
                        "Journal Source, Spend/Revenue Category)")
    p.add_argument("--sheet", default=None, help="Sheet name (default: search all sheets)")
    p.add_argument("--no-source", action="store_true",
                   help="Don't add a column showing which file each row came from")
    p.add_argument("-o", "--output", default="ledger_by_category.xlsx",
                   help="Output filename; saved into output/ with today's date appended")
    args = p.parse_args()

    wanted = list(args.columns)
    if CATEGORY_COL.lower() not in [w.lower() for w in wanted]:
        wanted.insert(0, CATEGORY_COL)

    paths = [resolve_path(f) for f in args.files] if args.files else data_files()
    frames = []
    for path in paths:
        df, sheet, header_row = load_file(path, args.sheet)
        cols = match_columns(df, wanted, path)
        df = df[cols]
        df.columns = wanted  # consistent names across files
        if not args.no_source:
            df["Source File"] = path.name
        frames.append(df)
        print(f'{path.name}: sheet "{sheet}", headers on row {header_row + 1}, {len(df)} rows')

    result = pd.concat(frames, ignore_index=True)
    if DATE_COL in result.columns:
        result[DATE_COL] = pd.to_datetime(result[DATE_COL], errors="coerce")

    result["_stripped_category"] = result[CATEGORY_COL].apply(strip_category_code)
    categories = sorted(c for c in result["_stripped_category"].dropna().unique() if c)
    if not categories:
        sys.exit(f'No values found in "{CATEGORY_COL}"')

    chosen = choose_category(categories)
    filtered = result[result["_stripped_category"] == chosen].drop(columns=["_stripped_category"])
    if DATE_COL in filtered.columns:
        filtered = filtered.sort_values(DATE_COL, kind="stable")

    out_path = resolve_output_path(args.output)
    if out_path.suffix.lower() == ".csv":
        filtered.to_csv(out_path, index=False, date_format="%Y-%m-%d")
    else:
        with pd.ExcelWriter(out_path, engine="openpyxl", datetime_format="YYYY-MM-DD") as w:
            filtered.to_excel(w, index=False, sheet_name="Filtered")
            ws = w.sheets["Filtered"]
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 50)

    print(f'\nCategory: "{chosen}" - {len(filtered)} of {len(result)} rows')
    for c in ("Debit", "Credit"):
        if c in filtered.columns:
            total = pd.to_numeric(filtered[c], errors="coerce").sum()
            print(f"{c} total: {total:,.2f}")
    print(f"Saved to {out_path}")


if __name__ == "__main__":
    main()
