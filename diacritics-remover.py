#!/usr/bin/env python3
# script name: Exambank/ExamsProcess/diacritics-remover.py

import sys
import os
import pandas as pd
from datetime import datetime
from filename_utils import clean_filename

def main():
    if len(sys.argv) != 2:
        print("Usage: python diacritics-remover.py Input_Spreadsheet.xlsx")
        sys.exit(1)

    input_file = sys.argv[1]

    if not os.path.isfile(input_file):
        print(f"❌ Error: File {input_file} not found.")
        sys.exit(1)

    output_file = input_file.replace(".xlsx", "_cleaned.xlsx")
    log_file = input_file.replace(".xlsx", "_cleanup.log")

    # Read spreadsheet
    df = pd.read_excel(input_file)

    if df.empty or df.shape[1] < 1:
        print("❌ Error: Spreadsheet must have at least one column (Filename).")
        sys.exit(1)

    log_entries = []
    cleaned_names = []

    for i, cell_value in enumerate(df.iloc[:, 0], start=1):
        if not isinstance(cell_value, str) or not cell_value.strip():
            cleaned_names.append(cell_value)
            log_entries.append(f"Row {i}: {cell_value} (unchanged or empty)")
            continue

        # Split multiple filenames in one cell using '|'
        parts = cell_value.split("|")
        cleaned_parts = [clean_filename(part)[1] for part in parts]
        cleaned_cell = "|".join(cleaned_parts)
        cleaned_names.append(cleaned_cell)

        if cell_value != cleaned_cell:
            log_entries.append(f"Row {i}: {cell_value}  -->  {cleaned_cell}")
        else:
            log_entries.append(f"Row {i}: {cell_value} (unchanged)")

    # Replace first column with cleaned names
    df.iloc[:, 0] = cleaned_names

    # Save cleaned Excel file
    df.to_excel(output_file, index=False)

    # Write enhanced log
    with open(log_file, "w", encoding="utf-8") as f:
        f.write(f"Cleanup log - {datetime.now()}\n")
        f.write("=" * 80 + "\n")
        for entry in log_entries:
            f.write(entry + "\n")
        f.write(f"\nTotal rows processed: {len(df)}\n")
        cleaned_count = sum(1 for entry in log_entries if " --> " in entry)
        f.write(f"Total filenames changed: {cleaned_count}\n")

    print(f"✅ Cleaning done!")
    print(f"Saved cleaned spreadsheet: {output_file}")
    print(f"Log file: {log_file}")

if __name__ == "__main__":
    main()
