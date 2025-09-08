#!/usr/bin/env python3
# script: rename_raws_in_excel.py

import sys
import pandas as pd
from pathlib import Path

def convert_to_jpeg(filename: str) -> str:
    """
    Convert CR2/CR3 files to .jpeg.
    Skip existing jpeg/jpg/png/gif files.
    """
    if not isinstance(filename, str):
        return filename
    
    ext = Path(filename).suffix.lower()
    if ext in [".jpeg", ".jpg", ".png", ".gif"]:
        return filename  # leave untouched
    if ext in [".cr2", ".cr3"]:
        return str(Path(filename).with_suffix(".jpeg"))
    return filename

def process_excel(input_file: str):
    # Output filenames
    output_file = input_file.replace(".xlsx", "_jpeg.xlsx")
    log_file = input_file.replace(".xlsx", "_converted_ext.log")

    # Load spreadsheet
    df = pd.read_excel(input_file)

    if df.empty or df.shape[1] < 1:
        print("❌ Error: Spreadsheet must have at least one column (Filename).")
        sys.exit(1)

    # Prepare log list
    log_changes = []

    # Process first column (row by row)
    cleaned_rows = []
    for idx, cell in enumerate(df.iloc[:, 0], start=2):  # row numbers (Excel rows start at 2 with header)
        if not isinstance(cell, str):
            cleaned_rows.append(cell)
            continue

        filenames = [fn.strip() for fn in cell.split("|")]
        converted = []
        changed = False

        for fn in filenames:
            new_fn = convert_to_jpeg(fn)
            converted.append(new_fn)
            if new_fn != fn:
                changed = True

        if changed:
            log_changes.append(f"Row {idx}: {cell}  ->  {' | '.join(converted)}")

        cleaned_rows.append(" | ".join(converted))

    # Replace first column with cleaned version
    df.iloc[:, 0] = cleaned_rows

    # Save new Excel
    df.to_excel(output_file, index=False)
    print(f"✅ Updated spreadsheet saved to: {output_file}")

    # Write log file
    if log_changes:
        with open(log_file, "w", encoding="utf-8") as f:
            f.write("\n".join(log_changes))
        print(f"📝 Change log written to: {log_file}")
    else:
        print("ℹ️ No CR2/CR3 filenames found, nothing converted.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python rename_raws_in_excel.py Input_Spreadsheet.xlsx")
        sys.exit(1)

    process_excel(sys.argv[1])
