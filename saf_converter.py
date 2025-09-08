#!/usr/bin/env python3
"""
Flexible Excel -> DSpace Dublin Core SAF packager.

Input:
  - Input_Spreadsheet_Cleanitup_cleaned.xlsx  (first rows may contain template text; header row is auto-detected)
  - ./bitstreams_dir/                         (bitstreams to copy; names come from filename-like column)

Output:
  - ./saf_converted/<N>/
      - dublin_core.xml
      - contents
      - copied bitstream files (if present)

Logging:
  - process.log  (missing files, created folders, header detection info)
"""

import os
import re
import math
import shutil
import logging
from pathlib import Path

try:
    import pandas as pd
except Exception:
    print("ERROR: pandas not installed. Install with:\n  pip install pandas openpyxl")
    raise

import xml.etree.ElementTree as ET

# ---------------------------------------------------
# Configuration
# ---------------------------------------------------
class Config:
    INPUT_XLSX = "Input_Spreadsheet_cleaned.xlsx"
    FILES_DIR = Path("bitstreams_dir")
    OUTPUT_DIR = Path("saf_converted")
    LOG_FILE = "saf_converter_process.log"
    LANG_ATTR = "en"
    MAX_HEADER_SCAN = 25
    ALLOWED_FILE_EXTENSIONS = [".cr2",".cr3",".pdf", ".doc", ".docx", ".jpg", ".jpeg", ".png",".gif", ".tiff", ".mp3", ".mp4",".wav", ".avi", ".mov",".mpeg", ".mpg", ".txt", ".rtf", ".xls", ".xlsx", ".zip", ".rar", ".7z"]
    REQUIRED_FIELDS = ["title", "creator", "date.issued"]  # Required DC fields

# ---------------------------------------------------
# Logging setup
# ---------------------------------------------------
logging.basicConfig(
    filename=Config.LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
console = logging.getLogger("console")
console.setLevel(logging.INFO)
_console_handler = logging.StreamHandler()
_console_handler.setLevel(logging.INFO)
_console_handler.setFormatter(logging.Formatter("%(message)s"))
console.addHandler(_console_handler)

# ---------------------------------------------------
# Functions
# ---------------------------------------------------

# ---- helpers ----
def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and v.strip().lower() in ["", "nan", "none"]:
        return True
    return False

def safe_mkdir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def basename_only(name: str) -> str:
    return os.path.basename(str(name).strip())

def norm_header(h: str) -> str:
    return str(h).strip()

def norm_key_for_match(h: str) -> str:
    """normalize header for matching (filename detection)"""
    return re.sub(r"[\s_\-\.]+", "", str(h).strip().lower())

def find_header_row(xlsx_path: str, max_scan: int = Config.MAX_HEADER_SCAN) -> int:
    """
    Find the header row by scanning the first few rows for:
      - a 'filename' cell (case/space-insensitive), and
      - several 'dc.' cells
    Fallback: first row that has any 'dc.' cell; else 0
    """
    probe = pd.read_excel(xlsx_path, header=None, nrows=max_scan, dtype=str)
    best_idx = None
    best_dc_count = -1

    for i, row in probe.iterrows():
        cells = [str(x).strip() for x in row.tolist() if not is_blank(x)]
        lc = [c.lower() for c in cells]
        has_filename = any("filename" in norm_key_for_match(c) for c in cells)
        dc_count = sum(1 for c in lc if c.startswith("dc."))

        console.info(f"Row {i}: has_filename={has_filename}, dc_count={dc_count}, cells={cells}")

        if has_filename and dc_count >= 2:
            logging.info(f"Header row detected at index {i} (has 'Filename' + {dc_count} dc.* cells).")
            return i
        # remember the row with most dc.* if no filename found yet
        if dc_count > best_dc_count:
            best_dc_count = dc_count
            best_idx = i

    if best_dc_count > 0 and best_idx is not None:
        logging.info(f"Header row guessed at index {best_idx} (max dc.* cells = {best_dc_count}).")
        return best_idx

    logging.info("Header row fallback to index 0.")
    return 0

def detect_filename_columns(columns):
    """
    Return a list of columns to consider as filename sources.
    Accepts: Filename, file name, file_name, dc.filename (etc.)
    """
    cands = []
    for c in columns:
        if c is None:
            continue
        key = norm_key_for_match(c)
        if "filename" in key or "file" in key or "bitstream" in key:
            cands.append(c)
    return cands

def is_dc_column(col: str) -> bool:
    if col is None:
        return False
    return col.strip().lower().startswith("dc.")

def base_header(col: str) -> str:
    """strip trailing '.<n>' so dc.subject.1 -> dc.subject for mapping"""
    if col is None:
        return ""
    return re.sub(r"\.(\d+)$", "", col)

def parse_dc(col: str):
    """
    dc.title                -> ('title', None)
    dc.description.abstract -> ('description', 'abstract')
    dc.language.iso         -> ('language', 'iso')
    """
    if col is None:
        return None, None

    c = col.strip().lower()
    if not c.startswith("dc."):
        return None, None
    parts = c.split(".")
    if len(parts) == 2:
        return parts[1], None
    elif len(parts) >= 3:
        return parts[1], ".".join(parts[2:])  # allow multi-qualifier like description.abstract
    return None, None

# ---- Enhanced file discovery ----
def find_files_for_item(filenames, bitstreams_dir):
    """Try multiple strategies to locate files"""
    found_files = []

    for fname in filenames:
        # Strategy 1: Exact match
        src = bitstreams_dir / fname
        if src.exists():
            found_files.append(fname)
            continue

        # Strategy 2: Case-insensitive match
        pattern = re.compile(re.escape(fname), re.IGNORECASE)
        for actual_file in bitstreams_dir.iterdir():
            if pattern.fullmatch(actual_file.name):
                found_files.append(actual_file.name)
                break

        # Strategy 3: Try common extensions if no extension provided
        if not Path(fname).suffix:
            for ext in Config.ALLOWED_FILE_EXTENSIONS:
                candidate = bitstreams_dir / (fname + ext)
                if candidate.exists():
                    found_files.append(candidate.name)
                    break

    return found_files

# ---- Validation for critical metadata fields ----
def validate_required_fields(columns, row_values):
    """Check for required DSpace fields"""
    missing = []

    for element in Config.REQUIRED_FIELDS:
        found = False
        for col, val in zip(columns, row_values):
            if is_blank(val):
                continue

            col_base = base_header(norm_header(col))
            if not is_dc_column(col_base):
                continue

            element_name, qualifier = parse_dc(col_base)
            if element_name == element.split('.')[0]:
                if '.' in element:
                    if qualifier == element.split('.')[1]:
                        found = True
                        break
                else:
                    found = True
                    break

        if not found:
            missing.append(element)

    return missing

# ---- XML writing ----

def write_dc_xml(columns, row_values, output_path: Path):
   root = ET.Element("dublin_core", schema="dc")
   seen = set()  # dedupe exact triples

   for raw_col, val in zip(columns, row_values):
       if is_blank(val) or raw_col is None:
           continue

       raw_col = norm_header(raw_col)
       base_col = base_header(raw_col)

       if not is_dc_column(base_col):
           continue

       element, qualifier = parse_dc(base_col)
       if element is None:
           continue

       text = str(val).strip()
       if not text:
           continue

       trip = (element, qualifier or "none", text)
       if trip in seen:
           continue
       seen.add(trip)

       attrs = {
           "element": element,
           "qualifier": qualifier if qualifier else "none",
           "language": Config.LANG_ATTR
       }
       attrs = {k: v for k, v in attrs.items() if v is not None}

       dcvalue = ET.SubElement(root, "dcvalue", attrs)
       dcvalue.text = text

   # Pretty-print and enforce closing tag
   if len(root) == 0:
       root.text = "\n"
   else:
       root.text = "\n  "
       for child in root:
           if child.tail is None or not child.tail.strip():
               child.tail = "\n  "
       root[-1].tail = "\n"

   tree = ET.ElementTree(root)
   try:
       ET.indent(tree, space="  ", level=0)
   except Exception:
       pass

   # Write UTF-8, no BOM
   tree.write(output_path, encoding="utf-8", xml_declaration=True)

   # Ensure trailing newline only
   with open(output_path, "r+", encoding="utf-8") as f:
       content = f.read()
       if not content.endswith("\n"):
           content += "\n"
           f.seek(0)
           f.write(content)
           f.truncate()

   console.info(f"Created XML with {len(root)} elements: {output_path}")



# ---- main ----
def main():
    if not Path(Config.INPUT_XLSX).exists():
        console.error(f"Input spreadsheet not found: {Config.INPUT_XLSX}")
        return

    console.info(f"Starting processing of {Config.INPUT_XLSX}")
    console.info(f"Looking for header row...")

    # find header row & load
    header_idx = find_header_row(Config.INPUT_XLSX, max_scan=Config.MAX_HEADER_SCAN)
    console.info(f"Using header row index: {header_idx}")

    df = pd.read_excel(Config.INPUT_XLSX, header=header_idx, dtype=str)
    console.info(f"Loaded DataFrame with shape: {df.shape}")

    # drop 'Unnamed' columns, trim header names
    cleaned_cols = []
    for c in df.columns:
        cname = norm_header(str(c))
        if "unnamed" in cname.lower():
            cleaned_cols.append(None)
        else:
            cleaned_cols.append(cname)
    df.columns = cleaned_cols

    # Filter out None columns
    df = df.loc[:, [c for c in df.columns if c is not None]]
    console.info(f"After cleaning, DataFrame shape: {df.shape}")
    console.info(f"Columns: {list(df.columns)}")

    # make output dir
    safe_mkdir(Config.OUTPUT_DIR)

    # locate filename columns
    filename_cols = detect_filename_columns(df.columns)
    if filename_cols:
        console.info(f"Filename column(s) detected: {', '.join(filename_cols)}")
    else:
        console.info("No filename column detected; will create XML only (contents will be empty).")

    total_records = 0
    total_files_copied = 0
    total_missing_files = 0
    error_count = 0

    for idx, row in df.iterrows():
        total_records += 1
        console.info(f"Processing record {total_records}...")

        try:
            # Validate required fields (but don't skip if missing, just warn)
            missing_fields = validate_required_fields(df.columns, row.values)
            if missing_fields:
                logging.warning(f"Record {idx+1} missing required fields: {', '.join(missing_fields)}")

            item_dir = Config.OUTPUT_DIR / str(total_records)
            safe_mkdir(item_dir)
            logging.info(f"Created folder → {item_dir}")

            # write dublin_core.xml (from dc.* columns only)
            write_dc_xml(df.columns, row.values, item_dir / "dublin_core.xml")

            # gather filenames from any detected filename-like column(s)
            filenames = set()
            for col in filename_cols:
                raw = row.get(col, "")
                if not is_blank(raw):
                    # Split on various common separators
                    parts = re.split(r"[;\n|,]+", str(raw))  # Handles semicolons, newlines, pipes, commas, back slash
                    for p in parts:
                        p = basename_only(p.strip())
                        if p and p.lower() not in ['', 'none', 'nan']:  # Additional filtering
                            filenames.add(p)

            console.info(f"Looking for files: {list(filenames)}")

            # Enhanced file discovery
            found_files = find_files_for_item(filenames, Config.FILES_DIR)
            missing_files = filenames - set(found_files)

            # write contents and copy files
            copied = []
            for fname in found_files:
                src = Config.FILES_DIR / fname
                dst = item_dir / fname
                try:
                    shutil.copy2(src, dst)
                    copied.append(fname)
                    total_files_copied += 1
                    console.info(f"Copied file: {fname}")
                except Exception as e:
                    logging.error(f"Copy error {src} -> {dst}: {e}")
                    console.error(f"Error copying {fname}: {e}")

            # Log missing files
            for fname in missing_files:
                total_missing_files += 1
                logging.warning(f"Missing file → {fname}")
                console.warning(f"Missing file: {fname}")

            # Write contents file
            with open(item_dir / "contents", "w", encoding="utf-8") as cf:
                for fname in sorted(copied):
                    cf.write(fname + "\n")

            console.info(f"Completed record {total_records}: {len(copied)} files copied")

        except Exception as e:
            logging.error(f"Error processing record {idx+1}: {str(e)}", exc_info=True)
            console.error(f"Failed to process record {idx+1}: {e}")
            error_count += 1

    # Generate summary report
    console.info("=" * 50)
    console.info("PROCESSING SUMMARY REPORT")
    console.info("=" * 50)
    console.info(f"Total records processed: {total_records}")
    console.info(f"Records with successfully created packages: {total_records - error_count}")
    console.info(f"Records with errors: {error_count}")
    console.info(f"Files successfully copied: {total_files_copied}")
    console.info(f"Missing files: {total_missing_files}")
    console.info(f"Output directory: {Config.OUTPUT_DIR.resolve()}")
    console.info("=" * 50)

    console.info(f"Processing complete: {total_records} records processed.")
    console.info(f"Files copied: {total_files_copied}")
    console.info(f"Missing files: {total_missing_files} (see {Config.LOG_FILE})")

if __name__ == "__main__":
    if not Config.FILES_DIR.exists():
        logging.warning(f"Files directory not found: {Config.FILES_DIR.resolve()} (will still create records)")
        console.warning(f"Files directory not found: {Config.FILES_DIR.resolve()}")
    main()

