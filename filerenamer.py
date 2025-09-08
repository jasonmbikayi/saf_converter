#!/usr/bin/env python3
import sys
import os
from datetime import datetime
from filename_utils import clean_filename

def main():
    if len(sys.argv) != 2:
        print("Usage: python filerenamer.py bitstreams_dir")
        sys.exit(1)

    bitstreams_dir = sys.argv[1]

    if not os.path.isdir(bitstreams_dir):
        print(f"❌ Error: {bitstreams_dir} is not a valid directory")
        sys.exit(1)

    log_file = os.path.join(bitstreams_dir, "bitstreams_dir_cleanup.log")
    log_entries = []

    for root, dirs, files in os.walk(bitstreams_dir):
        for filename in files:
            if filename.startswith("."):
                continue  # skip hidden/system files

            original, cleaned = clean_filename(filename)
            if original != cleaned:
                old_path = os.path.join(root, original)
                new_path = os.path.join(root, cleaned)

                # Handle duplicates
                counter = 1
                base, ext = os.path.splitext(new_path)
                while os.path.exists(new_path):
                    new_path = f"{base}_{counter}{ext}"
                    counter += 1

                try:
                    os.rename(old_path, new_path)
                    log_entries.append(f"{old_path}  -->  {new_path}")
                except Exception as e:
                    log_entries.append(f"❌ Failed: {old_path} ({e})")

    # Write log
    with open(log_file, "w", encoding="utf-8") as f:
        f.write(f"bitstreams_dir cleanup log - {datetime.now()}\n")
        f.write("=" * 80 + "\n")
        for entry in log_entries:
            f.write(entry + "\n")
        f.write(f"\nTotal files cleaned: {len(log_entries)}\n")

    print(f"✅ Cleanup done! Log saved at: {log_file}")

if __name__ == "__main__":
    main()
# Exambank/ExamsProcess/filerenamer.py  