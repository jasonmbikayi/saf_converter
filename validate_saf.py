#!/usr/bin/env python3
import os
import sys
import logging
from pathlib import Path
import xml.etree.ElementTree as ET

# ---- Config ----
REQUIRED_FIELDS = ["title", "creator", "date.issued"]

def validate_item(item_dir: Path, report):
    issues = []
    
    # 1. Check dublin_core.xml exists
    dc_path = item_dir / "dublin_core.xml"
    if not dc_path.exists():
        issues.append("❌ Missing dublin_core.xml")
        report.append((item_dir, issues))
        return

    # 2. Parse XML
    try:
        tree = ET.parse(dc_path)
        root = tree.getroot()
    except Exception as e:
        issues.append(f"❌ Invalid XML: {e}")
        report.append((item_dir, issues))
        return

    # 3. Check required metadata
    found_fields = {field: False for field in REQUIRED_FIELDS}
    for dcvalue in root.findall("dcvalue"):
        element = dcvalue.get("element")
        qualifier = dcvalue.get("qualifier")
        text = dcvalue.text.strip() if dcvalue.text else ""
        if not text:
            continue

        if element == "title":
            found_fields["title"] = True
        elif element == "creator":
            found_fields["creator"] = True
        elif element == "date" and qualifier == "issued":
            found_fields["date.issued"] = True

    for f, ok in found_fields.items():
        if not ok:
            issues.append(f"⚠ Missing required metadata: dc.{f}")

    # 4. Check contents file
    contents_path = item_dir / "contents"
    if not contents_path.exists():
        issues.append("❌ Missing contents file")
    else:
        with open(contents_path, "r", encoding="utf-8") as cf:
            for line in cf:
                fname = line.strip().split()[0]  # ignore bundle specifiers
                if not (item_dir / fname).exists():
                    issues.append(f"❌ Listed in contents but missing: {fname}")

    if issues:
        report.append((item_dir, issues))


def main():
    if len(sys.argv) != 2:
        print("Usage: python validate_saf.py saf_converted/")
        sys.exit(1)

    saf_dir = Path(sys.argv[1])
    if not saf_dir.is_dir():
        print(f"❌ Not a directory: {saf_dir}")
        sys.exit(1)

    report = []
    item_dirs = [p for p in saf_dir.iterdir() if p.is_dir()]
    print(f"🔍 Validating {len(item_dirs)} SAF item(s) in {saf_dir}...\n")

    for item_dir in sorted(item_dirs):
        validate_item(item_dir, report)

    # Print report
    if not report:
        print("✅ All items passed validation!")
    else:
        print("⚠ Issues found:\n")
        for item_dir, issues in report:
            print(f"Item {item_dir}:")
            for issue in issues:
                print(f"  - {issue}")
            print()

    print("Validation complete.")

if __name__ == "__main__":
    main()
