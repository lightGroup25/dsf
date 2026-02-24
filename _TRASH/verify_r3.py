
import openpyxl
import os
import sys

def verify_r3(filename):
    print(f"Verifying {filename}...")
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
    except Exception as e:
        print(f"Error: {e}")
        return

    sheet_name = None
    for name in wb.sheetnames:
        if "R3" in name:
            sheet_name = name
            break
    
    if not sheet_name:
        print("FAIL: Sheet R3 not found")
        return

    ws = wb[sheet_name]
    print(f"Sheet: {sheet_name}")

    checks = [
        ("A4", "GULFCAM SAS"),
        ("A5", "M050900027774W"),
        ("A11", "NYODOG"),
        ("B11", "PERRIAL JEAN"),
        ("C11", "PRESIDENT"),
        ("A12", "BOU"),
        ("B26", "MBAYEN"), # Row 26 (Index 0)
        ("C26", "RENE"),
        ("D26", "PRESIDENT CONSEIL SURVEILLANCE"),
        ("B27", "NYODOG"), # Row 27 (Index 1)
        ("B33", "NGO MBAYEN"), # Row 33 (Index 5)
        ("D33", "MEMBRE")
    ]

    all_passed = True
    for coord, expected in checks:
        val = ws[coord].value
        passed = val and expected.upper() in str(val).upper()
        status = "PASS" if passed else "FAIL"
        print(f"[{status}] {coord}: {val} (Expected: {expected})")
        if not passed:
            all_passed = False

    if all_passed:
        print("SUCCESS: Fiche R3 is correctly filled.")
    else:
        print("FAILURE: Some checks failed.")
        sys.exit(1)

if __name__ == "__main__":
    # Use the specific file generated
    filename = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_OUTPUT_2024_20260213_183235.xlsx"
    verify_r3(filename)
