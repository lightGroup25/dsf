
import openpyxl
import os
import sys
from openpyxl.utils import get_column_letter

def inspect_note13(filename):
    print(f"Loading {filename}...")
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
    except Exception as e:
        print(f"Error loading workbook: {e}")
        return

    sheet_name = None
    for name in wb.sheetnames:
        if "NOTE 13" in name.upper() or "NOTE13" in name.upper():
            sheet_name = name
            break
    
    if not sheet_name:
        print(f"Sheet containing 'NOTE 13' not found in {wb.sheetnames}")
        return

    print(f"Inspecting sheet: {sheet_name}")
    ws = wb[sheet_name]

    # Inspect first 60 rows
    for r in range(1, 60):
        row_values = []
        for c in range(1, 10):  # Inspect first A-I columns
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val:
                row_values.append(f"{cell.coordinate}:{val}")
        if row_values:
            print(f"Row {r}: {', '.join(row_values)}")

    print("\n--- Detailed Grid A8-H25 ---")
    for r in range(8, 26):
        row_vals = []
        for c in range(1, 9): # A-H
            val = ws.cell(row=r, column=c).value
            row_vals.append(f"{get_column_letter(c)}{r}: {val or ''}")
        print(" | ".join(row_vals))

if __name__ == "__main__":
    filename = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_OUTPUT_2024_20260213_183235.xlsx"
    if not os.path.exists(filename):
         # Search for latest output
         output_dir = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output"
         files = sorted([os.path.join(output_dir, f) for f in os.listdir(output_dir) if f.startswith("DSF_OUTPUT_2024") and f.endswith(".xlsx")], key=os.path.getmtime, reverse=True)
         if files:
             filename = files[0]
         else:
             filename = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\data\MODELE_O_01_DSF_S_N_SY_2024.xlsx"
    
    inspect_note13(filename)
