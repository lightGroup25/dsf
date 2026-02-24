
import openpyxl
import os
import sys

def verify_universal_headers(filename):
    print(f"Verifying {filename}...")
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
    except Exception as e:
        print(f"Error: {e}")
        return

    # Sample a few random sheets to check if headers are filled
    # "Fiche R2", "Fiche R3", "NOTE 13" were manually mapped, so they should be correct regardless (since explicit overwrites implicit).
    # But let's check a sheet NOT in the manual mapping, e.g., "NOTE 1" or "NOTE 2" if they exist and have the header.
    
    target_sheets = ["NOTE 1", "NOTE 2A", "NOTE 30"] 
    # Adjust based on what sheets actually exist.
    
    found_any = False
    
    for sheet_name in wb.sheetnames:
        # Just check a few sheets that look like Notes
        if "NOTE" in sheet_name.upper():
            ws = wb[sheet_name]
            # Check A3/A4/A5 for the patterns
            
            header_found = False
            for r in range(1, 10):
                cell = ws.cell(row=r, column=1)
                val = str(cell.value)
                if "GULFCAM SAS" in val and "31-12-2024" in val:
                     print(f"[PASS] {sheet_name} Header 1 found at {cell.coordinate}")
                     header_found = True
                
                if "M050900027774W" in val and "12" in val: # Duration is 12
                     print(f"[PASS] {sheet_name} Header 2 found at {cell.coordinate}")
                     header_found = True
            
            if header_found:
                found_any = True

    if found_any:
        print("SUCCESS: Universal headers verified in at least some sheets.")
    else:
        print("WARNING: No universal headers confirmed in sampled notes (maybe they don't have this structure?).")

if __name__ == "__main__":
    filename = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_OUTPUT_2024_20260213_184245.xlsx"
    verify_universal_headers(filename)
