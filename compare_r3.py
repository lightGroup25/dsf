
import openpyxl
import os
from openpyxl.utils import get_column_letter

def compare_sheets(template_file, output_file, sheet_name="Fiche R3"):
    """Compare the same sheet in template and output."""
    print(f"=== Comparing Sheet: {sheet_name} ===\n")
    
    try:
        wb_tpl = openpyxl.load_workbook(template_file, data_only=False)  # Don't load data only for template
        ws_tpl = wb_tpl[sheet_name] if sheet_name in wb_tpl.sheetnames else None
        
        if not ws_tpl:
            print(f"Sheet {sheet_name} not found in template")
            return
        
        print("TEMPLATE STRUCTURE (Rows 1-35):")
        print("-" * 150)
        for r in range(1, 36):
            row_data = []
            for c in range(1, 9):  # A-H
                cell = ws_tpl.cell(row=r, column=c)
                val = cell.value
                merged = any(cell.coordinate in merged_range for merged_range in ws_tpl.merged_cells.ranges)
                marker = "[MERGED]" if merged else ""
                row_data.append(f"{get_column_letter(c)}{r}:{val or ''}{marker}")
            
            # Only print if row has content
            if any(row_data):
                print(f"Row {r:2d}: " + " | ".join(row_data[:6]))  # Show first 6 columns
        
        print("\n" + "="*150)
        print("OUTPUT FILE STRUCTURE (Rows 1-35):")
        print("-" * 150)
        
        wb_out = openpyxl.load_workbook(output_file, data_only=True)
        ws_out = wb_out[sheet_name] if sheet_name in wb_out.sheetnames else None
        
        if not ws_out:
            print(f"Sheet {sheet_name} not found in output")
            return
        
        for r in range(1, 36):
            row_data = []
            for c in range(1, 9):  # A-H
                cell = ws_out.cell(row=r, column=c)
                val = cell.value
                row_data.append(f"{get_column_letter(c)}{r}:{val or ''}")
            
            if any(row_data):
                print(f"Row {r:2d}: " + " | ".join(row_data[:6]))
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    # Use the prefilled template (before final write) as comparison
    template_file = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\dsf_prefilled_template_20260213_190131.xlsx"
    output_file = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_OUTPUT_2024_20260213_190157.xlsx"
    
    if not os.path.exists(template_file):
        print(f"Template file not found: {template_file}")
    else:
        compare_sheets(template_file, output_file)
