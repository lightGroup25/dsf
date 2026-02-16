
import openpyxl
import os

def inspect_merged_cells(filename, sheet_name="Fiche R3"):
    """Inspect merged cells in a sheet."""
    print(f"=== Merged Cells in {sheet_name} ===\n")
    
    try:
        wb = openpyxl.load_workbook(filename, data_only=False)
        ws = wb[sheet_name]
        
        print(f"Total merged cells: {len(ws.merged_cells.ranges)}\n")
        
        print("Merged ranges:")
        for merged_range in ws.merged_cells.ranges:
            print(f"  {merged_range}")
        
        # Focus on rows 25-35
        print("\n=== Merged cells in rows 25-35 ===")
        for merged_range in ws.merged_cells.ranges:
            range_str = str(merged_range)
            # Check if row is in 25-35
            if any(str(r) in range_str for r in range(25, 36)):
                print(f"  {merged_range}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    template_file = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\dsf_prefilled_template_20260213_184227.xlsx"
    output_file = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_OUTPUT_2024_20260213_184245.xlsx"
    
    print("TEMPLATE MERGED CELLS:")
    print("=" * 100)
    inspect_merged_cells(template_file)
    
    print("\n" + "=" * 100)
    print("\nOUTPUT MERGED CELLS:")
    print("=" * 100)
    inspect_merged_cells(output_file)
