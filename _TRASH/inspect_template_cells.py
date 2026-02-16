
import openpyxl

def inspect_sheet(wb, sheet_name, rows=20, cols=15):
    if sheet_name not in wb.sheetnames:
        print(f"Sheet '{sheet_name}' not found in {wb.sheetnames}")
        # Try to find a partial match
        matching = [s for s in wb.sheetnames if sheet_name.strip() in s]
        if matching:
            print(f"  Did you mean: {matching}?")
            sheet_name = matching[0] # Try the first match
        else:
            return

    ws = wb[sheet_name]
    print(f"\n--- Inspecting {sheet_name} ---")
    
    # Print Merged Cells (first 5 relevant)
    count = 0
    print("Merged Cells (Top 5 relevant):")
    sorted_merged = sorted(ws.merged_cells.ranges, key=lambda x: (x.min_row, x.min_col))
    for merged in sorted_merged:
        if merged.min_row <= rows and merged.min_col <= cols:
            print(f"  {merged}")
            count += 1
            if count >= 5: break

    # Print Cell Values
    print("\nCell Values:")
    for r in range(1, rows + 1):
        row_data = []
        for c in range(1, cols + 1):
            cell = ws.cell(row=r, column=c)
            coord = cell.coordinate
            val = str(cell.value).strip() if cell.value is not None else ""
            
            # Check if this cell is part of a merge
            is_master = True
            for m in ws.merged_cells.ranges:
                if coord in m:
                    if coord != m.start_cell.coordinate:
                        is_master = False
                        val = f"[->{m.start_cell.coordinate}]"
                    break
            
            if len(val) > 15: val = val[:12] + "..."
            row_data.append(f"{val}")
        print(f"Row {r:2}: " + " | ".join(row_data))

filename = "DSF Normal standard.xlsx"
try:
    print(f"Loading {filename}...")
    wb = openpyxl.load_workbook(filename, data_only=True)
    print("Defined Names (Sample):")
    # wb.defined_names is a DefinedNameList in newer openpyxl, or dict in older?
    # Trying iteration
    try:
        count = 0
        for name in wb.defined_names:
            print(f"  {name} -> {wb.defined_names[name].value}") # This might fail depending on version
            count += 1
            if count >= 20: break
    except:
         for name_obj in getattr(wb.defined_names, "definedName", []):
             print(f"  {name_obj.name} -> {name_obj.value}")


    inspect_sheet(wb, "ENTETE")
    inspect_sheet(wb, "Fiche R1")
    inspect_sheet(wb, "Fiche R2")
    inspect_sheet(wb, "Fiche R3")

def inspect_sheet_range(wb, sheet_name, start_row, end_row, cols=15):
    if sheet_name not in wb.sheetnames:
        print(f"Sheet {sheet_name} not found")
        return
    ws = wb[sheet_name]
    print(f"\n--- Inspecting {sheet_name} Rows {start_row}-{end_row} ---")
    for r in range(start_row, end_row + 1):
        row_data = []
        for c in range(1, cols + 1):
            cell = ws.cell(row=r, column=c)
            val = str(cell.value) if cell.value else ""
            if len(val) > 10: val = val[:10]
            row_data.append(f"{val}")
        print(f"Row {r}: " + " | ".join(row_data))

# ... existing code ...
try:
    print(f"Loading {filename}...")
    wb = openpyxl.load_workbook(filename, data_only=True)
    inspect_sheet_range(wb, "Fiche R2", 25, 35)
except Exception as e:
    print(f"Error: {e}")
