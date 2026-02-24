"""
Analyze the ACTUAL structure of each sheet to see which columns
actually exist and should be filled
"""

import openpyxl

dsf_file = r'templates/DSF Normal standard.xlsx'
wb = openpyxl.load_workbook(dsf_file)

sheets_to_check = ['BILAN PAYSAGE', 'NOTE 3A', 'NOTE 4', 'NOTE 6', 'NOTE 7', 'NOTE 13']

for sheet_name in sheets_to_check:
    if sheet_name not in wb.sheetnames:
        continue
    
    ws = wb[sheet_name]
    print(f"\n{'='*70}")
    print(f"Sheet: {sheet_name}")
    print(f"{'='*70}")
    
    # Scan first 50 rows to find headers and data structure
    print("\nFirst 20 rows (showing non-empty cells):")
    
    for row_idx in range(1, 21):
        row_content = []
        has_content = False
        
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
            cell = ws[f'{col}{row_idx}']
            val = cell.value
            
            if val:
                has_content = True
                if isinstance(val, (int, float)):
                    row_content.append(f"{col}:{val}")
                else:
                    val_str = str(val)[:20]
                    row_content.append(f"{col}:{val_str}")
        
        if has_content:
            print(f"  Row {row_idx}: {', '.join(row_content)}")
    
    # Find actual data rows (with account structure)
    print(f"\nAnalyzing data rows (rows 15-30):")
    for row_idx in range(15, 31):
        col_a = ws[f'A{row_idx}'].value
        col_b = ws[f'B{row_idx}'].value
        
        if col_a or col_b:
            row_content = f"Row {row_idx}: "
            
            # Check all columns
            non_empty_cols = []
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                cell = ws[f'{col}{row_idx}']
                if cell.value:
                    non_empty_cols.append(col)
            
            if non_empty_cols:
                row_content += f"Cols with data/headers: {non_empty_cols}"
                print(f"  {row_content}")
