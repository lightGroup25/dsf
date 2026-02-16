"""
Check what was actually filled in the output file and verify if it's in oranged zones
"""

import openpyxl

output_file = r'output/DSF_OPTION3_COMPLETE_WITH_NOTES.xlsx'
template_file = r'templates/DSF Normal standard.xlsx'

print("Checking what was filled in BILAN PAYSAGE...\n")

wb_output = openpyxl.load_workbook(output_file)
wb_template = openpyxl.load_workbook(template_file)

ws_output = wb_output['BILAN PAYSAGE']
ws_template = wb_template['BILAN PAYSAGE']

filled_cells = []

# Check rows 12-40 (data rows)
for row_idx in range(12, 40):
    for col in ['D', 'E', 'F', 'G']:
        output_val = ws_output[f'{col}{row_idx}'].value
        template_val = ws_template[f'{col}{row_idx}'].value
        col_b = ws_output[f'B{row_idx}'].value
        
        # If something was filled (value in output where template was empty)
        if output_val and not template_val:
            filled_cells.append({
                'cell': f'{col}{row_idx}',
                'column': col,
                'row_label': str(col_b)[:50] if col_b else '',
                'value': output_val if not isinstance(output_val, str) else output_val[:30]
            })

print(f"Found {len(filled_cells)} filled cells in BILAN PAYSAGE\n")
for cell_info in filled_cells[:20]:
    print(f"{cell_info['cell']} (Col{cell_info['column']}): {cell_info['row_label']}")
    print(f"  Value: {cell_info['value']}\n")

# Analyze which columns are being filled
columns_filled = {}
for cell in filled_cells:
    col = cell['column']
    if col not in columns_filled:
        columns_filled[col] = 0
    columns_filled[col] += 1

print(f"\nColumns that were filled:")
for col, count in sorted(columns_filled.items()):
    print(f"  Column {col}: {count} cells")
