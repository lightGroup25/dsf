"""
Inspect cell colors in DSF template to identify oranged zones (should not be filled)
"""

import openpyxl
from openpyxl.styles import PatternFill

dsf_file = r'templates/DSF Normal standard.xlsx'
wb = openpyxl.load_workbook(dsf_file)

# Check main sheets
sheets_to_check = ['BILAN PAYSAGE', 'NOTE 3A', 'NOTE 4', 'NOTE 6']

for sheet_name in sheets_to_check:
    if sheet_name not in wb.sheetnames:
        continue
    
    ws = wb[sheet_name]
    print(f"\n{'='*70}")
    print(f"Sheet: {sheet_name}")
    print(f"{'='*70}")
    
    colored_cells = {'orange': [], 'yellow': [], 'red': [], 'green': [], 'blue': [], 'gray': []}
    
    for row_idx in range(1, min(100, ws.max_row + 1)):
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            cell = ws[f'{col}{row_idx}']
            
            if cell.fill and cell.fill.fgColor:
                color_type = None
                color_value = str(cell.fill.fgColor.rgb) if cell.fill.fgColor.rgb else str(cell.fill.fgColor.theme)
                
                # Detect color
                if 'FF' in str(color_value):  # Has color code
                    rgb_str = str(color_value).upper()
                    
                    # Orange: FFA500, FFB366, FFC266, etc
                    if 'FF' in rgb_str and ('A5' in rgb_str or 'FF' in rgb_str[:2]):
                        if 'B3' in rgb_str or 'C2' in rgb_str or 'A5' in rgb_str:  # Orange range
                            color_type = 'orange'
                    # Yellow: FFFF00, FFFF99, etc
                    elif rgb_str.startswith('FFFF'):
                        color_type = 'yellow'
                    # Red: FF0000, FF6666, etc
                    elif 'FF0000' in rgb_str or 'FF6666' in rgb_str:
                        color_type = 'red'
                    # Green: 00B050, 92D050
                    elif '00B0' in rgb_str or '92D0' in rgb_str:
                        color_type = 'green'
                    # Gray: CCCCCC, D3D3D3
                    elif 'CC' in rgb_str or 'D3' in rgb_str or 'E0' in rgb_str:
                        color_type = 'gray'
                    # Blue: 0070C0, 4472C4
                    elif '0070' in rgb_str or '4472' in rgb_str:
                        color_type = 'blue'
                
                if color_type and cell.value:
                    colored_cells[color_type].append({
                        'cell': f'{col}{row_idx}',
                        'value': str(cell.value)[:50],
                        'color': color_value
                    })
    
    for color, cells in colored_cells.items():
        if cells:
            print(f"\n{color.upper()} cells ({len(cells)}):")
            for cell_info in cells[:10]:
                print(f"  {cell_info['cell']}: {cell_info['value']}")
            if len(cells) > 10:
                print(f"  ... and {len(cells)-10} more")
