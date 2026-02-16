"""
Option 3 Corrected - Only fill DSF columns that should be filled
Avoid filling orange zones (amortissements, calculs)

DSF Structure (Bilan Paysage):
- Column D: BRUT (opening - FILL THIS)
- Column E: AMORT et DEPREC (amortissements - ORANGE ZONE, DON'T FILL)
- Column F: NET (calculated - ORANGE ZONE, DON'T FILL)
- Column G: Comparatif N-1 (ORANGE ZONE, DON'T FILL)

Note sections typically have:
- Column for opening balance (FILL)
- Column for movements/changes (FILL if specific)
- Column for closing balance (FILL or FORMULA if empty)
"""

import openpyxl
from difflib import SequenceMatcher
import json
from pathlib import Path
from datetime import datetime

BALANCE_FILE = r"input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
DSF_TEMPLATE = r"templates/DSF Normal standard.xlsx"
OUTPUT_FILE = r"output/DSF_OPTION3_CORRECTED_NO_ORANGE.xlsx"

def load_balance_data(filepath):
    """Load balance file"""
    print(f"Loading balance file: {filepath}")
    
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        
        accounts = {}
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, 
                                                    min_col=1, max_col=30, 
                                                    values_only=False), 1):
            if row_idx < 9:
                continue
                
            account_num = row[0].value
            if account_num is None:
                continue
            
            try:
                int(str(account_num).split('.')[0])
            except:
                continue
            
            account_label = row[3].value if row[3] else None
            opening_balance = row[13].value if row[13] else None
            movements = row[20].value if row[20] else None
            closing_balance = row[26].value if row[26] else None
            
            if account_label and (opening_balance or closing_balance):
                accounts[str(account_num).strip()] = {
                    'label': str(account_label).strip(),
                    'opening': opening_balance,
                    'movements': movements,
                    'closing': closing_balance
                }
        
        print(f"✓ Loaded {len(accounts)} accounts")
        return accounts
        
    except Exception as e:
        print(f"✗ Error: {e}")
        raise

def identify_fillable_cells_in_sheet(ws, sheet_name):
    """Identify exactly which cells should be filled by analyzing the sheet structure"""
    fillable_rows = []
    
    # Scan for data rows (rows that have an account reference in column A or B)
    for row_idx in range(1, min(500, ws.max_row + 1)):
        col_a = ws[f'A{row_idx}'].value
        col_b = ws[f'B{row_idx}'].value
        
        if col_a is None and col_b is None:
            continue
        
        col_a_str = str(col_a).strip() if col_a else ""
        col_b_str = str(col_b).strip() if col_b else ""
        
        # Skip header rows and section headers
        is_header = any(x in col_a_str.upper() for x in 
                       ['TOTAL', 'DESIGNATION', 'BILAN', 'EXERCICE', 'REF', 'NET', 'BRUT', 'AMORT'])
        is_header = is_header or any(x in col_b_str.upper() for x in 
                                     ['TOTAL', 'DESIGNATION', 'EXERCICE'])
        
        if is_header:
            continue
        
        # Determine which columns to fill based on sheet type
        if 'BILAN' in sheet_name:
            # BILAN PAYSAGE: Only fill column D (BRUT/opening)
            # Do NOT fill E (AMORT), F (NET), G (comparatif)
            fillable_cols = ['D']  # Only BRUT column
        else:
            # NOTE sections: Fill D (opening), maybe K (closing)
            fillable_cols = ['D', 'K']  # Opening and closing columns
        
        if col_a_str or col_b_str:
            fillable_rows.append({
                'row': row_idx,
                'label': col_b_str if col_b_str else col_a_str,
                'fillable_columns': fillable_cols
            })
    
    return fillable_rows

def fuzzy_match(balance_label, target_label):
    """Simple fuzzy matching"""
    if not target_label:
        return 0
    return SequenceMatcher(None, balance_label.lower(), target_label.lower()).ratio()

def process_sheet_corrected(wb_template, sheet_name, balance_accounts):
    """Process sheet, filling ONLY appropriate columns"""
    try:
        ws = wb_template[sheet_name]
    except:
        return 0
    
    fillable_rows = identify_fillable_cells_in_sheet(ws, sheet_name)
    cells_filled = 0
    
    for row_info in fillable_rows:
        row_idx = row_info['row']
        row_label = row_info['label']
        fillable_cols = row_info['fillable_columns']
        
        # Find best matching account
        best_score = 0
        best_account = None
        
        for account_data in balance_accounts.values():
            score = fuzzy_match(account_data['label'], row_label)
            if score > best_score:
                best_score = score
                best_account = account_data
        
        # Fill only if good match and only fillable columns
        if best_score >= 0.4 and best_account:
            for col in fillable_cols:
                try:
                    cell = ws[f'{col}{row_idx}']
                    
                    if col == 'D':  # Opening/BRUT column
                        if best_account['opening'] and not cell.value:
                            cell.value = best_account['opening']
                            cells_filled += 1
                    
                    elif col == 'K':  # Closing/NET column  
                        if best_account['closing'] and not cell.value:
                            cell.value = best_account['closing']
                            cells_filled += 1
                
                except:
                    pass
    
    return cells_filled

def main():
    print("="*70)
    print("OPTION 3 CORRECTED - ONLY FILLABLE COLUMNS (NO ORANGE ZONES)")
    print("="*70)
    
    try:
        # Load data
        balance_accounts = load_balance_data(BALANCE_FILE)
        
        print(f"Loading DSF template: {DSF_TEMPLATE}")
        wb_template = openpyxl.load_workbook(DSF_TEMPLATE)
        print(f"✓ Template has {len(wb_template.sheetnames)} sheets")
        
        # Process each sheet
        total_cells_filled = 0
        sheets_processed = 0
        
        print(f"\nProcessing sheets (only fillable columns)...")
        
        for sheet_name in wb_template.sheetnames:
            cells_filled = process_sheet_corrected(wb_template, sheet_name, balance_accounts)
            
            if cells_filled > 0:
                total_cells_filled += cells_filled
                sheets_processed += 1
                if 'BILAN' in sheet_name or 'NOTE' in sheet_name:
                    print(f"  ✓ {sheet_name}: {cells_filled} cells")
        
        # Save output
        print(f"\nSaving to {OUTPUT_FILE}")
        Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)
        wb_template.save(OUTPUT_FILE)
        
        print("\n" + "="*70)
        print("SUMMARY - OPTION 3 CORRECTED")
        print("="*70)
        print(f"Sheets processed: {sheets_processed}/{len(wb_template.sheetnames)}")
        print(f"Total cells filled (ONLY fillable columns): {total_cells_filled}")
        print(f"✓ Orange zones: NOT FILLED")
        print(f"\n✓ Output: {OUTPUT_FILE}")
        
    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
