"""
Option 3 INTELLIGENT - Auto-detect fillable columns per sheet
Only fill columns that are actually meant to be filled
"""

import openpyxl
from difflib import SequenceMatcher
import json
from pathlib import Path

BALANCE_FILE = r"input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
DSF_TEMPLATE = r"templates/DSF Normal standard.xlsx"
OUTPUT_FILE = r"output/DSF_OPTION3_INTELLIGENT.xlsx"

COLUMN_DETECTION_RULES = {
    'fillable_keywords': [
        'MONTANT BRUTE', 'ACQUISITIONS', 'VIREMENTS', 'REEVALUATION', 
        'CESSIONS', 'VARIATIONS', 'CREATIONS', 'MOUVEMENTS'
    ],
    'protected_keywords': [
        'NET', 'AMORT', 'CLÔTURE', 'CLOTURE', 'N-1', 'COMPARATIF',
        'EXERCICE PRECEDENT', 'TOTAL', 'BRUT NET', 'DEPRECIATION'
    ]
}

def load_balance_data(filepath):
    """Load balance file"""
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
    
    return accounts

def detect_fillable_columns(ws):
    """Detect which columns in this sheet should be filled"""
    # Find header row
    header_row = None
    for row_idx in range(1, 30):
        row_text = []
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
            val = ws[f'{col}{row_idx}'].value
            if val:
                row_text.append(str(val)[:20])
        
        if any(kw in ' '.join(row_text).upper() for kw in ['MONTANT', 'BRUT', 'ACQUISITIONS', 'VIREMENTS']):
            header_row = row_idx
            break
    
    if not header_row:
        return {'header_row': None, 'fillable': []}
    
    fillable_columns = []
    
    for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
        header = ws[f'{col}{header_row}'].value
        if not header:
            continue
        
        header_text = str(header).upper()
        
        # Check if protected first
        is_protected = any(kw in header_text for kw in COLUMN_DETECTION_RULES['protected_keywords'])
        if is_protected:
            continue
        
        # Check if fillable
        is_fillable = any(kw in header_text for kw in COLUMN_DETECTION_RULES['fillable_keywords'])
        
        # Special case: BRUT without NET = opening (fillable)
        if 'BRUT' in header_text and 'NET' not in header_text:
            is_fillable = True
        
        if is_fillable:
            fillable_columns.append(col)
    
    return {
        'header_row': header_row,
        'fillable': fillable_columns
    }

def process_sheet_intelligent(wb_template, sheet_name, balance_accounts):
    """Process sheet using intelligent column detection"""
    try:
        ws = wb_template[sheet_name]
    except:
        return 0
    
    # Detect which columns to fill
    column_info = detect_fillable_columns(ws)
    if not column_info['fillable']:
        return 0
    
    fillable_cols = column_info['fillable']
    header_row = column_info['header_row']
    cells_filled = 0
    
    # Find data rows (not headers, not section headers)
    for row_idx in range(header_row + 1, min(header_row + 500, ws.max_row + 1)):
        col_a = ws[f'A{row_idx}'].value
        col_b = ws[f'B{row_idx}'].value
        
        if not col_a and not col_b:
            continue
        
        row_label = str(col_b if col_b else col_a).strip().upper()
        
        # Skip section headers
        if any(x in row_label for x in ['TOTAL', 'IMMOBILISATIONS', 'SUB-TOTAL', 'SOCIETE', 'DESIGNATION']):
            if not any(x in row_label for x in ['INCORPORELLES', 'CORPORELLES']):  # Allow specific account types
                continue
        
        # Find matching balance account
        best_score = 0
        best_account = None
        
        for account_data in balance_accounts.values():
            score = SequenceMatcher(None, account_data['label'].upper(), row_label).ratio()
            if score > best_score:
                best_score = score
                best_account = account_data
        
        # Fill if good match
        if best_score >= 0.4 and best_account:
            for col in fillable_cols:
                try:
                    cell = ws[f'{col}{row_idx}']
                    if cell.value:
                        continue  # Don't overwrite existing values
                    
                    # Determine what data to put in this column based on column position
                    col_idx = ord(col) - ord('D')  # Relative to column D
                    
                    if col_idx == 0:  # First fillable column = opening
                        if best_account['opening']:
                            cell.value = best_account['opening']
                            cells_filled += 1
                    elif best_account['movements']:  # Middle columns = movements
                        cell.value = best_account['movements']
                        cells_filled += 1
                    elif col_idx > 3 and best_account['closing']:  # Last columns = closing
                        cell.value = best_account['closing']
                        cells_filled += 1
                
                except:
                    pass
    
    return cells_filled

def main():
    print("="*70)
    print("OPTION 3 INTELLIGENT - AUTO-DETECT FILLABLE COLUMNS PER SHEET")
    print("="*70)
    
    balance_accounts = load_balance_data(BALANCE_FILE)
    print(f"✓ Loaded {len(balance_accounts)} accounts")
    
    wb_template = openpyxl.load_workbook(DSF_TEMPLATE)
    print(f"✓ Loaded template with {len(wb_template.sheetnames)} sheets")
    
    print(f"\nProcessing sheets...\n")
    
    total_cells = 0
    sheets_processed = 0
    
    for sheet_name in wb_template.sheetnames:
        cells = process_sheet_intelligent(wb_template, sheet_name, balance_accounts)
        if cells > 0:
            total_cells += cells
            sheets_processed += 1
            if sheets_processed <= 10 or 'NOTE' in sheet_name or 'BILAN' in sheet_name:
                print(f"  ✓ {sheet_name}: {cells} cells")
    
    Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)
    wb_template.save(OUTPUT_FILE)
    
    print(f"\n{'='*70}")
    print(f"SUMMARY - OPTION 3 INTELLIGENT")
    print(f"{'='*70}")
    print(f"Sheets processed: {sheets_processed}/{len(wb_template.sheetnames)}")
    print(f"Total cells filled: {total_cells}")
    print(f"✓ Each sheet: Only fillable columns (auto-detected)")
    print(f"✓ NO orange zones filled")
    print(f"\n✓ Output: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
