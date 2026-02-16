"""
Direct Column Mapper - Option 3 (Intelligent)
Maps balance file columns to DSF with auto-calculating formulas
Col 14 (opening) → DSF Col 4
Col 21 (movements) → DSF Col 5  
DSF Col 10 → FORMULA = Col 4 + Col 5 (auto-calculate closing)
"""

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from difflib import SequenceMatcher
import json
from pathlib import Path
from datetime import datetime

# Configuration
BALANCE_FILE = r"input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
DSF_TEMPLATE = r"templates/DSF Normal standard.xlsx"
OUTPUT_FILE = r"output/DSF_OPTION3_AUTO_CALCULATED.xlsx"

def load_balance_data(filepath):
    """Load balance file and extract account data"""
    print(f"Loading balance file: {filepath}")
    
    try:
        # Load Excel with openpyxl to inspect structure
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        
        accounts = {}
        
        # Find data range
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, 
                                                    min_col=1, max_col=30, 
                                                    values_only=False), 1):
            # Skip header rows
            if row_idx < 9:
                continue
                
            # Extract account data
            account_num_cell = row[0]  # Col A (Col 1)
            account_num = account_num_cell.value
            
            # Check if it's a valid account number (numeric or starts with number)
            if account_num is None:
                continue
            
            try:
                int(str(account_num).split('.')[0])
            except:
                continue
            
            account_label = row[3].value if row[3] else None  # Col D (Col 4)
            opening_balance = row[13].value if row[13] else None  # Col N (Col 14)
            movements = row[20].value if row[20] else None  # Col U (Col 21)
            closing_balance = row[26].value if row[26] else None  # Col AA (Col 27)
            
            # Only add if we have key data
            if account_label and (opening_balance or closing_balance):
                accounts[str(account_num).strip()] = {
                    'row': row_idx,
                    'label': str(account_label).strip(),
                    'opening': opening_balance,
                    'movements': movements,
                    'closing': closing_balance
                }
        
        print(f"✓ Loaded {len(accounts)} accounts from balance file")
        return accounts
        
    except Exception as e:
        print(f"✗ Error loading balance file: {e}")
        raise

def load_dsf_template(filepath):
    """Load DSF template and identify data rows"""
    print(f"Loading DSF template: {filepath}")
    
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        
        # Find account rows in DSF (typically start after headers)
        dsf_accounts = {}
        
        for row_idx in range(1, ws.max_row + 1):
            col_a = ws[f'A{row_idx}'].value  # Account number or label
            col_b = ws[f'B{row_idx}'].value  # Often account label
            
            if col_a is None or col_b is None:
                continue
            
            col_a_str = str(col_a).strip()
            col_b_str = str(col_b).strip()
            
            # Look for account patterns
            if col_a_str and not col_a_str.startswith('NOTE') and len(col_a_str) < 50:
                # This might be an account row
                dsf_accounts[row_idx] = {
                    'col_a': col_a_str,
                    'col_b': col_b_str,
                    'account_label': col_b_str if col_b_str else col_a_str
                }
        
        print(f"✓ Identified {len(dsf_accounts)} potential data rows in DSF")
        return wb, ws, dsf_accounts
        
    except Exception as e:
        print(f"✗ Error loading DSF template: {e}")
        raise

def fuzzy_match_accounts(balance_accounts, dsf_accounts, threshold=0.5):
    """Match balance accounts to DSF rows using fuzzy matching
    
    Strategy:
    1. First try to match main accounts (10, 11, 12, etc.) to DSF
    2. For sub-accounts, try to match to same main account row
    3. Use fuzzy matching as fallback
    """
    print("\nMatching accounts between balance and DSF...")
    
    matches = {}
    matched_dsf_rows = set()  # Track which DSF rows have been matched
    
    # Separate main accounts from sub-accounts
    main_accounts = {k: v for k, v in balance_accounts.items() 
                     if len(str(k).split('.')[0]) <= 2}
    sub_accounts = {k: v for k, v in balance_accounts.items() 
                    if len(str(k).split('.')[0]) > 2}
    
    # First pass: Match main accounts with stricter threshold
    for bal_num, bal_data in main_accounts.items():
        bal_label = bal_data['label'].lower()
        best_score = 0
        best_dsf_row = None
        
        for dsf_row, dsf_data in dsf_accounts.items():
            if dsf_row in matched_dsf_rows:
                continue
            dsf_label = dsf_data['account_label'].lower()
            score = SequenceMatcher(None, bal_label, dsf_label).ratio()
            
            if score > best_score:
                best_score = score
                best_dsf_row = dsf_row
        
        if best_score >= threshold:
            matches[bal_num] = {
                'balance_label': bal_data['label'],
                'dsf_row': best_dsf_row,
                'dsf_label': dsf_accounts[best_dsf_row]['account_label'],
                'score': best_score,
                'opening': bal_data['opening'],
                'movements': bal_data['movements'],
                'closing': bal_data['closing']
            }
            matched_dsf_rows.add(best_dsf_row)
    
    # Second pass: Match remaining accounts (with more flexibility)
    for bal_num, bal_data in balance_accounts.items():
        if bal_num in matches:
            continue
            
        bal_label = bal_data['label'].lower()
        best_score = 0
        best_dsf_row = None
        
        for dsf_row, dsf_data in dsf_accounts.items():
            dsf_label = dsf_data['account_label'].lower()
            score = SequenceMatcher(None, bal_label, dsf_label).ratio()
            
            if score > best_score:
                best_score = score
                best_dsf_row = dsf_row
        
        if best_score >= 0.4:  # Lower threshold for additional matches
            matches[bal_num] = {
                'balance_label': bal_data['label'],
                'dsf_row': best_dsf_row,
                'dsf_label': dsf_accounts[best_dsf_row]['account_label'],
                'score': best_score,
                'opening': bal_data['opening'],
                'movements': bal_data['movements'],
                'closing': bal_data['closing']
            }
    
    print(f"✓ Matched {len(matches)} accounts")
    print(f"  Main accounts (threshold={threshold}): {len([m for m in main_accounts if m in matches])}")
    print(f"  Sub-accounts (threshold=0.4): {len([m for m in sub_accounts if m in matches])}")
    return matches

def apply_mappings_option3(wb, ws, matches):
    """Apply Option 3 mapping: direct columns + auto-calculating formula
    
    Handle merged cells by unmerging them first
    """
    print("\nApplying Option 3 mapping (direct columns + formulas)...")
    
    stats = {
        'total_matches': len(matches),
        'cells_filled_direct': 0,
        'cells_with_formulas': 0,
        'accounts_processed': 0,
        'merged_cells_handled': 0
    }
    
    # Collect all merged cell ranges to unmerge
    merged_ranges = list(ws.merged_cells.ranges)
    
    for bal_num, match_data in matches.items():
        dsf_row = match_data['dsf_row']
        
        try:
            # Unmerge cells if necessary
            for merged_range in merged_ranges:
                if dsf_row in merged_range.bounds[1:]:  # Check if row is in merged range
                    ws.unmerge_cells(str(merged_range))
                    stats['merged_cells_handled'] += 1
            
            # Column D (4): opening balance
            cell_d = ws[f'D{dsf_row}']
            if match_data['opening'] is not None:
                cell_d.value = match_data['opening']
                stats['cells_filled_direct'] += 1
            
            # Column E (5): movements total
            cell_e = ws[f'E{dsf_row}']
            if match_data['movements'] is not None:
                cell_e.value = match_data['movements']
                stats['cells_filled_direct'] += 1
            
            # Column J (10): auto-calculated closing = opening + movements
            cell_j = ws[f'J{dsf_row}']
            # Set formula: =D{row} + E{row}
            cell_j.value = f"=D{dsf_row}+E{dsf_row}"
            stats['cells_with_formulas'] += 1
            
            stats['accounts_processed'] += 1
            
        except Exception as e:
            print(f"  ✗ Error processing account {bal_num}: {e}")
    
    print(f"✓ Applied mappings:")
    print(f"  - Direct cell fills: {stats['cells_filled_direct']}")
    print(f"  - Formula cells: {stats['cells_with_formulas']}")
    print(f"  - Accounts processed: {stats['accounts_processed']}")
    print(f"  - Merged cells handled: {stats['merged_cells_handled']}")
    
    return stats

def save_output(wb, output_path):
    """Save modified workbook"""
    print(f"\nSaving output to: {output_path}")
    
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    
    print(f"✓ Output file saved")

def generate_report(matches, stats):
    """Generate validation report"""
    report = {
        'timestamp': datetime.now().isoformat(),
        'strategy': 'Option 3 - Direct Column Mapping with Auto-Formulas',
        'description': 'Col 4←Balance Col 14(opening), Col 5←Balance Col 21(movements), Col 10←Formula(=D+E)',
        'data': {
            'total_accounts_in_balance': len(matches),
            'accounts_matched_to_dsf': stats['accounts_processed'],
            'cells_filled_direct': stats['cells_filled_direct'],
            'cells_with_formulas': stats['cells_with_formulas'],
            'total_cells_impacted': stats['cells_filled_direct'] + stats['cells_with_formulas']
        },
        'validation': {
            'success_rate': (stats['accounts_processed'] / len(matches) * 100) if matches else 0,
            'average_match_score': sum(m['score'] for m in matches.values()) / len(matches) if matches else 0
        },
        'sample_mappings': {
            k: {
                'balance_label': v['balance_label'],
                'dsf_row': v['dsf_row'],
                'opening_balance': v['opening'],
                'movements': v['movements'],
                'match_score': v['score']
            }
            for k, v in list(matches.items())[:5]
        }
    }
    
    return report

def main():
    print("="*70)
    print("DIRECT COLUMN MAPPER - OPTION 3 (Intelligent with Auto-Formulas)")
    print("="*70)
    
    try:
        # Load data
        balance_accounts = load_balance_data(BALANCE_FILE)
        wb, ws, dsf_accounts = load_dsf_template(DSF_TEMPLATE)
        
        # Match accounts
        matches = fuzzy_match_accounts(balance_accounts, dsf_accounts, threshold=0.6)
        
        # Apply Option 3 mapping
        stats = apply_mappings_option3(wb, ws, matches)
        
        # Save output
        save_output(wb, OUTPUT_FILE)
        
        # Generate report
        report = generate_report(matches, stats)
        
        # Save report
        report_path = r"output/reports/option3_report.json"
        Path(report_path).parent.mkdir(parents=True, exist_ok=True)
        with open(report_path, 'w') as f:
            json.dump(report, f, indent=2, default=str)
        
        print(f"\n✓ Report saved to: {report_path}")
        
        # Print summary
        print("\n" + "="*70)
        print("SUMMARY - OPTION 3 RESULTS")
        print("="*70)
        print(f"Accounts matched: {stats['accounts_processed']}/{len(balance_accounts)}")
        print(f"Cells filled (direct): {stats['cells_filled_direct']}")
        print(f"Cells with formulas: {stats['cells_with_formulas']}")
        print(f"Total impact: {stats['cells_filled_direct'] + stats['cells_with_formulas']} cells")
        print(f"Success rate: {report['validation']['success_rate']:.1f}%")
        print(f"\n✓ Output: {OUTPUT_FILE}")
        
    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
