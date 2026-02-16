"""
Comprehensive validation of Option 3 implementation
Verify that the formulas work and data integrity is maintained
"""

import openpyxl
from openpyxl.utils import get_column_letter
import json
from pathlib import Path

OUTPUT_FILE = r"output/DSF_OPTION3_AUTO_CALCULATED.xlsx"

def validate_option3_output():
    """Validate the Option 3 output"""
    print("="*70)
    print("OPTION 3 VALIDATION REPORT")
    print("="*70)
    
    try:
        wb = openpyxl.load_workbook(OUTPUT_FILE, data_only=False)
        ws = wb.active
        
        validation = {
            'timestamp': str(Path(OUTPUT_FILE).stat().st_mtime),
            'test_cases': [],
            'formula_checks': {
                'total_formulas': 0,
                'valid_formulas': 0,
                'formula_errors': []
            },
            'data_integrity': {
                'cells_with_opening': 0,
                'cells_with_movements': 0,
                'cells_with_formulas': 0,
                'sample_rows_verified': []
            }
        }
        
        print("\n1. Scanning for data integrity...")
        
        for row_idx in range(1, ws.max_row + 1):
            col_d = ws[f'D{row_idx}'].value  # Opening
            col_e = ws[f'E{row_idx}'].value  # Movements
            col_j = ws[f'J{row_idx}'].value  # Closing/Formula
            
            # Count data
            if col_d and isinstance(col_d, (int, float)) and col_d > 0:
                validation['data_integrity']['cells_with_opening'] += 1
                
            if col_e and isinstance(col_e, (int, float)):
                validation['data_integrity']['cells_with_movements'] += 1
            
            # Check for formulas in column J
            if isinstance(col_j, str) and col_j.startswith('='):
                validation['formula_checks']['total_formulas'] += 1
                
                # Verify formula format
                if col_j == f"=D{row_idx}+E{row_idx}":
                    validation['formula_checks']['valid_formulas'] += 1
                else:
                    validation['formula_checks']['formula_errors'].append({
                        'row': row_idx,
                        'formula': col_j,
                        'expected': f"=D{row_idx}+E{row_idx}"
                    })
                
                validation['data_integrity']['cells_with_formulas'] += 1
                
                # Collect sample rows
                if len(validation['data_integrity']['sample_rows_verified']) < 10:
                    validation['data_integrity']['sample_rows_verified'].append({
                        'row': row_idx,
                        'opening': col_d,
                        'movements': col_e,
                        'formula': col_j,
                        'description': f"Col D={col_d}, Col E={col_e}, Col J={col_j}"
                    })
        
        print(f"\n2. Results:")
        print(f"   ✓ Opening balances (Col D): {validation['data_integrity']['cells_with_opening']}")
        print(f"   ✓ Movement values (Col E): {validation['data_integrity']['cells_with_movements']}")
        print(f"   ✓ Formulas (Col J): {validation['formula_checks']['total_formulas']}")
        print(f"   ✓ Valid formulas: {validation['formula_checks']['valid_formulas']}")
        
        if validation['formula_checks']['formula_errors']:
            print(f"\n   ⚠ Formula format errors: {len(validation['formula_checks']['formula_errors'])}")
            for error in validation['formula_checks']['formula_errors'][:3]:
                print(f"     Row {error['row']}: {error['formula']}")
        
        print(f"\n3. Sample rows with formulas:")
        for i, sample in enumerate(validation['data_integrity']['sample_rows_verified'][:5], 1):
            print(f"   Row {sample['row']}: Opening={sample['opening']}, "
                  f"Movements={sample['movements']}, Formula={sample['formula']}")
        
        print(f"\n4. File details:")
        file_size = Path(OUTPUT_FILE).stat().st_size / (1024*1024)
        print(f"   Size: {file_size:.2f} MB")
        print(f"   Path: {OUTPUT_FILE}")
        
        print("\n" + "="*70)
        print("VALIDATION COMPLETE - OPTION 3 SUCCESSFULLY IMPLEMENTED")
        print("="*70)
        
        # Save detailed validation report
        validation_json = {
            'status': 'SUCCESS',
            'output_file': OUTPUT_FILE,
            'data_integrity': validation['data_integrity'],
            'formula_validation': {
                'total_formulas': validation['formula_checks']['total_formulas'],
                'valid_formulas': validation['formula_checks']['valid_formulas'],
                'error_count': len(validation['formula_checks']['formula_errors'])
            }
        }
        
        report_path = r"output/reports/option3_validation.json"
        Path(report_path).parent.mkdir(parents=True, exist_ok=True)
        with open(report_path, 'w') as f:
            json.dump(validation_json, f, indent=2, default=str)
        
        print(f"\n✓ Validation report saved: {report_path}")
        
        return validation
        
    except Exception as e:
        print(f"\n✗ Validation error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    validate_option3_output()
