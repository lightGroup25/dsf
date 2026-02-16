"""
Analyser NOTE 3B en détail pour comprendre sa structure
"""

import openpyxl
from pathlib import Path

TEMPLATE_FILE = Path(r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\templates\DSF Normal standard.xlsx")

def main():
    wb = openpyxl.load_workbook(TEMPLATE_FILE)
    ws = wb['NOTE 3B']
    
    print("="*80)
    print("ANALYSE DÉTAILLÉE - NOTE 3B")
    print("="*80)
    
    # Analyser colonnes A et B sur plusieurs lignes
    for col in ['A', 'B']:
        print(f"\nColonne {col}:")
        print("-" * 60)
        for row_idx in range(10, 30):
            val = ws[f'{col}{row_idx}'].value
            val_type = type(val).__name__ if val else "None"
            val_display = str(val)[:60] if val else "(vide)"
            print(f"  {col}{row_idx}: [{val_type:8s}] {val_display}")

if __name__ == "__main__":
    main()
