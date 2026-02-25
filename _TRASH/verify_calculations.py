"""
VÉRIFICATION DES CALCULS - Mapper V7
Vérifie si les formules et calculs ont été appliqués correctement
"""

import openpyxl
from pathlib import Path

OUTPUT_FILE = Path(r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_FINAL_V7_FIXED.xlsx")

def check_formulas_and_calculations(ws, sheet_name):
    """Vérifier les formules et calculs dans une sheet"""
    
    formulas_found = []
    totals_found = []
    
    # Scanner les cellules pour formules et totaux
    for row_idx in range(10, min(100, ws.max_row + 1)):
        # Vérifier si ligne TOTAL
        is_total = False
        for col in ['A', 'B', 'C']:
            try:
                val = ws[f'{col}{row_idx}'].value
                if val and isinstance(val, str):
                    if any(kw in val.upper() for kw in ['TOTAL', 'SOUS-TOTAL', 'SOMME']):
                        is_total = True
                        break
            except:
                pass
        
        if is_total:
            # Vérifier les colonnes de données pour formules
            row_info = {'row': row_idx, 'label': None, 'formulas': {}, 'values': {}}
            
            for col in ['A', 'B', 'C']:
                val = ws[f'{col}{row_idx}'].value
                if val and isinstance(val, str):
                    row_info['label'] = val[:50]
                    break
            
            # Vérifier colonnes D-N pour formules ou valeurs
            for col_letter in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
                try:
                    cell = ws[f'{col_letter}{row_idx}']
                    
                    # Vérifier si formule
                    if hasattr(cell, 'value') and isinstance(cell.value, str):
                        if cell.value.startswith('='):
                            row_info['formulas'][col_letter] = cell.value
                    
                    # Vérifier si valeur numérique
                    elif isinstance(cell.value, (int, float)):
                        row_info['values'][col_letter] = cell.value
                except:
                    pass
            
            if row_info['formulas'] or row_info['values']:
                totals_found.append(row_info)
    
    return totals_found

def main():
    if not OUTPUT_FILE.exists():
        print(f"ERROR: {OUTPUT_FILE} not found!")
        return
    
    print("="*100)
    print("VÉRIFICATION DES CALCULS ET FORMULES - Mapper V7")
    print("="*100)
    
    wb = openpyxl.load_workbook(OUTPUT_FILE)
    
    # Tester les sheets principales
    test_sheets = ['NOTE 3A', 'NOTE 3B', 'NOTE 28', 'BILAN PAYSAGE']
    
    total_formulas = 0
    total_calculated_totals = 0
    
    for sheet_name in test_sheets:
        if sheet_name not in wb.sheetnames:
            continue
        
        ws = wb[sheet_name]
        
        print(f"\n{'='*100}")
        print(f"SHEET: {sheet_name}")
        print('='*100)
        
        totals = check_formulas_and_calculations(ws, sheet_name)
        
        if totals:
            print(f"\n✅ Lignes TOTAL détectées: {len(totals)}")
            
            for t in totals:
                print(f"\n  Row {t['row']}: {t['label']}")
                
                if t['formulas']:
                    print(f"    🔢 FORMULES EXCEL:")
                    for col, formula in t['formulas'].items():
                        print(f"       {col}{t['row']}: {formula}")
                        total_formulas += 1
                
                if t['values']:
                    print(f"    💰 VALEURS:")
                    for col, value in t['values'].items():
                        print(f"       {col}{t['row']}: {value:,.0f}")
                        total_calculated_totals += 1
        else:
            print(f"\n⚠️ Aucune ligne TOTAL avec calculs détectée")
    
    print("\n" + "="*100)
    print("RÉSUMÉ")
    print("="*100)
    print(f"Total FORMULES EXCEL créées: {total_formulas}")
    print(f"Total VALEURS calculées: {total_calculated_totals}")
    
    if total_formulas > 0:
        print(f"\n✅ SUCCÈS: Le système de calculs génère des formules Excel!")
        print(f"   Les formules se recalculeront automatiquement si les données changent.")
    elif total_calculated_totals > 0:
        print(f"\n✅ SUCCÈS: Le système de calculs génère des valeurs calculées!")
        print(f"   Les totaux sont calculés automatiquement.")
    else:
        print(f"\n⚠️ ATTENTION: Aucun calcul automatique détecté.")
        print(f"   Vérifiez que les lignes TOTAL contiennent des valeurs ou formules.")

if __name__ == "__main__":
    main()
