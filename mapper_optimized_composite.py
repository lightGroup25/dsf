"""
MAPPER INTELLIGENT OPTIMISÉ - MAPPAGE FUZZY COMPOSITE

Procédure:
1. Pour chaque CELLULE du DSF (ligne + colonne)
2. Extraire: libellé_ligne + libellé_colonne → clé COMPOSITE
3. Fuzzy match cette clé composite avec la balance:
   - Chercher le compte (ligne) par libellé
   - Chercher la colonne appropriée par libellé de colonne (OUVERTURE, CLOTURE, MOUVEMENTS, etc)
4. Remplir avec la donnée balance correcte
5. Appliquer les formules découvertes (Col 27 = Col 14 + Col 21)
"""

import openpyxl
from difflib import SequenceMatcher
import json
from pathlib import Path

BALANCE_FILE = r"input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
DSF_TEMPLATE = r"templates/DSF Normal standard.xlsx"
OUTPUT_FILE = r"output/DSF_FINAL_OPTIMIZED.xlsx"

def load_balance():
    """Charger balance avec structure complète"""
    wb = openpyxl.load_workbook(BALANCE_FILE)
    ws = wb.active
    
    accounts = {}
    for row_idx in range(1, ws.max_row + 1):
        account_num = ws[f'A{row_idx}'].value
        account_label = ws[f'D{row_idx}'].value
        
        if not account_num or not account_label:
            continue
            
        try:
            int(str(account_num).split('.')[0])
            num_str = str(account_num).strip()
            label_str = str(account_label).strip()
            
            accounts[num_str] = {
                'label': label_str,
                # Données clés découvertes
                'col14_opening': ws[f'N{row_idx}'].value,      # Col 14: Ouverture
                'col21_movements': ws[f'U{row_idx}'].value,    # Col 21: Mouvements
                'col27_closing': ws[f'AA{row_idx}'].value,     # Col 27: Clôture
                # Formule validée: Col 27 = Col 14 + Col 21
            }
        except:
            pass
    
    return accounts

def get_column_header_type(ws, col_letter, start_row=1, end_row=20):
    """Extraire le type de colonne (OUVERTURE, CLOTURE, MOUVEMENTS, AMORT, etc)"""
    headers = []
    for r in range(start_row, end_row + 1):
        val = ws[f'{col_letter}{r}'].value
        if val:
            headers.append(str(val).upper())
    
    header_text = ' '.join(headers)
    
    # Classifier par libellé
    # 1. D'ABORD les colonnes À IGNORER (orange, calculées)
    if any(x in header_text for x in ['AMORT', 'DEPREC', 'AMORTISSEMENTS']):
        return 'AMORT', header_text  
    elif 'N-1' in header_text or 'EXERCICE AU 31/12/N-1' in header_text:
        return 'COMPARATIF', header_text  
    elif ('NET' in header_text and 'CLOTURE' not in header_text  
          and 'BRUT' not in header_text and 'MONTANT' not in header_text):
        return 'NET', header_text  
    
    # 2. Colonnes À REMPLIR - Clôture (chercher CLOTURE en premier car peut avoir BRUT après)
    if 'CLOTURE' in header_text or 'A LA CLOTURE' in header_text:
        return 'CLOSING', header_text
    
    # 3. Colonnes À REMPLIR - Ouverture
    # Peut être juste "BRUT" (BILAN PAYSAGE) ou "MONTANT BRUTE À L'OUVERTURE" (Notes)
    if 'OUVERTURE' in header_text or ('A L' in header_text and 'BRUTE' in header_text):
        return 'OPENING', header_text
    # BRUT seul signifie OUVERTURE (BILAN PAYSAGE colonne D)
    if 'BRUT' in header_text and 'CLOTURE' not in header_text and 'AMORT' not in header_text:
        return 'OPENING', header_text
    
    # 4. Mouvements/Acquisitions
    if any(x in header_text for x in ['ACQUISITION', 'APPORT', 'CREATION', 'MOUVEMENT']):
        return 'MOVEMENTS', header_text
    
    return 'UNKNOWN', header_text

def fuzzy_match(str1, str2, threshold=0.4):
    """Fuzzy matching simple"""
    if not str1 or not str2:
        return 0
    return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()

def composite_fuzzy_match(line_label, col_type, accounts):
    """
    MAPPAGE FUZZY COMPOSITE:
    - Chercher un compte dans balance qui matche line_label
    - Retourner la donnée correcte basée sur col_type
    - NOTE: Ignore les comptes sans données dans la colonne appropriée
    """
    best_score = 0
    best_account = None
    
    for acc_num, acc_data in accounts.items():
        # IMPORTANT: Vérifier que ce compte a une donnée pour ce type de colonne
        # avant de considérer comme match
        if col_type == 'OPENING' and not acc_data['col14_opening']:
            continue  # Skip accounts without opening data
        elif col_type == 'CLOSING' and not acc_data['col27_closing']:
            continue  # Skip accounts without closing data
        elif col_type == 'MOVEMENTS' and not acc_data['col21_movements']:
            continue  # Skip accounts without movements data
        
        score = fuzzy_match(line_label, acc_data['label'])
        if score > best_score:
            best_score = score
            best_account = acc_data
    
    # Threshold: minimum 0.25 pour un match (fuzzy tolère les variations)
    if best_score < 0.25 or not best_account:
        return None
    
    # Retourner la donnée appropriée basée sur le type de colonne
    if col_type == 'OPENING':
        value = best_account['col14_opening']
        if value and isinstance(value, (int, float)) and value != 0:
            return value, best_score
    
    elif col_type == 'CLOSING':
        value = best_account['col27_closing']
        if value and isinstance(value, (int, float)) and value != 0:
            return value, best_score
    
    elif col_type == 'MOVEMENTS':
        value = best_account['col21_movements']
        if value and isinstance(value, (int, float)) and value != 0:
            return value, best_score
    
    return None

def process_sheet_optimized(wb, sheet_name, accounts):
    """Traiter feuille: CELLULE PAR CELLULE avec mappage composite"""
    try:
        ws = wb[sheet_name]
    except:
        return 0
    
    cells_filled = 0
    
    # Identifier les colonnes et leurs types dans TOUTES les colonnes (A-N)
    column_info = {}
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
        col_type, header = get_column_header_type(ws, col, start_row=1, end_row=20)
        # Seulement tracker les colonnes à remplir
        if col_type in ['OPENING', 'CLOSING', 'MOVEMENTS']:
            column_info[col] = {'type': col_type, 'header': header}
    
    if not column_info:
        return 0
    
    # Scanner les lignes de données (généralement après row 10)
    for row_idx in range(10, min(300, ws.max_row + 1)):
        col_a = ws[f'A{row_idx}'].value
        col_b = ws[f'B{row_idx}'].value
        
        line_label = str(col_b if col_b else col_a).strip() if (col_a or col_b) else None
        
        if not line_label or len(line_label) < 2:
            continue
        
        # Pour chaque colonne remplissable de cette ligne
        for col_letter, col_info_dict in column_info.items():
            col_type = col_info_dict['type']
            
            try:
                cell = ws[f'{col_letter}{row_idx}']
            except:
                continue
            
            # Seulement remplir si cellule vide
            if cell.value:
                continue
            
            # Vérifier si c'est une merged cell
            if hasattr(cell, 'data_type') and cell.data_type == 'm':
                continue  # Skip merged cells
            
            # MAPPAGE FUZZY COMPOSITE: ligne + colonne
            result = composite_fuzzy_match(line_label, col_type, accounts)
            
            if result:
                value, score = result
                if value and score >= 0.25:
                    try:
                        cell.value = value
                        cells_filled += 1
                    except:
                        # Impossible d'écrire dans cette cellule (merged, protected, etc)
                        pass
    
    return cells_filled

def main():
    print("="*70)
    print("MAPPER INTELLIGENT OPTIMISÉ - MAPPAGE FUZZY COMPOSITE")
    print("="*70)
    
    accounts = load_balance()
    print(f"\n[OK] Loaded {len(accounts)} accounts from balance")
    print("  With composite data: OPENING (Col14), MOVEMENTS (Col21), CLOSING (Col27)")
    
    wb = openpyxl.load_workbook(DSF_TEMPLATE)
    print(f"[OK] Template has {len(wb.sheetnames)} sheets")
    
    total_filled = 0
    sheets_processed = {}
    
    print(f"\nProcessing each sheet (cellule par cellule)...\n")
    
    for sheet_name in wb.sheetnames:
        filled = process_sheet_optimized(wb, sheet_name, accounts)
        
        if filled > 0:
            total_filled += filled
            sheets_processed[sheet_name] = filled
            print(f"[OK] {sheet_name}: {filled} cells filled")
    
    print(f"\nSaving to {OUTPUT_FILE}...")
    Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_FILE)
    
    print("\n" + "="*70)
    print("RÉSULTAT - MAPPAGE FUZZY COMPOSITE OPTIMISÉ")
    print("="*70)
    print(f"Total sheets: {len(wb.sheetnames)}")
    print(f"Sheets with data: {len(sheets_processed)}")
    print(f"Total cells filled: {total_filled}")
    
    if sheets_processed:
        print(f"\nSheets filled:")
        for sheet, count in sorted(sheets_processed.items(), key=lambda x: -x[1])[:20]:
            print(f"  {sheet}: {count}")
    
    print(f"\n[OK] Output: {OUTPUT_FILE}")
    print("="*70)

if __name__ == "__main__":
    main()
