"""
MAPPER V7 - OPTIMIZED WITH FORMULA SUPPORT + CALCULATIONS
==========================================================

Améliorations:
1. Remplir les 3 colonnes: OPENING, MOVEMENTS, CLOSING
2. Calculer CLOSING = OPENING + MOVEMENTS automatiquement
3. Couvrir TOUS les comptes (pas de skipping)
4. Utiliser formules découvertes dans les en-têtes pour guider la classification
5. Couvrir actifs ET passifs complètement
6. 🆕 CALCULS AUTOMATIQUES: TOTAL, DIFFÉRENCES, POURCENTAGES, RATIOS
7. 🆕 GÉNÉRATION FORMULES EXCEL pour calculs dynamiques
8. 🆕 PROTECTION INTELLIGENTE: Détection colonnes labels par sheet

Procédure:
- Pour chaque cellule du DSF:
  1. Identifier type de colonne (OPENING/MOVEMENTS/CLOSING/AMORT/etc)
  2. 🆕 Vérifier si cellule nécessite un CALCUL (ligne TOTAL, colonne VARIATION, etc)
  3. SI calcul requis → Appliquer calcul/formule
  4. SINON → Matcher la ligne avec le compte balance
  5. Remplir avec donnée appropriée
"""

import openpyxl
from difflib import SequenceMatcher
from pathlib import Path
from calculation_manager import get_calculation_manager

BALANCE_FILE = r"input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
DSF_TEMPLATE = r"templates/DSF Normal standard.xlsx"
OUTPUT_FILE = r"output/DSF_FINAL_V7_FIXED.xlsx"  # Nouveau nom pour éviter conflit

def load_balance_complete():
    """Charger balance COMPLÈTE - toutes les colonnes possibles"""
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
            
            # Charger TOUTES les données disponibles
            # Vérifier les colonnes numériques (debit/credit) aussi
            col_j_debit = ws[f'J{row_idx}'].value    # Debit 2023
            col_m_credit = ws[f'M{row_idx}'].value   # Credit 2023
            col_n_solde = ws[f'N{row_idx}'].value    # Balance 2023 (opening)
            
            col_q_debit = ws[f'Q{row_idx}'].value    # Debit mouv
            col_t_credit = ws[f'T{row_idx}'].value   # Credit mouv
            col_u_total = ws[f'U{row_idx}'].value    # Total mouv (col 21)
            
            col_w_debit = ws[f'W{row_idx}'].value    # Debit cumul
            col_z_credit = ws[f'Z{row_idx}'].value   # Credit cumul
            col_aa_solde = ws[f'AA{row_idx}'].value  # Balance cumul (closing)
            
            # Calculer les mouvements si pas disponibles directement
            movements = col_u_total
            if not movements and (col_q_debit or col_t_credit):
                q_val = col_q_debit or 0
                t_val = col_t_credit or 0
                movements = (q_val or 0) - (t_val or 0)
            
            # Calculer closing si pas disponible
            closing = col_aa_solde
            if not closing and col_n_solde and movements:
                closing = col_n_solde + movements
            
            accounts[num_str] = {
                'label': label_str,
                'opening': col_n_solde,           # Col 14: Solde ouverture
                'movements': movements,            # Col 21: Total mouvements
                'closing': closing,                # Col 27: Solde clôture
                'debit_2023': col_j_debit,
                'credit_2023': col_m_credit,
                'debit_mouv': col_q_debit,
                'credit_mouv': col_t_credit,
            }
        except:
            pass
    
    return accounts

def is_label_column(ws, col_letter, data_start_row=15, sample_rows=20):
    """Déterminer si une colonne contient des LABELS (libellés de lignes) au lieu de données
    
    Critères:
    - Analyse les lignes de données (pas les headers)
    - Si majorité de texte descriptif long → LABEL
    - Si majorité de nombres ou vide → DATA
    """
    text_count = 0
    number_count = 0
    empty_count = 0
    long_text_count = 0  # Texte > 10 caractères = description
    
    # Analyser un échantillon de lignes de données
    for row_idx in range(data_start_row, min(data_start_row + sample_rows, ws.max_row + 1)):
        try:
            cell_value = ws[f'{col_letter}{row_idx}'].value
        except:
            continue
        
        if cell_value is None or cell_value == '':
            empty_count += 1
        elif isinstance(cell_value, str):
            text_count += 1
            # Si texte long et descriptif → indicateur de colonne de labels
            if len(cell_value.strip()) > 10:
                long_text_count += 1
        elif isinstance(cell_value, (int, float)):
            number_count += 1
    
    total_analyzed = text_count + number_count + empty_count
    
    # Si pas assez de données, pas une colonne de labels
    if total_analyzed < 5:
        return False
    
    # Si > 50% de texte descriptif long → LABEL COLUMN
    if long_text_count > (total_analyzed * 0.5):
        return True
    
    # Si > 70% de texte (même court) → probablement LABEL
    if text_count > (total_analyzed * 0.7):
        return True
    
    return False

def get_column_type_v7(ws, col_letter, start_row=1, end_row=20):
    """Détecter type de colonne avec support des formules et patterns"""
    headers = []
    for r in range(start_row, min(end_row + 1, ws.max_row + 1)):
        val = ws[f'{col_letter}{r}'].value
        if val:
            headers.append(str(val).upper())
    
    header_text = ' '.join(headers)
    
    # 1. Colonnes À IGNORER (protégées)
    if any(x in header_text for x in ['AMORT', 'DEPREC', 'AMORTISSEMENT']):
        return 'SKIP_AMORT'
    if 'N-1' in header_text or 'EXERCICE AU 31/12/N-1' in header_text or 'COMPARATIF' in header_text:
        return 'SKIP_COMPARATIF'
    if 'NET' in header_text and 'CLOTURE' not in header_text and 'BRUT' not in header_text:
        # Mais pas si c'est "NET À LA CLOTURE"
        if 'A LA' not in header_text:
            return 'SKIP_NET'
    
    # 2. Colonnes CALCULÉES (TOTAL, RÉSULTAT)
    if 'TOTAL' in header_text and any(x in header_text for x in ['ACQUISIT', 'APPORT', 'CREATION']):
        return 'CALC_TOTAL_MOVEMENTS'
    if 'TOTAL' in header_text and ('OUVERTUR' in header_text or 'BRUT' in header_text):
        return 'CALC_TOTAL_OPENING'
    if 'RESULTAT' in header_text or 'SOLDE' in header_text.upper():
        return 'CALC_RESULT'
    
    # 3. Colonnes À REMPLIR - CLÔTURE (en premier!)
    if 'CLOTURE' in header_text or 'A LA CLOTURE' in header_text:
        return 'CLOSING'
    
    # 4. Colonnes À REMPLIR - OUVERTURE
    if 'OUVERTUR' in header_text or ('A L' in header_text and 'BRUT' in header_text):
        return 'OPENING'
    if 'BRUT' in header_text and 'CLOTURE' not in header_text and 'AMORT' not in header_text:
        return 'OPENING'
    
    # 5. Colonnes À REMPLIR - MOUVEMENTS
    if any(x in header_text for x in ['ACQUISITION', 'APPORT', 'CREATION', 'MOUVEMENT', 'CHANGE']):
        return 'MOVEMENTS'
    
    # 6. Autres types de données (virements, révaluations, etc)
    if 'VIREMENT' in header_text or 'RECLASSEMENT' in header_text:
        return 'MOVEMENTS'
    
    if 'REEVALUATION' in header_text or 'REESTIMATION' in header_text:
        return 'MOVEMENTS'
    
    if 'CESSION' in header_text or 'SORTIE' in header_text or 'DECOMM' in header_text:
        return 'MOVEMENTS'
    
    return 'UNKNOWN'

def fuzzy_match(str1, str2, threshold=0.25):
    """Fuzzy matching - très permissif (0.25 = 25%)"""
    if not str1 or not str2:
        return 0
    return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()

def find_best_account(line_label, col_type, accounts):
    """Trouver le meilleur compte pour cette ligne"""
    best_score = 0
    best_account = None
    
    for acc_num, acc_data in accounts.items():
        # IMPORTANT: Ne pas skipper basé sur col_type - vérifier la valeur réelle
        # Chaque compte peut avoir une données pour n'importe quel type
        score = fuzzy_match(line_label, acc_data['label'])
        
        if score > best_score:
            best_score = score
            best_account = (acc_num, acc_data)
    
    return best_account, best_score

def get_value_for_column(acc_data, col_type):
    """Extraire la valeur appropriée pour ce type de colonne"""
    if col_type == 'OPENING':
        return acc_data.get('opening')
    elif col_type == 'CLOSING':
        return acc_data.get('closing')
    elif col_type == 'MOVEMENTS':
        return acc_data.get('movements')
    elif col_type == 'CALC_TOTAL_OPENING':
        return acc_data.get('opening')
    elif col_type == 'CALC_TOTAL_MOVEMENTS':
        return acc_data.get('movements')
    
    return None

def process_sheet_v7(wb, sheet_name, accounts):
    """Traiter feuille avec support complet des formules"""
    try:
        ws = wb[sheet_name]
    except:
        return 0, {}
    
    cells_filled = 0
    debug_info = {}
    
    # Étape 1: Identifier colonnes remplissables
    column_info = {}
    # Analyser TOUTES les colonnes - détecter dynamiquement les labels
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
        # PROTECTION: Vérifier si c'est une colonne de labels
        if is_label_column(ws, col, data_start_row=15, sample_rows=20):
            # Skip cette colonne - c'est une colonne de libellés
            continue
        
        col_type = get_column_type_v7(ws, col, start_row=1, end_row=20)
        
        # Tracker seulement les colonnes à remplir
        if col_type in ['OPENING', 'CLOSING', 'MOVEMENTS', 
                        'CALC_TOTAL_OPENING', 'CALC_TOTAL_MOVEMENTS', 'CALC_RESULT']:
            column_info[col] = col_type
    
    if not column_info:
        return 0, debug_info
    
    debug_info['columns_found'] = column_info
    
    # NOUVEAU: Créer gestionnaire de calculs
    calc_mgr = get_calculation_manager(ws, column_info)
    
    # Étape 2: Scanner les rows de données
    for row_idx in range(10, min(300, ws.max_row + 1)):
        col_a = ws[f'A{row_idx}'].value
        col_b = ws[f'B{row_idx}'].value
        
        line_label = str(col_b if col_b else col_a).strip() if (col_a or col_b) else None
        
        if not line_label or len(line_label) < 2:
            continue
        
        # Pour chaque colonne remplissable
        for col_letter, col_type in column_info.items():
            try:
                cell = ws[f'{col_letter}{row_idx}']
            except:
                continue
            
            # PROTECTION CRITIQUE: Ne JAMAIS écraser une cellule avec du texte (libellé)
            # - Si c'est une string non-vide → SKIP (c'est un header/libellé)
            # - Si c'est un nombre → SKIP (déjà rempli)
            # - Si c'est vide ou None → OK pour remplir
            if cell.value is not None:
                if isinstance(cell.value, str):
                    # Si la string contient du texte significatif (pas juste espaces)
                    if cell.value.strip():
                        continue  # SKIP: libellé détecté
                else:
                    # C'est un nombre - déjà rempli
                    continue
            
            # Vérifier merged cells
            if hasattr(cell, 'data_type') and cell.data_type == 'm':
                continue
            
            # NOUVEAU: Vérifier si cette cellule nécessite un CALCUL au lieu d'un remplissage
            should_calc, calc_type, extra_info = calc_mgr.should_calculate_cell(row_idx, col_letter, col_type)
            
            if should_calc:
                # CALCULER (somme, différence, formule complexe) au lieu de remplir depuis balance
                try:
                    calc_value = calc_mgr.apply_calculation(
                        row_idx, col_letter, col_type, calc_type,
                        extra_info=extra_info,  # Passer les infos supplémentaires (formules template)
                        use_formulas=True  # True = formules Excel, False = valeurs
                    )
                    
                    if calc_value:
                        cell.value = calc_value
                        cells_filled += 1
                except Exception as e:
                    # Si calcul échoue, log et continue
                    pass
                
                continue  # Passer à la cellule suivante
            
            # SINON: Remplir depuis balance (code existant)
            # Matcher avec balance
            best_account, score = find_best_account(line_label, col_type, accounts)
            
            if not best_account or score < 0.25:
                continue
            
            acc_num, acc_data = best_account
            
            # Obtenir la valeur appropriée
            value = get_value_for_column(acc_data, col_type)
            
            # Si c'est une cellule de CLOSING et qu'on n'a pas la valeur directe
            # Mais qu'on a OPENING + MOVEMENTS → CALCULER
            if col_type == 'CLOSING' and not value:
                opening = acc_data.get('opening')
                movements = acc_data.get('movements')
                if opening and movements:
                    value = opening + movements
            
            # Remplir si valeur valide
            if value and isinstance(value, (int, float)) and value != 0:
                try:
                    cell.value = value
                    cells_filled += 1
                except:
                    pass
    
    return cells_filled, debug_info

def main():
    print("="*80)
    print("MAPPER V7 - OPTIMIZED WITH FORMULAS")
    print("="*80)
    
    accounts = load_balance_complete()
    print(f"\n[OK] Loaded {len(accounts)} accounts from balance")
    print(f"     Available data: OPENING, MOVEMENTS, CLOSING columns")
    
    wb = openpyxl.load_workbook(DSF_TEMPLATE)
    print(f"[OK] Template has {len(wb.sheetnames)} sheets")
    
    print(f"\nProcessing each sheet...\n")
    
    total_filled = 0
    sheets_data = {}
    
    for sheet_name in wb.sheetnames:
        filled, debug = process_sheet_v7(wb, sheet_name, accounts)
        
        if filled > 0:
            total_filled += filled
            sheets_data[sheet_name] = filled
            print(f"[+] {sheet_name:30s}: {filled:3d} cells")
        elif 'columns_found' in debug and debug['columns_found']:
            # Afficher aussi les sheets où des colonnes ont été détectées mais rien rempli
            col_count = len(debug['columns_found'])
            # print(f"[-] {sheet_name:30s}: 0 cells ({col_count} columns detected)")
    
    print(f"\nSaving to {OUTPUT_FILE}...")
    Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_FILE)
    
    print("\n" + "="*80)
    print("RESULT - MAPPER V7 OPTIMIZED")
    print("="*80)
    print(f"Total sheets: {len(wb.sheetnames)}")
    print(f"Sheets with data: {len(sheets_data)}")
    print(f"Total cells filled: {total_filled}")
    
    if sheets_data:
        print(f"\nTop sheets by cells filled:")
        for sheet, count in sorted(sheets_data.items(), key=lambda x: -x[1])[:20]:
            print(f"  {sheet:40s}: {count:4d} cells")
        
        if len(sheets_data) > 20:
            print(f"  ... and {len(sheets_data) - 20} more sheets")
    
    print(f"\n[OK] Output: {OUTPUT_FILE}")
    print("="*80)

if __name__ == "__main__":
    main()
