"""
AUDIT DSF V7 OUTPUT - Vérifier si des libellés ont été écrasés ou des zones protégées remplies
"""

import openpyxl
from pathlib import Path

OUTPUT_FILE = Path(r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_FINAL_V7_OPTIMIZED.xlsx")
TEMPLATE_FILE = Path(r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\templates\DSF Normal standard.xlsx")

def audit_sheet(ws_output, ws_template, sheet_name):
    """Compare template vs output pour détecter les anomalies"""
    issues = []
    
    # Vérifier les 20 premières lignes (headers) pour écrasement de libellés
    for row_idx in range(1, 21):
        for col_idx in range(1, 15):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            cell_template = ws_template.cell(row=row_idx, column=col_idx)
            cell_output = ws_output.cell(row=row_idx, column=col_idx)
            
            # Si le template a un libellé texte et que l'output a un nombre → PROBLÈME
            if cell_template.value:
                template_val = cell_template.value
                output_val = cell_output.value
                
                if isinstance(template_val, str) and template_val.strip():
                    if isinstance(output_val, (int, float)):
                        issues.append({
                            'type': 'HEADER_OVERWRITTEN',
                            'cell': f'{col_letter}{row_idx}',
                            'template': template_val[:50],
                            'output': output_val,
                        })
    
    # Vérifier les lignes de données (21-100) pour colonnes protégées remplies
    for row_idx in range(21, min(100, ws_output.max_row + 1)):
        for col_idx in range(1, 15):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            cell_output = ws_output.cell(row=row_idx, column=col_idx)
            
            # Détecter si c'est une colonne protégée (en vérifiant les headers)
            header_text = ""
            for h_row in range(1, 21):
                h_val = ws_output.cell(row=h_row, column=col_idx).value
                if h_val:
                    header_text += str(h_val).upper() + " "
            
            is_protected = any(kw in header_text for kw in 
                             ['AMORT', 'DEPREC', 'NET', 'N-1', 'COMPARATIF'])
            
            if is_protected and cell_output.value:
                if isinstance(cell_output.value, (int, float)) and cell_output.value > 1000:
                    issues.append({
                        'type': 'PROTECTED_ZONE_FILLED',
                        'cell': f'{col_letter}{row_idx}',
                        'value': cell_output.value,
                        'column_type': header_text[:50],
                    })
    
    return issues

def main():
    if not OUTPUT_FILE.exists():
        print(f"ERROR: {OUTPUT_FILE} not found!")
        return
    
    if not TEMPLATE_FILE.exists():
        print(f"ERROR: {TEMPLATE_FILE} not found!")
        return
    
    print("="*100)
    print("AUDIT DSF V7 - Vérification des anomalies")
    print("="*100)
    
    wb_output = openpyxl.load_workbook(OUTPUT_FILE)
    wb_template = openpyxl.load_workbook(TEMPLATE_FILE)
    
    all_issues = {}
    total_header_issues = 0
    total_protected_issues = 0
    
    # Vérifier les principales sheets
    test_sheets = ['BILAN PAYSAGE', 'NOTE 3A', 'NOTE 3B', 'NOTE 28', 'Note 1']
    
    for sheet_name in test_sheets:
        if sheet_name not in wb_output.sheetnames or sheet_name not in wb_template.sheetnames:
            continue
        
        ws_output = wb_output[sheet_name]
        ws_template = wb_template[sheet_name]
        
        issues = audit_sheet(ws_output, ws_template, sheet_name)
        
        if issues:
            all_issues[sheet_name] = issues
            header_count = sum(1 for i in issues if i['type'] == 'HEADER_OVERWRITTEN')
            protected_count = sum(1 for i in issues if i['type'] == 'PROTECTED_ZONE_FILLED')
            
            total_header_issues += header_count
            total_protected_issues += protected_count
            
            print(f"\n[ISSUES TROUVÉS] {sheet_name}")
            print(f"  - Libellés écrasés: {header_count}")
            print(f"  - Zones protégées remplies: {protected_count}")
            
            # Montrer quelques exemples
            for issue in issues[:5]:
                if issue['type'] == 'HEADER_OVERWRITTEN':
                    print(f"    ⚠️ {issue['cell']}: '{issue['template']}' → {issue['output']}")
                else:
                    print(f"    ⚠️ {issue['cell']}: valeur {issue['value']} dans zone protégée")
    
    print("\n" + "="*100)
    print("RÉSUMÉ AUDIT")
    print("="*100)
    print(f"Sheets auditées: {len(test_sheets)}")
    print(f"Sheets avec problèmes: {len(all_issues)}")
    print(f"Total libellés écrasés: {total_header_issues}")
    print(f"Total zones protégées remplies: {total_protected_issues}")
    
    if total_header_issues > 0 or total_protected_issues > 0:
        print("\n❌ ANOMALIES DÉTECTÉES - Le mapper V7 a des bugs!")
        print("\nActions requises:")
        print("1. Ne remplir QUE les lignes de données (row >= 21 généralement)")
        print("2. Skip les colonnes AMORT, NET, COMPARATIF, N-1")
        print("3. Vérifier que cell.value est vide AVANT d'écrire")
    else:
        print("\n✅ Aucune anomalie détectée")

if __name__ == "__main__":
    main()
