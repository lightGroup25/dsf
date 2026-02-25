#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import openpyxl

wb = openpyxl.load_workbook(r'templates\DSF Normal standard.xlsx', data_only=False)

# Vérifier les feuilles qui pourraient contenir les taux
target_sheets = ['ENTETE', 'INFORMATIONS GENERALES', 'Fiche R1', 'Fiche R2']

for sheet_name in target_sheets:
    if sheet_name not in wb.sheetnames:
        print(f'\n[SKIP] {sheet_name} - not found')
        continue
    
    ws = wb[sheet_name]
    print(f"\n{'='*120}")
    print(f"FEUILLE: {sheet_name}")
    print(f"{'='*120}")
    
    # Chercher des lignes avec "TAUX", "%", "AMORT", "DEPRECI"
    found = False
    for row_num in range(1, min(100, ws.max_row + 1)):
        for col_num in range(1, min(15, ws.max_column + 1)):
            cell = ws.cell(row_num, col_num)
            if cell.value:
                val_str = str(cell.value).upper()
                if any(keyword in val_str for keyword in ['TAUX', '%', 'AMORT', 'DEPRECI', 'IMMOBIL', 'TERRAIN', 'BATIMENT', 'VEHICULE']):
                    if not found:
                        print(f"\nElements trouves:")
                        found = True
                    print(f"  Row {row_num:3d}, Col {col_num:2d}: {str(cell.value)[:70]}")
    
    if not found:
        # Afficher les 10 premieres lignes meme si rien de special
        print(f"\nPremiers contenus (rows 1-15, cols 1-6):")
        for row_num in range(1, 16):
            for col_num in range(1, 7):
                cell = ws.cell(row_num, col_num)
                if cell.value:
                    print(f"  Row {row_num}, Col {col_num}: {str(cell.value)[:60]}")

wb.close()
