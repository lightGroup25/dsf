#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SOLUTION FINALE - Remplir DSF avec TOUTES les données du bilan

Structure:
- Col 14 = Solde Ouverture (DSF Col 4)
- Col 21 = Mouvements année (DSF Col 5-9)
- Col 27 = Solde Clôture (DSF Col 10)

Formule: Col 27 = Col 14 + Col 21  (= Ouverture + Mouvements)
"""

import openpyxl
from pathlib import Path

print("="*130)
print("STRATEGIE CORRECTE")
print("="*130)

balance_file = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\input\BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"
dsf_template = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\templates\DSF Normal standard.xlsx"

wb_balance = openpyxl.load_workbook(balance_file, data_only=True)
wb_dsf = openpyxl.load_workbook(dsf_template, data_only=False)

ws_balance = wb_balance.active
ws_dsf = wb_dsf['NOTE 3A']

print("""
POUR CHAQUE COMPTE DANS BALANCE:
================================

1. Lire:  
   - Col 1: Numéro compte
   - Col 4: Intitulé
   - Col 14: Solde Ouverture
   - Col 21: Mouvements  
   - Col 27: Solde Clôture

2. Chercher la CORRESPONDANCE dans DSF:
   - Row avec même label de compte
   
3. Remplir DSF:
   - Col 4: = Balance Col 14 (Ouverture)
   - Col 10: = Balance Col 27 (Clôture)
   
4. Pour Col 5-9 (Mouvements détail):
   - QUESTION: Comment repartir le mouvement total (Col 21)?
   
OPTION A: Mettre TOUT le mouvement dans Col 5 (acquisitions)
   - Honnête si on sait que ce sont des acquisitions
   
OPTION B: Laisser vide
   - Transparent sur ce qu'on ne sait pas
   
OPTION C: Diviser entre colonnes
   - Risqué sans justification
""")

# Compter les comptes qui peuvent être remplis
filled_count = 0
for row_num in range(11, ws_balance.max_row + 1):
    col_1 = ws_balance.cell(row_num, 1).value
    col_14 = ws_balance.cell(row_num, 14).value
    col_27 = ws_balance.cell(row_num, 27).value
    
    # Vérifier si c'est un compte valide
    if col_1 and col_14 is not None or col_27 is not None:
        filled_count += 1

print(f"\n\nCOMPTES DISPONIBLES: {filled_count}")
print(f"CELLULES DSF à REMPLIR:")
print(f"  - Colonnes 4 (Ouverture): {filled_count} cellules")
print(f"  - Colonnes 10 (Clôture): {filled_count} cellules")  
print(f"  - TOTAL: {filled_count * 2} cellules")

print(f"\nPLUS, si on remplit Col 5 (Mouvements):")
print(f"  - Colonnes 5 (Mouvements): {filled_count} cellules")
print(f"  - TOTAL: {filled_count * 3} cellules")

wb_balance.close()
wb_dsf.close()

print("\n" + "="*130)
print("DECISION POUR VOUS")
print("="*130)

print("""
Voulez-vous que je code:

OPTION 1 (RECOMMANDEE - Honnête et complet):
   ✓ Col 4 = Balance Col 14 (Ouverture)
   ✓ Col 10 = Balance Col 27 (Clôture)
   ✗ Col 5-9 = Vides (mouvement total pas détaillé)
   → ~150-200 cellules remplies CORRECTEMENT
   → Rapport montrera ce qui était possible

OPTION 2 (Agressif - Plus de remplissage):
   ✓ Col 4 = Balance Col 14 (Ouverture)
   ✓ Col 5 = Balance Col 21 (Mouvements totaux)
   ✓ Col 10 = Balance Col 27 (Clôture)
   ✗ Col 6-9 = Vides
   → ~300-400 cellules remplies
   → Note: Col 5 contient TOUS les mouvements (pas juste acquisitions)

OPTION 3 (Intelligent - Utiliser formules):
   ✓ Col 4 = Balance Col 14
   ✓ Col 5 = Balance Col 21
   ✓ Col 10 = FORMULE Excel = Col 4 + Col 5 (auto-calcul!)
   → Les utilisateurs voient la logique
   → Si balance change, DSF se recalcule
   
Quelle option choisissez-vous?
""")
