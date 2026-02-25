#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
REMARQUE: La formule LOGIQUE est:
Col 4 (Ouverture) + Col 5-9 (Mouvements) = Col 10 (Cloture)

Cela signifie que si on remplit Col 4 et Col 5-9, Excel peut AUTO-CALCULER Col 10 avec une formule!
"""

print("="*130)
print("ANALYSE - FORMULE LOGIQUE ENTRE COLONNES DSF")
print("="*130)

print("""
COLONNE 4: MONTANT BRUTE A L'OUVERTURE DE L'EXERCICE
COLONNE 5: ACQUISITIONS APPORTS CREATIONS
COLONNE 6: VIREMENTS DE POSTE A POSTE
COLONNE 7: REEVALUATION 
COLONNE 8: CESSIONS SCISSIONS HORS SERVICE  [MOINS/NEGATIF]
COLONNE 9: VIREMENTS
COLONNE 10: MONTANT BRUT A LA CLOTURE

LA FORMULE LOGIQUE EST:
=========================================

Col 10 = Col 4 + Col 5 + Col 6 + Col 7 - Col 8 + Col 9

Ou plus simplement:
Col 10 = Col 4 + SOMME(Col 5 a Col 9)

CETTE FORMULE PEUT ETRE IMPLANTEE DANS EXCEL!
==============================================

Cela signifie que SI on peut remplir:
- Col 4 (Ouverture) 
- Col 5-9 (Détails mouvements)

ALORS Col 10 se CALCULE AUTOMATIQUEMENT!
""")

print("\n" + "="*130)
print("ANALYSE DES AUTRES COLONNES")
print("="*130)

print("""
NOTE 3B (DEPRECIATIONS/AMORTISSEMENTS):
========================================

Col 4: NATURE DU CONTRAT (codes I; M; A)
   → I = Immobilisé
   → M = ?
   → A = Amortissable?

Col 5: Dépréciations à l'ouverture (ouverture)
Col 7: AUGMENTATIONS (augmentation d'amortissement pendant l'année)
Col 11: DIMINUTIONS (diminution d'amortissement)
Col ?: Dépréciations à la clôture

LA FORMULE:
Col Clôture = Col 5 + Col 7 - Col 11

DE PLUS: Les augmentations de dépréciations (Col 7) peuvent être CALCULEES comme:
   Augmentation = (Immobilisation Brute * Taux d'Amortissement * Mois/12)
   
Si on a les TAUX D'AMORTISSEMENT quelque part, on peut calculer les dépréciations!
""")

print("\n" + "="*130)
print("QUESTION CLES")
print("="*130)

print("""
1. Y A-T-IL UNE FEUILLE OU UNE ZONE AVEC LES TAUX D'AMORTISSEMENT?
   Exemple: Terrain = 0%, Bâtiment = 4%, Véhicules = 25%, etc.
   
2. SI OUI, ALORS ON PEUT:
   ✓ Remplir Col 4/10 (ouverture/clôture) avec Balance
   ✓ CALCULER Col 7 (augmentations dépréciation) = Col 10 * Taux * Jours/365
   ✓ CALCULER Col Clôture = Col 5 + Col 7 - Col 11
   
3. SI NON, ALORS ON REVIENT A L'APPROCHE CONSERVATIVE:
   ✓ Remplir UNIQUEMENT Col 10
   ✓ Laisser le reste vide
""")
