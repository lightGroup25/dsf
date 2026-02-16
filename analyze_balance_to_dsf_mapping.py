#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analyse des correspondances entre Balance de Compte et DSF
Basé sur les principes comptables OHADA/SYSCOHADA
"""
import openpyxl

print("=" * 120)
print("CORRESPONDANCE BALANCE - DSF - ANALYSE COMPLETE")
print("=" * 120)

# Charger les fichiers
dsf_template = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\templates\DSF Normal standard.xlsx"
balance_file = r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\input\BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx"

wb_dsf = openpyxl.load_workbook(dsf_template, data_only=False)
wb_balance = openpyxl.load_workbook(balance_file, data_only=True)

print("\n" + "=" * 120)
print("PARTIE 1: STRUCTURE ACTIF - NOTE 3A (Immobilisations)")
print("=" * 120)

if "NOTE 3A" in wb_dsf.sheetnames:
    ws_dsf = wb_dsf["NOTE 3A"]
    
    print("\nColonnes du DSF NOTE 3A:")
    print("-" * 120)
    
    # Afficher structure
    cols_structure = {
        4: ("MONTANT BRUTE A L'OUVERTURE", "Brut", "Valeur d'acquisition/ouverture"),
        5: ("ACQUISITIONS APPORTS CREATIONS", "Mvt", "Achats d'immobilisations pendant l'exercice"),
        6: ("VIREMENTS DE POSTE A POSTE", "Mvt", "Virements/transferts entre categories"),
        7: ("REEVALUATION PRATIQUEE AU COURS DE L'EXERCICE", "Mvt", "Reevaluations effectuees"),
        8: ("CESSIONS SCISSIONS HORS SERVICE", "Mvt", "Cessions/mise au rebut"),
        9: ("VIREMENTS", "Mvt", "Autres virements"),
        10: ("MONTANT BRUT A LA CLOTURE", "Brut", "Valeur brute = ouverture + mouvements"),
    }
    
    for col_num, (header, type_, meaning) in cols_structure.items():
        print("  Col {:2d}: {}".format(col_num, header))
        print("     Type: {}".format(type_))
        print("     Signification: {}\n".format(meaning))

print("\n" + "=" * 120)
print("PARTIE 2: STRUCTURE BALANCE DE COMPTE")
print("=" * 120)

ws_balance = wb_balance.active

print("\nHierarchie des colonnes BALANCE:")
print("-" * 120)

balance_structure = {
    "Identification": {
        1: "Numero de compte (ex: 21110, 21120, ...)",
        6: "Intitule du compte (ex: 'Terrains', 'Constructions', etc.)",
    },
    "DONNEES DISPONIBLES": {
        14: "Solde DEBIT final (31 décembre 2024)",
        27: "Solde CREDIT final (31 décembre 2024)",
        10: "Mouvements Débit du 31/12 (tous types mélangés)",
        13: "Mouvements Crédit du 31/12 (tous types mélangés)",
    },
}

for category, cols in balance_structure.items():
    print("\n  {}:".format(category))
    for col, desc in cols.items():
        if isinstance(col, str):
            print("     -> {}".format(desc))
        else:
            print("     Col {}: {}".format(col, desc))

print("\n" + "=" * 120)
print("PARTIE 3: LE PROBLEME FONDAMENTAL")
print("=" * 120)

print("""
DSF.NOTE 3A Col 5: "ACQUISITIONS APPORTS CREATIONS"
    ← Cherche dans Balance...
    ✗ La balance donne TOUS les mouvements melanges (col 10 ou 13)
    ✗ Aucune separation entre:
      - Acquisitions
      - Virements
      - Reevaluations
      - Cessions

CONCLUSION:
  Balance SYNTHÉTISE (total des mouvements)
  DSF DÉTAILLE (besoin du détail par type)
  
  -> ON NE PEUT PAS REMPLIR LE DÉTAIL SANS SOURCE ADDITIONNELLE !
""")

print("\n" + "=" * 120)
print("PARTIE 4: SOLUTION - CE QU'ON PEUT FAIRE")
print("=" * 120)

print("""
APPROCHE RECOMMANDEE: Remplir UNIQUEMENT ce qu'on peut faire honnetement
=======================================================================

DSF.NOTE 3A - Colonnes remplissables:
  
  [?] Col 4: "MONTANT BRUTE A L'OUVERTURE"
      ← Probleme: On n'a que la balance 2024
      ← Solution manquante: Besoin du bilan 2023 ou balance d'ouverture
      ← Recommandation: LAISSER VIDE (données manquantes)

  [!] Col 5-9: "Détails des mouvements"
      ← Balance ne les fournit pas en détail
      ← La balance n'a que TOTAL des mouvements (melanges)
      ← Recommandation: LAISSER VIDE (données manquantes)

  [OK] Col 10: "MONTANT BRUT A LA CLOTURE"
      ← Balance Col 14 (Débit) OU Col 27 (Crédit)
      ← C'est CE QU'ON PEUT FAIRE AVEC CERTITUDE !
      ← Recommandation: REMPLIR CE COLONNE !

RÉSULTAT HONNÊTE:
- Remplir UNIQUEMENT col 10 (clôture) = ~100-150 cellules
- Transparent sur ce qu'on ne peut pas faire
- États financiers fiables (pas d'estimation risquée)
""")

print("\n" + "=" * 120)
print("PARTIE 5: PROCHAINES ETAPES")
print("=" * 120)

print("""
QUESTION POUR VOUS:

Avez-vous d'autres fichiers/donnees (sourcedata)?
========================================================

A) Balance d'ouverture (01/01/2024)?
   -> Oui: On peut remplir Col 4
   -> Non: Col 4 reste vide

B) Détail des acquisitions, virements, réévaluations, cessions (par type)?
   -> Oui: On peut remplir Col 5-9
   -> Non: Col 5-9 restent vides

C) États de dépréciations/amortissements?
   -> Oui: On peut remplir NOTE 3B
   -> Non: NOTE 3B reste vide pour les détails

VOTRE CHOIX:

Option 1: APPROCHE CONSERVATIVE (Recommandée)
   - Remplir UNIQUEMENT les clôtures (col 10)
   - Résultat: ~150-200 cellules correctement remplies
   - Résultat: Honnête et fiable
   - Temps: 5 minutes

Option 2: RECHERCHER LES DONNEES MANQUANTES
   - Vérifier avec le client/comptable s'il a:
     * Balance d'ouverture
     * Détail des mouvements
     * États de dépréciations
   - Si les données existent: Code modifié pour les utiliser
   - Si données n'existent pas: Revenir à Option 1

QUELLE OPTION PREFEREZ-VOUS?
""")

wb_dsf.close()
wb_balance.close()
