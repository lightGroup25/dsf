#!/usr/bin/env python3
"""
Analyse des correspondances entre Balance de Compte et DSF
Basé sur les principes comptables OHADA/SYSCOHADA
"""
import openpyxl
from pathlib import Path

print("=" * 120)
print("CORRESPONDANCE BALANCE - DSF - ANALYSE COMPLETE")
print("=" * 120)

# Charger le fichier template DSF
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
        6: ("VIREMENTS DE POSTE A POSTE", "Mvt", "Virements/transferts entre catégories"),
        7: ("SUITE A UNE REEVALUATION PRATIQUEE AU COURS DE L'EXERCICE", "Mvt", "Réévaluations effectuées"),
        8: ("CESSIONS SCISSIONSHORS SERVICE", "Mvt", "Cessions/mise au rebut"),
        9: ("VIREMENTS DE POSTE A POSTE", "Mvt", "Autres virements"),
        10: ("MONTANT BRUT A LA CLOTURE", "Brut", "Valeur brute de clôture = ouverture + mouvements"),
    }
    
    for col_num, (header, type_, meaning) in cols_structure.items():
        print(f"\n  Col {col_num}: {header}")
        print(f"    Type: {type_}")
        print(f"    Signification: {meaning}")

print("\n" + "=" * 120)
print("PARTIE 2: STRUCTURE BALANCE")
print("=" * 120)

ws_balance = wb_balance.active
print(f"\nSheets balance: {wb_balance.sheetnames}")

print("\nHiérarchie des colonnes BALANCE:")
print("-" * 120)

balance_structure = {
    "Identification": {
        1: "Numéro de compte (ex: 21110, 21120, ...)",
        6: "Intitulé du compte (ex: 'Terrains', 'Constructions', 'Immobilisations corporelles', etc.)",
    },
    "Soldes Initiaux (1er janvier 2024)": {
        "Note": "Supposé = Col 23 Débit ou Col 26 Crédit (vérifié dans votre analyse)",
    },
    "Mouvements au 31/12/23 (Exercice 2024)": {
        10: "Débit - transactions débitrices",
        13: "Crédit - transactions créditrices",
        "Note": "Type de mouvement: NON DÉTAILLÉ (acquisition? transfert? réévaluation? tous mélangés?)",
    },
    "Mouvements (colonnes ??)": {
        17: "Débit - mouvements intermédiaires(?)",
        20: "Crédit - mouvements intermédiaires(?)",
        "Note": "Fonction peu claire",
    },
    "Soldes Finaux (31 décembre 2024)": {
        14: "Débit (solde final côté débit)",
        27: "Crédit (solde final côté crédit)",
        "Note": "C'est le solde final de chaque compte après tous les mouvements",
    }
}

for category, cols in balance_structure.items():
    print(f"\n  {category}:")
    for col, desc in cols.items():
        if isinstance(col, str):
            print(f"    → {desc}")
        else:
            print(f"    Col {col}: {desc}")

print("\n" + "=" * 120)
print("PARTIE 3: LE PROBLÈME")
print("=" * 120)

print("""
DSF.NOTE 3A Col 5: "ACQUISITIONS APPORTS CREATIONS"
    ↓ Cherche dans Balance...
    ↗ La balance donne TOUS les mouvements mélangés (col 10/13)
    ✗ Aucune séparation entre:
      - Acquisitions (acquisitions d'immobilisations)
      - Virements (transferts entre categories)
      - Réévaluations (revaluations comptables)
      - Cessions (ventes/sorties d'immobilisations)

CONCLUSION: 
    La balance SYNTHÉTISE, le DSF DÉTAILLE
    → Balance col 10/13 = TOTAL de tous les mouvements de l'année
    → DSF col 5-9 = DEMANDE le détail par TYPE de mouvement
    
    → ON NE PEUT PAS REMPLIR le détail sans source additionnelle!
""")

print("\n" + "=" * 120)
print("PARTIE 4: QU'EST-CE QU'ON DEVRAIT FAIRE?")
print("=" * 120)

print("""
OPTION A: APPROCHE MINIMALISTE (Recommandée)
==========================================
Remplir UNIQUEMENT les colonnes qu'on a les données pour:

1. DSF Col 4 "MONTANT BRUTE A L'OUVERTURE"
   ← À chercher dans: Balance colonnes 12/13? (soldes initiaux)
   → Ou à calculer: solde final 2023 = soldes initiaux 2024
   → DÉFI: On n'a que balance 2024... pas les données d'ouverture!

2. DSF Col 10 "MONTANT BRUT A LA CLOTURE"
   ← Balance Col 14 (Débit) ou Col 27 (Crédit)
   → C'est ce qu'on peut faire avec certitude!

3. DSF Col 5-9 "Détails des mouvements"
   ← Balance ne les fournit pas en détail
   → LAISSER VIDE ✓ (Honnête)

RÉSULTAT HONNÊTE:
- ~200-300 cellules remplies (colonnes 4 + 10)
- Transparent sur ce qu'on ne peut pas faire
- États financiers fiables

OPTION B: VÉRIFier si d'autres données existent
===============================================
Questions pour vous:
1. Avez-vous un fichier "Balance d'ouverture" (01/01/2024)?
2. Avez-vous des journaux détaillés (acquisitions, transferts, réévaluations, cessions)?
3. Avez-vous des états d'amortissement (dépréciations)?
4. Ces fichiers contiennent-ils le détail par type de mouvement?

Si OUI → On pourrait remplir plus complètement!
Si NON → Option A (minimaliste mais honnête)
""")

print("\n" + "=" * 120)
print("PARTIE 5: NOTES COMPLÉMENTAIRES - NOTE 3B (Dépréciations)")
print("=" * 120)

print("""
NOTE 3B a une structure SIMILAIRE mais pour les dépréciations:
- Col 4: Dépréciation d'ouverture
- Col 5-9: Mouvements de dépréciations (augmentations, reprises, transferts, cessions)
- Col 10: Dépréciations de clôture

MÊME PROBLÈME: Balance ne donne PAS le détail des dépréciations!

Seul ce qu'on peut faire:
✓ Col 4: Récupérer les amortissements cumulés du bilan d'ouverture (si disponible)
✓ Col 10: Calculer = Tous les amortissements cumulés clôture (à chercher ailleurs)
✗ Col 5-9: Impossible (pas de détail)
""")

print("\n" + "=" * 120)
print("PARTIE 6: POINTS D'ENTRÉE - Pour les autresNOTE (4, 5, etc...)")
print("=" * 120)

print("""
Appliquer la MÊME LOGIQUE à chaque NOTE:

NOTE 4 (Stocks):
- Col 4: Stock d'ouverture → Laisser vide (pas d'inventaire d'ouverture)
- Col 5-9: Variations stocks → Laisser vide (pas disponible)
- Col 10: Stock de clôture → Balance compte d'inventaire? Oui si existe

NOTE 5 (Créances):
- Col 4: Créances d'ouverture → À chercher ailleurs
- Col 5-9: Mouvements créances → À chercher ailleurs
- Col 10: Créances de clôture → Balance compte créances clients

...et ainsi de suite

PATTERN GÉNÉRAL:
Pour chaque NOTE × chaque colonne:
- Si c'est une CLÔTURE → Chercher dans Balance soldes finaux (cols 14/27)
- Si c'est une OUVERTURE → Non disponible actuellement
- Si c'est un DÉTAIL de mouvement → Non disponible dans cette balance
""")

wb_dsf.close()
wb_balance.close()

print("\n" + "=" * 120)
print("RECOMMANDATION: ACTION À PRENDRE")
print("=" * 120)

print("""
ÉTAPE 1: Confirmez votre préférence
☐ Option A: Remplir UNIQUEMENT les clôtures (col 10 partout) = ~300 cellules, honnête
☐ Option B: Vérifier si d'autres sources de données existent (contact client/comptable)

ÉTAPE 2: Si Option A
- Modifier le code pour CIBLER UNIQUEMENT les colonnes de clôture
- Mapper DSF col 10 → Balance cols 14/27
- Résultat: 3-5 min pour générer le fichier
- Résultat: ~200-300 cellules remplies correctement

ÉTAPE 3: Tester et valider
- Vérifier quelques comptes manuellement
- Créer rapport final avec explications

Voulez-vous procéder avec Option A (recommandée)?
""")
