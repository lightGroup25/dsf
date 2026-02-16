"""
Create the inverse mapping: Intermediate Categories → DSF Line Codes
This bridges: SYSCOHADA → Intermediate → DSF cell locations
"""

import json

# Create category-to-DSF-line mapping
# Based on analyzing the structure and the accounting logic

CATEGORY_TO_DSF_LINES = {
    # CLASS 1 - EQUITY (Capitaux propres)
    "EQ_CAPITAL": {
        "sheet": "BILAN PAYSAGE",
        "code": "CA",           # Capital line on BILAN PAYSAGE
        "label": "Capital",
        "location": {"col_ref": "I", "label_row": 12, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif"
    },
    "EQ_RESERVES": {
        "sheet": "BILAN PAYSAGE",
        "code": "CB",           # Capital adjustments
        "label": "Réserves",
        "location": {"col_ref": "I", "label_row": 16, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif"
    },
    "EQ_EARNINGS": {
        "sheet": "BILAN PAYSAGE",
        "code": "CH",           # Report à nouveau
        "label": "Report à nouveau",
        "location": {"col_ref": "I", "label_row": 18, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif"
    },
    "EQ_RESULT": {
        "sheet": "BILAN PAYSAGE",
        "code": "CJ",           # Résultat net
        "label": "Résultat net de l'exercice",
        "location": {"col_ref": "I", "label_row": 19, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif"
    },
    "EQ_SUBSIDY": {
        "sheet": "BILAN PAYSAGE",
        "code": "CL",           # Subventions
        "label": "Subventions d'investissement",
        "location": {"col_ref": "I", "label_row": 20, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif"
    },
    "EQ_PROVISIONS": {
        "sheet": "BILAN PAYSAGE",
        "code": "CM",           # Provisions
        "label": "Provisions réglementées",
        "location": {"col_ref": "I", "label_row": 21, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif"
    },
    
    # CLASS 2 - FIXED ASSETS (Immobilisations)
    "FA_INTANGIBLE": {
        "sheet": "BILAN PAYSAGE",
        "code": "AD",
        "label": "IMMOBILISATIONS INCORPORELLES",
        "location": {"col_ref": "B", "label_row": 12, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif",
        "note_ref": "3"
    },
    "FA_LAND": {
        "sheet": "BILAN PAYSAGE",
        "code": "AJ",
        "label": "Terrains",
        "location": {"col_ref": "B", "label_row": 18, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif"
    },
    "FA_BUILDING": {
        "sheet": "BILAN PAYSAGE",
        "code": "AK",
        "label": "Bâtiments",
        "location": {"col_ref": "B", "label_row": 19, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif"
    },
    "FA_INSTALL": {
        "sheet": "BILAN PAYSAGE",
        "code": "AL",
        "label": "Aménagements, agencements et installations",
        "location": {"col_ref": "B", "label_row": 20, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif"
    },
    "FA_EQUIPMENT": {
        "sheet": "BILAN PAYSAGE",
        "code": "AM",
        "label": "Matériel, mobilier et actifs biologiques",
        "location": {"col_ref": "B", "label_row": 21, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif"
    },
    "FA_TRANSPORT": {
        "sheet": "BILAN PAYSAGE",
        "code": "AN",
        "label": "Matériel de transport",
        "location": {"col_ref": "B", "label_row": 22, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif"
    },
    "FA_FINANCIAL": {
        "sheet": "BILAN PAYSAGE",
        "code": "AQ",
        "label": "IMMOBILISATION FINANCIERES",
        "location": {"col_ref": "B", "label_row": 24, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif",
        "note_ref": "4"
    },
    "FA_DEPRECIATION": {
        "sheet": "BILAN PAYSAGE",
        "code": "AZ",           # TOTAL IMMOBILISATIONS - calculated
        "label": "TOTAL ACTIF IMMOBILISE",
        "location": {"col_ref": "B", "label_row": 27, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif",
        "is_total": True
    },
    
    # CLASS 3 - CURRENT ASSETS (Actif circulant)
    "CA_INVENTORY": {
        "sheet": "BILAN PAYSAGE",
        "code": "BB",
        "label": "Stocks et encours",
        "location": {"col_ref": "B", "label_row": 29, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif",
        "note_ref": "6"
    },
    "CA_RECEIVABLES": {
        "sheet": "BILAN PAYSAGE",
        "code": "BC",
        "label": "Créances et emplois assimilés",
        "location": {"col_ref": "B", "label_row": 30, "values_cols": ["D", "E", "F", "G"]},
        "type": "balance_sheet_actif"
    },
    "CA_CASH": {
        "sheet": "BILAN PAYSAGE",
        "code": "BT",
        "label": "TOTAL TRESORERIE - ACTIF",
        "location": {"col_ref": "B", "label_row": 38, "values_cols": ["D", "F", "G"]},
        "type": "balance_sheet_actif",
        "is_total": True
    },
    
    # CLASS 4 - LIABILITIES (Passif courant)
    "LI_SUPPLIERS": {
        "sheet": "BILAN PAYSAGE",
        "code": "DJ",
        "label": "Fournisseurs d'exploitation",
        "location": {"col_ref": "I", "label_row": 30, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif",
        "note_ref": "17"
    },
    "LI_PERSONNEL": {
        "sheet": "BILAN PAYSAGE",
        "code": "DK",
        "label": "Dettes fiscales et sociales",
        "location": {"col_ref": "I", "label_row": 31, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif",
        "note_ref": "18"
    },
    "LI_TAXES": {
        "sheet": "BILAN PAYSAGE",
        "code": "DK",
        "label": "Dettes fiscales et sociales",
        "location": {"col_ref": "I", "label_row": 31, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif",
        "note_ref": "18"
    },
    "LI_DEBT_LT": {
        "sheet": "BILAN PAYSAGE",
        "code": "DA",
        "label": "Emprunts et dettes financières diverses",
        "location": {"col_ref": "I", "label_row": 23, "values_cols": ["K", "L"]},
        "type": "balance_sheet_passif",
        "note_ref": "16"
    },
    
    # CLASS 6 - EXPENSES (Charges)
    "EX_MATERIALS": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "RC",
        "label": "Achats de matières premières",
        "location": {"col_ref": "B", "label_row": 24, "values_cols": ["E", "F"]},
        "type": "income_statement"
    },
    "EX_SERVICES": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "RH",
        "label": "Services extérieurs",
        "location": {"col_ref": "B", "label_row": 29, "values_cols": ["E", "F"]},
        "type": "income_statement",
        "note_ref": "24"
    },
    "EX_PERSONNEL": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "RK",
        "label": "Charges de personnel",
        "location": {"col_ref": "B", "label_row": 33, "values_cols": ["E", "F"]},
        "type": "income_statement",
        "note_ref": "27"
    },
    "EX_TAXES": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "RI",
        "label": "Impôts et taxes",
        "location": {"col_ref": "B", "label_row": 30, "values_cols": ["E", "F"]},
        "type": "income_statement",
        "note_ref": "25"
    },
    "EX_DEPRECIATION": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "RL",
        "label": "Dotations aux amortissements",
        "location": {"col_ref": "B", "label_row": 36, "values_cols": ["E", "F"]},
        "type": "income_statement"
    },
    "EX_FINANCIAL": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "RM",
        "label": "Frais financiers et charges assimilées",
        "location": {"col_ref": "B", "label_row": 41, "values_cols": ["E", "F"]},
        "type": "income_statement",
        "note_ref": "29"
    },
    
    # CLASS 7 - REVENUES (Produits)
    "RE_SALES": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "TA",
        "label": "Vente de marchandises",
        "location": {"col_ref": "B", "label_row": 11, "values_cols": ["E", "F"]},
        "type": "income_statement",
        "note_ref": "21"
    },
    "RE_SERVICES": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "TC",
        "label": "Travaux, services vendus",
        "location": {"col_ref": "B", "label_row": 16, "values_cols": ["E", "F"]},
        "type": "income_statement",
        "note_ref": "21"
    },
    "RE_FINANCIAL": {
        "sheet": "COMPTE DE RESULTAT",
        "code": "TK",
        "label": "Revenus financiers et assimilés",
        "location": {"col_ref": "B", "label_row": 38, "values_cols": ["E", "F"]},
        "type": "income_statement",
        "note_ref": "29"
    },
}

if __name__ == "__main__":
    print(f"\n✅ Created category-to-DSF mapping")
    print(f"   Total mappings: {len(CATEGORY_TO_DSF_LINES)}")
    print(f"\n   Sheets covered:")
    sheets = set(v["sheet"] for v in CATEGORY_TO_DSF_LINES.values())
    for sheet in sorted(sheets):
        count = sum(1 for v in CATEGORY_TO_DSF_LINES.values() if v["sheet"] == sheet)
        print(f"     - {sheet}: {count}")
    
    print(f"\n   Example mapping:")
    for cat, info in list(CATEGORY_TO_DSF_LINES.items())[:3]:
        print(f"     {cat} → Row {info['location']['label_row']}, Sheet: {info['sheet']}")
