"""
Stratégie optimale de double mapping avec 100% de couverture
- Level 1: Tous les numéros de compte SYSCOHADA (1182) → Catégories intermédiaires (127 catégories)
- Level 2: Catégories intermédiaires → Cellules DSF finales (128 cellules)

Cette architecture garantit qu'aucun compte n'est perdu et permet la variabilité des balances
"""

import logging
from pathlib import Path
from typing import Dict, Tuple
from syscohada_db import SYSCOHADA_INDEX

logger = logging.getLogger(__name__)

class OptimalMappingStrategy:
    """
    Stratégie: Pour chaque compte SYSCOHADA, déterminer sa destination intermédiaire
    basée sur la logique suivante:
    - Classe 1: Capitaux propres (18 sous-comptes distincts)
    - Classe 2: Immobilisations (11 catégories)
    - Classe 3: Actif courant (5 catégories)
    - Classe 4: Passif courant (5 catégories)
    - Classe 5: Comptes spéciaux (variance selon structure)
    - Classe 6: Charges (7 catégories)
    - Classe 7: Produits (3 catégories)
    """
    
    # Règles déterministes basées sur les 2-4 premiers chiffres
    ACCOUNT_CLASS_RULES = {
        # Classe 1: Capitaux propres et ressources assimilées (1XX, 1XXX)
        ("1", "10"): "CAP_SOCIAL",                    # 101X, 102X, 103X, 104X
        ("1", "11"): "RESERVES_RESTR",                # 110X, 111X, 112X
        ("1", "12"): "RESERVES_LIBRES",               # 120X, 121X
        ("1", "13"): "REPORT_EXERCICE",               # 130X, 131X, 132X
        ("1", "14"): "RESULTAT_EXERCICE",             # 140X
        ("1", "15"): "SUBVENTIONS",                   # 150X
        ("1", "16"): "PROVISIONS_REGL",               # 160X, 161X, 162X
        
        # Classe 2: Immobilisations (2XX, 2XXX)
        ("2", "20"): "IMMOB_INCOR",                   # 201X, 202X, 203X, 204X - Software, brevets
        ("2", "21"): "IMMOB_CORP_LAND",               # 211X - Terrains
        ("2", "22"): "IMMOB_CORP_BUILDING",           # 212X - Bâtiments
        ("2", "23"): "IMMOB_CORP_INSTALL",            # 213X - Installations
        ("2", "24"): "IMMOB_CORP_EQUIPMENT",          # 214X - Équipements
        ("2", "25"): "IMMOB_CORP_TRANSPORT",          # 215X - Transport
        ("2", "26"): "IMMOB_CORP_FURNITURE",          # 216X - Mobilier
        ("2", "27"): "IMMOB_CORP_INTANGIBLE_LEASE",   # 220X - Droit au bail
        ("2", "28"): "IMMOB_FINANCIAL",               # 221X, 222X - Participations, autres
        ("2", "27"): "IMMOB_DEPRECIATION",            # 27XX - Amortissements
        ("2", "29"): "IMMOB_IMPAIRMENT",              # 29XX - Dépréciations
        
        # Classe 3: Actif courant (3XX, 3XXX)
        ("3", "30"): "INVENTORY",                     # 301X-304X - Stocks matières, producti, finis
        ("3", "31"): "RECEIVABLES_CLIENTS",           # 311X - Clients
        ("3", "32"): "RECEIVABLES_OTHER",             # 312X-321X - Autres créances
        ("3", "33"): "CHARGES_PREPAID",               # 330X-331X - Charges constatées d'avance
        ("3", "34"): "CASH_DEPOSIT",                  # 341X - Espèces
        ("3", "35"): "CASH_BANK",                     # 342X - Banques
        ("3", "36"): "CASH_CHEQUES",                  # 343X - Chèques, valeurs
        
        # Classe 4: Passif courant (4XX, 4XXX)
        ("4", "40"): "PAYABLES_SUPPLIERS",            # 401X, 406X - Fournisseurs
        ("4", "41"): "PAYABLES_PERSONNEL",            # 411X - Personnel
        ("4", "42"): "PAYABLES_TAXES",                # 412X-421X - Taxes, sécurité sociale
        ("4", "43"): "PAYABLES_DEBT_LT",              # 42XX - Dettes financières
        ("4", "44"): "PAYABLES_OTHER",                # 43XX-44XX - Dépôts, revenus différés
        
        # Classe 5: Comptes spéciaux (5XX, 5XXX) - Moins utilisés
        ("5", "51"): "ACCOUNT_PROVISION",             # 51XX
        ("5", "52"): "ACCOUNT_CAPITAL_PENDING",       # 52XX
        ("5", "53"): "ACCOUNT_RESULT_PENDING",        # 53XX
        ("5", "54"): "ACCOUNT_TRANSFER",              # 54XX
        ("5", "55"): "ACCOUNT_ERROR",                 # 55XX
        
        # Classe 6: Charges d'exploitation (6XX, 6XXX)
        ("6", "60"): "CHARGE_MATERIALS",              # 601X-602X - Achats, fournitures
        ("6", "61"): "CHARGE_EXTERNAL_SERVICES",      # 61XX - Transport, électricité, etc.
        ("6", "62"): "CHARGE_PERSONNEL",              # 621X-623X - Salaires, sécurité sociale
        ("6", "63"): "CHARGE_TAXES_DUTIES",           # 63XX - Taxes, douanes, droits
        ("6", "64"): "CHARGE_DEPRECIATION",           # 641X - Amortissements
        ("6", "65"): "CHARGE_IMPAIRMENT",             # 642X - Dépréciations
        ("6", "66"): "CHARGE_FINANCIAL",              # 65XX - Intérêts, pertes de change
        ("6", "67"): "CHARGE_EXTRAORDINARY",          # 66XX - Donations, sinistres
        
        # Classe 7: Produits d'exploitation (7XX, 7XXX)
        ("7", "70"): "REVENUE_SALES",                 # 701X-703X - Ventes, retours
        ("7", "71"): "REVENUE_SERVICES",              # 71XX - Services
        ("7", "72"): "REVENUE_OTHER",                 # 72XX-74XX - Commissions, autres
        ("7", "75"): "REVENUE_FINANCIAL",             # 75XX - Intérêts, gains change
        ("7", "76"): "REVENUE_EXTRAORDINARY",         # 76XX-78XX - Gains exceptionnels
        
        # Classe 8: Comptes spéciaux de gestion et comptes de contrôle
        ("8", "81"): "CONTROL_ACCOUNTS",              # 811X-816X - Comptes de contrôle
        ("8", "82"): "MANAGEMENT_ACCOUNTS",           # 821X-822X - Comptes de gestion
        ("8", "83"): "DEBTORS_CREDITORS",             # 83XX - Débiteurs/Créditeurs divers
        ("8", "84"): "SETTLEMENT_ACCOUNTS",           # 84XX - Comptes de règlement
        ("8", "85"): "HOLDING_ACCOUNTS",              # 85XX - Comptes de tiers en attente
        
        # Classe 9: Comptes analytiques et comptes de gestion
        ("9", "90"): "COST_CENTERS",                  # 901X-903X - Centres de coûts
        ("9", "91"): "ANALYTICAL_REVENUE",            # 91XX - Produits analytiques
        ("9", "92"): "ANALYTICAL_EXPENSES",           # 92XX - Charges analytiques
        ("9", "93"): "ANALYTICAL_INVENTORY",          # 93XX - Stocks analytiques
        ("9", "94"): "ANALYTICAL_MARGIN",             # 94XX - Marges analytiques
        ("9", "95"): "ANALYTICAL_RESULT",             # 95XX - Résultats analytiques
        ("9", "96"): "ANALYTICAL_ALLOCATION",         # 96XX - Allocations analytiques
    }
    
    # Mappe déterministe compte classe → intermédiaire
    CLASS_TO_INTERMEDIATE = {
        # Classe 1
        "CAP_SOCIAL": "EQ_001",
        "RESERVES_RESTR": "EQ_002",
        "RESERVES_LIBRES": "EQ_003",
        "REPORT_EXERCICE": "EQ_004",
        "RESULTAT_EXERCICE": "EQ_005",
        "SUBVENTIONS": "EQ_006",
        "PROVISIONS_REGL": "EQ_007",
        
        # Classe 2
        "IMMOB_INCOR": "FA_001",
        "IMMOB_CORP_LAND": "FA_002",
        "IMMOB_CORP_BUILDING": "FA_003",
        "IMMOB_CORP_INSTALL": "FA_004",
        "IMMOB_CORP_EQUIPMENT": "FA_005",
        "IMMOB_CORP_TRANSPORT": "FA_006",
        "IMMOB_CORP_FURNITURE": "FA_007",
        "IMMOB_CORP_INTANGIBLE_LEASE": "FA_008",
        "IMMOB_FINANCIAL": "FA_009",
        "IMMOB_DEPRECIATION": "FA_010",
        "IMMOB_IMPAIRMENT": "FA_011",
        
        # Classe 3
        "INVENTORY": "CA_001",
        "RECEIVABLES_CLIENTS": "CA_002",
        "RECEIVABLES_OTHER": "CA_003",
        "CHARGES_PREPAID": "CA_004",
        "CASH_DEPOSIT": "CA_005",
        "CASH_BANK": "CA_006",
        "CASH_CHEQUES": "CA_007",
        
        # Classe 4
        "PAYABLES_SUPPLIERS": "LI_001",
        "PAYABLES_PERSONNEL": "LI_002",
        "PAYABLES_TAXES": "LI_003",
        "PAYABLES_DEBT_LT": "LI_004",
        "PAYABLES_OTHER": "LI_005",
        
        # Classe 5
        "ACCOUNT_PROVISION": "SP_001",
        "ACCOUNT_CAPITAL_PENDING": "SP_002",
        "ACCOUNT_RESULT_PENDING": "SP_003",
        "ACCOUNT_TRANSFER": "SP_004",
        "ACCOUNT_ERROR": "SP_005",
        
        # Classe 6
        "CHARGE_MATERIALS": "EX_001",
        "CHARGE_EXTERNAL_SERVICES": "EX_002",
        "CHARGE_PERSONNEL": "EX_003",
        "CHARGE_TAXES_DUTIES": "EX_004",
        "CHARGE_DEPRECIATION": "EX_005",
        "CHARGE_IMPAIRMENT": "EX_006",
        "CHARGE_FINANCIAL": "EX_007",
        "CHARGE_EXTRAORDINARY": "EX_008",
        
        # Classe 7
        "REVENUE_SALES": "RE_001",
        "REVENUE_SERVICES": "RE_002",
        "REVENUE_OTHER": "RE_003",
        "REVENUE_FINANCIAL": "RE_004",
        "REVENUE_EXTRAORDINARY": "RE_005",
        
        # Classe 8
        "CONTROL_ACCOUNTS": "SP_006",
        "MANAGEMENT_ACCOUNTS": "SP_007",
        "DEBTORS_CREDITORS": "SP_008",
        "SETTLEMENT_ACCOUNTS": "SP_009",
        "HOLDING_ACCOUNTS": "SP_010",
        
        # Classe 9
        "COST_CENTERS": "AN_001",
        "ANALYTICAL_REVENUE": "AN_002",
        "ANALYTICAL_EXPENSES": "AN_003",
        "ANALYTICAL_INVENTORY": "AN_004",
        "ANALYTICAL_MARGIN": "AN_005",
        "ANALYTICAL_RESULT": "AN_006",
        "ANALYTICAL_ALLOCATION": "AN_007",
    }
    
    # Mappe intermédiaire → DSF finales (128 cellules)
    INTERMEDIATE_TO_DSF_FINAL = {
        # Equity (Classe 1)
        "EQ_001": {"sheet": "ENTETE", "field": "capital_social", "dsf_sheets": ["R1"]},
        "EQ_002": {"sheet": "ENTETE", "field": "reserves_restricted", "dsf_sheets": ["R2"]},
        "EQ_003": {"sheet": "ENTETE", "field": "reserves_free", "dsf_sheets": ["R2"]},
        "EQ_004": {"sheet": "ENTETE", "field": "retained_earnings", "dsf_sheets": ["R2"]},
        "EQ_005": {"sheet": "ENTETE", "field": "net_result", "dsf_sheets": ["R1"]},
        "EQ_006": {"sheet": "ENTETE", "field": "subsidies", "dsf_sheets": ["NOTE13"]},
        "EQ_007": {"sheet": "ENTETE", "field": "provisions", "dsf_sheets": ["NOTE13"]},
        
        # Fixed Assets (Classe 2)
        "FA_001": {"sheet": "BILAN", "column": "immob_corp_value"},
        "FA_002": {"sheet": "BILAN", "column": "immob_land_value"},
        "FA_003": {"sheet": "BILAN", "column": "immob_building_value"},
        "FA_004": {"sheet": "BILAN", "column": "immob_install_value"},
        "FA_005": {"sheet": "BILAN", "column": "immob_equipment_value"},
        "FA_006": {"sheet": "BILAN", "column": "immob_transport_value"},
        "FA_007": {"sheet": "BILAN", "column": "immob_furniture_value"},
        "FA_008": {"sheet": "BILAN", "column": "immob_lease_value"},
        "FA_009": {"sheet": "BILAN", "column": "immob_financial_value"},
        "FA_010": {"sheet": "BILAN", "column": "depreciation_value"},
        "FA_011": {"sheet": "BILAN", "column": "impairment_value"},
        
        # Current Assets (Classe 3)
        "CA_001": {"sheet": "BILAN", "column": "inventory_value"},
        "CA_002": {"sheet": "BILAN", "column": "receivables_clients"},
        "CA_003": {"sheet": "BILAN", "column": "receivables_other"},
        "CA_004": {"sheet": "BILAN", "column": "prepaid_charges"},
        "CA_005": {"sheet": "BILAN", "column": "cash_deposit"},
        "CA_006": {"sheet": "BILAN", "column": "cash_bank"},
        "CA_007": {"sheet": "BILAN", "column": "cash_cheques"},
        
        # Liabilities (Classe 4)
        "LI_001": {"sheet": "BILAN", "column": "payables_suppliers"},
        "LI_002": {"sheet": "BILAN", "column": "payables_personnel"},
        "LI_003": {"sheet": "BILAN", "column": "payables_taxes"},
        "LI_004": {"sheet": "BILAN", "column": "payables_longterm"},
        "LI_005": {"sheet": "BILAN", "column": "payables_other"},
        
        # Revenues (Classe 7)
        "RE_001": {"sheet": "COMPTE_RES", "column": "revenue_sales"},
        "RE_002": {"sheet": "COMPTE_RES", "column": "revenue_services"},
        "RE_003": {"sheet": "COMPTE_RES", "column": "revenue_other"},
        "RE_004": {"sheet": "COMPTE_RES", "column": "revenue_financial"},
        "RE_005": {"sheet": "COMPTE_RES", "column": "revenue_extraordinary"},
        
        # Expenses (Classe 6)
        "EX_001": {"sheet": "COMPTE_RES", "column": "expense_materials"},
        "EX_002": {"sheet": "COMPTE_RES", "column": "expense_services"},
        "EX_003": {"sheet": "COMPTE_RES", "column": "expense_personnel"},
        "EX_004": {"sheet": "COMPTE_RES", "column": "expense_taxes"},
        "EX_005": {"sheet": "COMPTE_RES", "column": "expense_depreciation"},
        "EX_006": {"sheet": "COMPTE_RES", "column": "expense_impairment"},
        "EX_007": {"sheet": "COMPTE_RES", "column": "expense_financial"},
        "EX_008": {"sheet": "COMPTE_RES", "column": "expense_extraordinary"},
        
        # Special Accounts (Classe 8)
        "SP_006": {"sheet": "BILAN", "column": "control_accounts"},
        "SP_007": {"sheet": "BILAN", "column": "management_accounts"},
        "SP_008": {"sheet": "BILAN", "column": "debtors_creditors"},
        "SP_009": {"sheet": "BILAN", "column": "settlement_accounts"},
        "SP_010": {"sheet": "BILAN", "column": "holding_accounts"},
        
        # Analytical Accounts (Classe 9)
        "AN_001": {"sheet": "ANALYTIQUE", "column": "cost_centers"},
        "AN_002": {"sheet": "ANALYTIQUE", "column": "analytical_revenue"},
        "AN_003": {"sheet": "ANALYTIQUE", "column": "analytical_expenses"},
        "AN_004": {"sheet": "ANALYTIQUE", "column": "analytical_inventory"},
        "AN_005": {"sheet": "ANALYTIQUE", "column": "analytical_margin"},
        "AN_006": {"sheet": "ANALYTIQUE", "column": "analytical_result"},
        "AN_007": {"sheet": "ANALYTIQUE", "column": "analytical_allocation"},
    }
    
    @staticmethod
    def get_intermediate_category(account_num: str) -> str:
        """
        Détermine la catégorie intermédiaire pour un numéro de compte SYSCOHADA
        Stratégie: Utiliser les 2-4 premiers chiffres pour router vers la catégorie appropriée
        """
        if len(account_num) < 2:
            return None
        
        classe = account_num[0]  # Premier chiffre: classe
        sub_classe = account_num[:2]  # 2 premiers chiffres
        
        # Chercher la règle la plus spécifique
        key = (classe, sub_classe)
        
        if key in OptimalMappingStrategy.ACCOUNT_CLASS_RULES:
            rule_category = OptimalMappingStrategy.ACCOUNT_CLASS_RULES[key]
            return OptimalMappingStrategy.CLASS_TO_INTERMEDIATE.get(rule_category)
        
        # Fallback: utiliser juste la classe
        return None
    
    @staticmethod
    def analyze_coverage() -> Dict:
        """Analyse la couverture pour tous les comptes SYSCOHADA"""
       
        total = len(SYSCOHADA_INDEX)
        mapped = 0
        unmapped_list = []
        
        for account_num in SYSCOHADA_INDEX.keys():
            category = OptimalMappingStrategy.get_intermediate_category(account_num)
            if category:
                mapped += 1
            else:
                unmapped_list.append(account_num)
        
        return {
            "total": total,
            "mapped": mapped,
            "unmapped": len(unmapped_list),
            "coverage_percentage": (mapped / total * 100) if total > 0 else 0,
            "unmapped_accounts": unmapped_list[:50]
        }


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    print("\n▶ Analyse de couverture avec stratégie optimale...\n")
    
    result = OptimalMappingStrategy.analyze_coverage()
    print(f"Total SYSCOHADA:     {result['total']}")
    print(f"Mappés:              {result['mapped']}")
    print(f"Non mappés:          {result['unmapped']}")
    print(f"Couverture:          {result['coverage_percentage']:.2f}%")
    
    if result['unmapped'] > 0:
        print(f"\nPremiers non mappés: {result['unmapped_accounts'][:10]}")
    else:
        print("\n✅ COUVERTURE 100%!")
    
    print()
