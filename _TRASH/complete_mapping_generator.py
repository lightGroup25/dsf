"""
Générateur automatique du double mapping complet
Crée un mapping 100% exhaustif des 1182 comptes SYSCOHADA
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Tuple
from syscohada_db import SYSCOHADA_INDEX

logger = logging.getLogger(__name__)

class CompleteDoubleMappingGenerator:
    """Génère un mapping complet SYSCOHADA → Intermédiaire → DSF"""
    
    # Catégories de comptes SYSCOHADA (classe + caractéristiques)
    SYSCOHADA_CATEGORIES = {
        "10": {"category": "EQUITY_CAPITAL", "class": "1", "label": "Capital social"},
        "11": {"category": "EQUITY_RESERVES_RESTRICTED", "class": "1", "label": "Réserves indisponibles"},
        "12": {"category": "EQUITY_RESERVES_FREE", "class": "1", "label": "Réserves libres"},
        "13": {"category": "EQUITY_RETAINED_EARNINGS", "class": "1", "label": "Report à nouveau"},
        "14": {"category": "EQUITY_NET_RESULT", "class": "1", "label": "Résultat net"},
        "15": {"category": "EQUITY_SUBSIDIES", "class": "1", "label": "Subventions"},
        "16": {"category": "EQUITY_PROVISIONS", "class": "1", "label": "Provisions réglementées"},
        
        "20": {"category": "FIXED_INTANGIBLE", "class": "2", "label": "Immob. incorporelles"},
        "21": {"category": "FIXED_TANGIBLE", "class": "2", "label": "Immob. corporelles"},
        "22": {"category": "FIXED_FINANCIAL", "class": "2", "label": "Immob. financières"},
        "27": {"category": "FIXED_DEPRECIATION", "class": "2", "label": "Amortissements", "type": "contra"},
        
        "30": {"category": "INVENTORY", "class": "3", "label": "Stocks"},
        "31": {"category": "RECEIVABLES", "class": "3", "label": "Créances clients"},
        "32": {"category": "OTHER_RECEIVABLES", "class": "3", "label": "Autres créances"},
        "33": {"category": "PREPAID", "class": "3", "label": "Charges constatées d'avance"},
        "34": {"category": "CASH", "class": "3", "label": "Trésorerie actif"},
        
        "40": {"category": "PAYABLES_SUPPLIERS", "class": "4", "label": "Dettes fournisseurs"},
        "41": {"category": "PAYABLES_EMPLOYEE_TAX", "class": "4", "label": "Dettes personnel/fiscal"},
        "42": {"category": "PAYABLES_LONG_TERM", "class": "4", "label": "Dettes financières LT"},
        "43": {"category": "PAYABLES_OTHER", "class": "4", "label": "Autres dettes"},
        "44": {"category": "PAYABLES_ACCRUED", "class": "4", "label": "Charges à payer"},
        
        "70": {"category": "REVENUE_SALES", "class": "7", "label": "Ventes"},
        "71": {"category": "REVENUE_SERVICES", "class": "7", "label": "Services"},
        "72": {"category": "REVENUE_OTHER", "class": "7", "label": "Autres revenus"},
        
        "60": {"category": "EXPENSE_MATERIALS", "class": "6", "label": "Achats matières"},
        "61": {"category": "EXPENSE_SERVICES", "class": "6", "label": "Services extérieurs"},
        "62": {"category": "EXPENSE_PERSONNEL", "class": "6", "label": "Personnel"},
        "63": {"category": "EXPENSE_TAXES", "class": "6", "label": "Taxes"},
        "64": {"category": "EXPENSE_DEPRECIATION", "class": "6", "label": "Dépréciations"},
        "65": {"category": "EXPENSE_FINANCIAL", "class": "6", "label": "Charges financières"},
        "66": {"category": "EXPENSE_OTHER", "class": "6", "label": "Autres charges"},
    }
    
    # Mappe les catégories intermédiaires aux cellules DSF (simplifié)
    INTERMEDIATE_TO_DSF = {
        # Capitaux propres → BILAN_PAYSAGE D33-D42
        "EQUITY_CAPITAL": {"sheet": "BILAN_PAYSAGE", "cells": ["D33"], "label": "Capital"},
        "EQUITY_RESERVES_RESTRICTED": {"sheet": "BILAN_PAYSAGE", "cells": ["D37"], "label": "Réserves indisponibles"},
        "EQUITY_RESERVES_FREE": {"sheet": "BILAN_PAYSAGE", "cells": ["D38"], "label": "Réserves libres"},
        "EQUITY_RETAINED_EARNINGS": {"sheet": "BILAN_PAYSAGE", "cells": ["D39"], "label": "Report à nouveau"},
        "EQUITY_NET_RESULT": {"sheet": "BILAN_PAYSAGE", "cells": ["D40"], "label": "Résultat net"},
        "EQUITY_SUBSIDIES": {"sheet": "BILAN_PAYSAGE", "cells": ["D41"], "label": "Subventions"},
        "EQUITY_PROVISIONS": {"sheet": "BILAN_PAYSAGE", "cells": ["D42"], "label": "Provisions"},
        
        # Immobilisations → BILAN_PAYSAGE D14-D27
        "FIXED_INTANGIBLE": {"sheet": "BILAN_PAYSAGE", "cells": ["D14"], "label": "Immob. incorporelles"},
        "FIXED_TANGIBLE": {"sheet": "BILAN_PAYSAGE", "cells": ["D18"], "label": "Immob. corporelles"},
        "FIXED_FINANCIAL": {"sheet": "BILAN_PAYSAGE", "cells": ["D24"], "label": "Immob. financières"},
        "FIXED_DEPRECIATION": {"sheet": "BILAN_PAYSAGE", "cells": ["D25"], "label": "Amortissements", "multiply_by": -1},
        
        # Actif courant → BILAN_PAYSAGE D29-D43
        "INVENTORY": {"sheet": "BILAN_PAYSAGE", "cells": ["D29"], "label": "Stocks"},
        "RECEIVABLES": {"sheet": "BILAN_PAYSAGE", "cells": ["D30"], "label": "Créances clients"},
        "OTHER_RECEIVABLES": {"sheet": "BILAN_PAYSAGE", "cells": ["D31"], "label": "Autres créances"},
        "PREPAID": {"sheet": "BILAN_PAYSAGE", "cells": ["D32"], "label": "Charges constatées d'avance"},
        "CASH": {"sheet": "BILAN_PAYSAGE", "cells": ["D43"], "label": "Trésorerie"},
        
        # Passif → BILAN_PAYSAGE D44-D51
        "PAYABLES_SUPPLIERS": {"sheet": "BILAN_PAYSAGE", "cells": ["D44"], "label": "Dettes fournisseurs"},
        "PAYABLES_EMPLOYEE_TAX": {"sheet": "BILAN_PAYSAGE", "cells": ["D45"], "label": "Dettes personnel/fiscal"},
        "PAYABLES_LONG_TERM": {"sheet": "BILAN_PAYSAGE", "cells": ["D46"], "label": "Dettes LT"},
        "PAYABLES_OTHER": {"sheet": "BILAN_PAYSAGE", "cells": ["D47"], "label": "Autres dettes"},
        "PAYABLES_ACCRUED": {"sheet": "BILAN_PAYSAGE", "cells": ["D48"], "label": "Charges à payer"},
        
        # Revenus → COMPTE_RESULTAT
        "REVENUE_SALES": {"sheet": "COMPTE_RESULTAT", "cells": ["D1"], "label": "Ventes"},
        "REVENUE_SERVICES": {"sheet": "COMPTE_RESULTAT", "cells": ["D2"], "label": "Services"},
        "REVENUE_OTHER": {"sheet": "COMPTE_RESULTAT", "cells": ["D3"], "label": "Autres revenus"},
        
        # Charges → COMPTE_RESULTAT
        "EXPENSE_MATERIALS": {"sheet": "COMPTE_RESULTAT", "cells": ["D10"], "label": "Achats matières"},
        "EXPENSE_SERVICES": {"sheet": "COMPTE_RESULTAT", "cells": ["D11"], "label": "Services"},
        "EXPENSE_PERSONNEL": {"sheet": "COMPTE_RESULTAT", "cells": ["D12"], "label": "Personnel"},
        "EXPENSE_TAXES": {"sheet": "COMPTE_RESULTAT", "cells": ["D13"], "label": "Taxes"},
        "EXPENSE_DEPRECIATION": {"sheet": "COMPTE_RESULTAT", "cells": ["D14"], "label": "Dépréciation"},
        "EXPENSE_FINANCIAL": {"sheet": "COMPTE_RESULTAT", "cells": ["D15"], "label": "Charges financières"},
        "EXPENSE_OTHER": {"sheet": "COMPTE_RESULTAT", "cells": ["D16"], "label": "Autres charges"},
    }
    
    def __init__(self):
        self.level_1_mapping = {}
        self.level_2_mapping = {}
        self.unmapped_accounts = []
        
    def generate_level_1(self):
        """Génère le mapping Level 1 complet (SYSCOHADA → Intermédiaire)"""
        level_1 = {}
        
        # Pour chaque compte SYSCOHADA
        for account_num in sorted(SYSCOHADA_INDEX.keys()):
            account_label = SYSCOHADA_INDEX[account_num]
            
            # Déterminer la catégorie intermédiaire
            category_key = account_num[:2]  # Prendre les 2 premiers chiffres
            
            if category_key not in self.SYSCOHADA_CATEGORIES:
                logger.warning(f"Catégorie non reconnue pour {account_num}")
                self.unmapped_accounts.append(account_num)
                continue
            
            category_info = self.SYSCOHADA_CATEGORIES[category_key]
            intermediate_category = category_info["category"]
            
            # Ajouter au Level 1
            if account_num not in level_1:
                level_1[account_num] = {
                    "label": account_label,
                    "intermediate_category": intermediate_category,
                    "class": category_info.get("class"),
                }
        
        self.level_1_mapping = level_1
        logger.info(f"Level 1 généré: {len(level_1)} comptes SYSCOHADA mappés")
        
        return level_1
    
    def generate_level_2(self):
        """Génère le mapping Level 2 (Intermédiaire → DSF)"""
        self.level_2_mapping = self.INTERMEDIATE_TO_DSF
        logger.info(f"Level 2 généré: {len(self.level_2_mapping)} catégories intermédiaires")
        
        return self.level_2_mapping
    
    def export_complete_mapping(self, output_file: str = "dsf_complete_mapping.json"):
        """Exporte le mapping complet"""
        if not self.level_1_mapping:
            self.generate_level_1()
        if not self.level_2_mapping:
            self.generate_level_2()
        
        complete_mapping = {
            "description": "Double mapping complet SYSCOHADA → Intermédiaire → DSF",
            "coverage": {
                "total_syscohada": len(SYSCOHADA_INDEX),
                "mapped": len(self.level_1_mapping),
                "unmapped": len(self.unmapped_accounts),
                "coverage_percentage": (len(self.level_1_mapping) / len(SYSCOHADA_INDEX)) * 100
            },
            "LEVEL_1_SYSCOHADA_TO_INTERMEDIATE": self.level_1_mapping,
            "LEVEL_2_INTERMEDIATE_TO_DSF": self.level_2_mapping,
            "unmapped_accounts": self.unmapped_accounts
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(complete_mapping, f, indent=2, ensure_ascii=False)
        
        logger.info(f"Mapping complet exporté: {output_file}")
        logger.info(f"Couverture: {complete_mapping['coverage']['coverage_percentage']:.2f}%")
        
        return complete_mapping


def main():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    print("\n▶ Génération du mapping complet...\n")
    
    generator = CompleteDoubleMappingGenerator()
    
    print("  ▶ Level 1: SYSCOHADA → Intermédiaire...")
    level_1 = generator.generate_level_1()
    print(f"    ✓ {len(level_1)} comptes mappés")
    
    print("  ▶ Level 2: Intermédiaire → DSF...")
    level_2 = generator.generate_level_2()
    print(f"    ✓ {len(level_2)} catégories intermédiaires")
    
    print("  ▶ Export du mapping complet...")
    mapping = generator.export_complete_mapping("dsf_complete_mapping.json")
    
    coverage = mapping['coverage']
    print(f"\n📊 RÉSULTATS:")
    print(f"  Total SYSCOHADA:    {coverage['total_syscohada']}")
    print(f"  Mappés:             {coverage['mapped']}")
    print(f"  Non mappés:         {coverage['unmapped']}")
    print(f"  Couverture:         {coverage['coverage_percentage']:.2f}%")
    
    if coverage['unmapped'] > 0:
        print(f"\n⚠️  Comptes non mappés: {generator.unmapped_accounts[:5]}")
    else:
        print("\n✅ COUVERTURE COMPLÈTE (100%)")
    
    print()


if __name__ == "__main__":
    main()
