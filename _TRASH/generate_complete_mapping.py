"""
Generate complete double-mapping JSON with 100% coverage
Output: dsf_final_mapping.json
"""

import json
from pathlib import Path
from syscohada_db import SYSCOHADA_INDEX
from universal_mapping_strategy import UniversalMappingStrategy

def generate_complete_mapping():
    """Generate complete Level 1 & Level 2 mapping"""
    
    # Level 1: SYSCOHADA (1182) → Intermediate Categories
    level_1 = {}
    for account_num in sorted(SYSCOHADA_INDEX.keys()):
        category = UniversalMappingStrategy.get_category(account_num)
        account_obj = SYSCOHADA_INDEX[account_num]
        level_1[account_num] = {
            "intermediate_category": category,
            "account_name": account_obj.intitule
        }
    
    # Level 2: Intermediate Categories → DSF Final Destinations
    level_2 = {}
    for category_key, dsf_info in UniversalMappingStrategy.CATEGORY_TO_DSF.items():
        level_2[category_key] = dsf_info
    
    # Mapping from intermediate categories (with category codes like EQ_CAPITAL)
    # back to readable category names for debugging
    intermediate_categories = {}
    for classe_num, classe_rules in UniversalMappingStrategy.CLASS_CATEGORIES.items():
        for prefix, category_code in classe_rules["mappings"].items():
            if category_code not in intermediate_categories:
                intermediate_categories[category_code] = {
                    "classe": classe_num,
                    "prefixes": []
                }
            intermediate_categories[category_code]["prefixes"].append(prefix)
        
        # Add default
        default_cat = classe_rules["default"]
        if default_cat not in intermediate_categories:
            intermediate_categories[default_cat] = {
                "classe": classe_num,
                "prefixes": ["others"]
            }
    
    # Verify coverage
    total = len(SYSCOHADA_INDEX)
    mapped = sum(1 for acc in level_1.values() if acc["intermediate_category"])
    
    complete_mapping = {
        "version": "1.0",
        "description": "Complete double-mapping: SYSCOHADA (1182) → Intermediate (32) → DSF",
        "generated_date": "2024-01-01",
        "coverage": {
            "total_syscohada": total,
            "explicitly_mapped": mapped,
            "coverage_percentage": 100.0 if mapped == total else (mapped / total * 100)
        },
        "LEVEL_1_SYSCOHADA_TO_INTERMEDIATE": level_1,
        "LEVEL_2_INTERMEDIATE_TO_DSF": level_2,
        "INTERMEDIATE_CATEGORIES_INFO": intermediate_categories,
        "class_distribution": {
            "Classe 1 - Equity": len([a for a in level_1.keys() if a[0] == "1"]),
            "Classe 2 - Fixed Assets": len([a for a in level_1.keys() if a[0] == "2"]),
            "Classe 3 - Current Assets": len([a for a in level_1.keys() if a[0] == "3"]),
            "Classe 4 - Liabilities": len([a for a in level_1.keys() if a[0] == "4"]),
            "Classe 5 - Special": len([a for a in level_1.keys() if a[0] == "5"]),
            "Classe 6 - Expenses": len([a for a in level_1.keys() if a[0] == "6"]),
            "Classe 7 - Revenues": len([a for a in level_1.keys() if a[0] == "7"]),
            "Classe 8 - Control": len([a for a in level_1.keys() if a[0] == "8"]),
            "Classe 9 - Analytical": len([a for a in level_1.keys() if a[0] == "9"]),
        }
    }
    
    return complete_mapping


if __name__ == "__main__":
    print("Generating complete double-mapping JSON with 100% coverage...\n")
    
    mapping = generate_complete_mapping()
    
    # Save to JSON
    output_file = Path("dsf_final_mapping.json")
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(mapping, f, indent=2, ensure_ascii=False)
    
    print(f"✅ Complete mapping saved: {output_file}")
    print(f"\nMapping Statistics:")
    print(f"  Total SYSCOHADA accounts: {mapping['coverage']['total_syscohada']}")
    print(f"  Coverage: {mapping['coverage']['coverage_percentage']:.2f}%")
    print(f"  Unique intermediate categories: {len(mapping['LEVEL_2_INTERMEDIATE_TO_DSF'])}")
    print(f"\nClass Distribution:")
    for classe, count in mapping['class_distribution'].items():
        print(f"  {classe}: {count} accounts")
    
    print(f"\nFile size: {output_file.stat().st_size / 1024:.1f} KB")
