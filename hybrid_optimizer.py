#!/usr/bin/env python3
"""
HYBRID OPTIMIZATION STRATEGY for DSF Balance Filling
Combines: Fuzzy + Hierarchical + Category-based + Pattern Matching
Target: 20-25% success rate (vs current 11.2% semantic-only)
"""

import re
from typing import Dict, List, Set, Tuple, Optional
from dataclasses import dataclass
from decimal import Decimal
from difflib import SequenceMatcher


@dataclass
class AccountCategory:
    """Account categorization for rules"""
    category: str  # ASSETS, LIABILITIES, EQUITY, INCOME, EXPENSES
    keywords: Set[str]
    regex_patterns: List[str]
    

class HybridOptimizer:
    """Multi-strategy matcher for unmatched cells"""
    
    def __init__(self):
        self._init_categories()
        self._init_rules()
    
    def _init_categories(self):
        """Define account categories and their keywords"""
        self.categories = {
            'IMMOBILISATIONS': AccountCategory(
                category='ASSETS',
                keywords={'immobil', 'outil', 'équip', 'construction', 'terrains', 'stock'},
                regex_patterns=[
                    r'(immobilis[é|ations]|outil|équip|construc|terrain)',
                    r'(stocks?|inventories)',
                ]
            ),
            'CREANCES': AccountCategory(
                category='ASSETS',
                keywords={'créance', 'client', 'débiteur', 'trésor', 'compte'},
                regex_patterns=[
                    r'(créance|client|débiteur|compte)',
                ]
            ),
            'DETTES': AccountCategory(
                category='LIABILITIES',
                keywords={'dette', 'fournisseur', 'créancier', 'emprunt'},
                regex_patterns=[
                    r'(dette|fournisseur|créancier|emprunt)',
                ]
            ),
            'RESULTAT': AccountCategory(
                category='INCOME',
                keywords={'chiffre', 'vente', 'revenue', 'income', 'produit'},
                regex_patterns=[
                    r'(chiffre|vente|revenue|income|produit)',
                ]
            ),
            'CHARGES': AccountCategory(
                category='EXPENSES',
                keywords={'charge', 'dépense', 'personnel', 'amort', 'provisionn'},
                regex_patterns=[
                    r'(charge|dépense|personnel|amort|provisionn)',
                ]
            )
        }
    
    def _init_rules(self):
        """Define mapping rules for common transformations"""
        self.transformation_rules = {
            # French account names → common alternatives
            'STOCKS': ['STOCKS FINIS', 'STOCKS MATERIAUX', 'STOCKS EN COURS'],
            'IMMOBILISATIONS': ['IMMOBI INCORPORELLES', 'IMMOBILISAT CORPORELLES', 'IMMOBILISAT FINANCIERES'],
            'CREANCES CLIENT': ['CREANCES CLIENTS', 'CREANCE CLIENT', 'CLIENT'],
            'DETTES FOURNISSEURS': ['DETTES FOURNISSEUR', 'FOURNISSEURS', 'FOURNISSEUR'],
            'RESULTATS': ['RESULTAT DE L EXERCICE', 'RESULTAT NET', 'BENEFICE', 'PERTE'],
        }
    
    def categorize_label(self, label: str) -> Optional[str]:
        """Assign a category to a cell label"""
        label_lower = label.lower()
        
        for cat_name, cat in self.categories.items():
            # Check keywords
            for keyword in cat.keywords:
                if keyword in label_lower:
                    return cat_name
            
            # Check regex patterns
            for pattern in cat.regex_patterns:
                if re.search(pattern, label_lower, re.IGNORECASE):
                    return cat_name
        
        return None
    
    def get_category_rules(self, label: str) -> List[str]:
        """Get transformation rules for a label"""
        alternatives = []
        label_upper = label.upper()
        
        for base, transforms in self.transformation_rules.items():
            if base.upper() in label_upper or label_upper in base.upper():
                alternatives.extend(transforms)
        
        return alternatives
    
    def hierarchical_match(self, label: str, all_accounts: Dict[str, str]) -> List[Tuple[str, float]]:
        """
        Match by hierarchy: if exact label not found,
        find accounts that START WITH same prefix
        """
        label_words = label.split()
        if not label_words:
            return []
        
        matches = []
        prefix = label_words[0].upper()  # First word as prefix
        
        for compte, account_label in all_accounts.items():
            if account_label.upper().startswith(prefix):
                # Calculate partial similarity
                ratio = SequenceMatcher(None, label.lower(), account_label.lower()).ratio()
                if ratio > 0.35:  # Lower threshold for hierarchical
                    matches.append((compte, ratio))
        
        return sorted(matches, key=lambda x: -x[1])
    
    def pattern_based_match(self, label: str, all_accounts: Dict[str, str]) -> List[Tuple[str, float]]:
        """
        Try regex pattern matching for common cases
        e.g. "STOCKS" → any account with "STOCKS" in name
        """
        matches = []
        label_upper = label.upper()
        
        for compte, account_label in all_accounts.items():
            account_upper = account_label.upper()
            
            # Exact word match (one word in label fully present in account)
            for word in label.split():
                if len(word) > 4 and word.upper() in account_upper:
                    ratio = 0.60  # Moderate confidence for word match
                    matches.append((compte, ratio))
                    break
        
        return sorted(matches, key=lambda x: -x[1])


def demo_hybrid_strategies():
    """Demonstrate optimization strategies"""
    
    optimizer = HybridOptimizer()
    
    print("=" * 80)
    print("HYBRID OPTIMIZATION STRATEGIES DEMO")
    print("=" * 80)
    
    # Test categorization
    test_labels = [
        "STOCKS FINIS",
        "CREANCES CLIENTS",
        "INSTALLATIONS TECHNIQUES",
        "CHIFFRE D'AFFAIRES",
        "CHARGES DE PERSONNEL"
    ]
    
    print("\n1. CATEGORIZATION:")
    print("-" * 40)
    for label in test_labels:
        cat = optimizer.categorize_label(label)
        print(f"  {label:<35} → {cat or 'UNKNOWN'}")
    
    # Test transformation rules
    print("\n2. TRANSFORMATION RULES:")
    print("-" * 40)
    for test_label in ['STOCKS', 'IMMOBILISATIONS', 'CREANCES CLIENT']:
        rules = optimizer.get_category_rules(test_label)
        print(f"  {test_label}:")
        for rule in rules:
            print(f"    → {rule}")
    
    # Test hierarchical matching (mock accounts)
    print("\n3. HIERARCHICAL MATCHING (Example):")
    print("-" * 40)
    mock_accounts = {
        '1010': 'STOCKS MATIERES PREMIERES',
        '1020': 'STOCKS PRODUITS FINIS',
        '1030': 'STOCKS EN COURS',
        '2010': 'CREANCES CLIENTS',
        '2020': 'CREANCES AUTRES',
    }
    
    test_cell = "STOCKS"
    matches = optimizer.hierarchical_match(test_cell, mock_accounts)
    print(f"  Cell label: '{test_cell}'")
    print(f"  Hierarchical matches:")
    for compte, ratio in matches[:3]:
        print(f"    {ratio:.2f} → {compte}: {mock_accounts[compte]}")
    
    # Test pattern matching
    print("\n4. PATTERN-BASED MATCHING (Example):")
    print("-" * 40)
    matches = optimizer.pattern_based_match(test_cell, mock_accounts)
    print(f"  Cell label: '{test_cell}'")
    print(f"  Pattern matches:")
    for compte, ratio in matches[:3]:
        print(f"    {ratio:.2f} → {compte}: {mock_accounts[compte]}")


if __name__ == "__main__":
    demo_hybrid_strategies()
