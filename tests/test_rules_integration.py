import unittest
from pathlib import Path
import sys
import os

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from dsf_rule_config import load_rule_set
from dsf_rule_engine import RuleEngine
from dsf_inventory import DSFInventory

import unittest
import json
from pathlib import Path
import sys
import os

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from dsf_rule_config import load_rule_set
from dsf_inventory import DSFInventory

class TestRulesIntegration(unittest.TestCase):
    def setUp(self):
        self.config_path = Path("config/dsf_rule_generated.yaml")
        if not self.config_path.exists():
            self.skipTest(f"Config file not found: {self.config_path}")
        
        self.rule_set = load_rule_set(self.config_path)
        
        self.inventory_path = Path("data/dsf_inventory.json")
        if not self.inventory_path.exists():
            self.skipTest(f"Inventory file not found: {self.inventory_path}")
            
        with open(self.inventory_path, 'r', encoding='utf-8') as f:
            payload = json.load(f)
        self.inventory = DSFInventory(payload)

    def test_rule_count(self):
        """Verify that we have at least the expected number of rules."""
        self.assertGreaterEqual(len(self.rule_set.rules), 63, 
                                f"Expected at least 63 rules, found {len(self.rule_set.rules)}")

    def test_critical_categories_exist(self):
        """Ensure specific categories added recently are present (checking converted IDs)."""
        rule_ids = {r.id for r in self.rule_set.rules}
        
        # We expect representative rules for each major section
        expected_ids = [
            "auto_fa_land",       # Fixed Assets
            "auto_eq_capital",    # Equity
            "auto_sp_provisions", # Special (Class 5)
            "auto_ct_other",      # Control (Class 8)
            "auto_an_revenue"     # Analytical (Class 9)
        ]
        
        for rid in expected_ids:
            self.assertIn(rid, rule_ids, f"Rule ID {rid} missing from generated rules")

    def test_mapping_logic_static(self):
        """Test that specific prefixes are mapped to the correct rules."""
        # Account 101 -> prefix "10" -> auto_eq_capital -> BILAN PAYSAGE
        
        target_rule = next((r for r in self.rule_set.rules if r.id == "auto_eq_capital"), None)
        self.assertIsNotNone(target_rule)
        self.assertIn("10", target_rule.filters.include_prefixes)
        self.assertEqual(target_rule.target.sheet, "BILAN PAYSAGE")

    def test_special_categories_provisions_static(self):
        """Test mapping for failure/catch-all categories like SP_PROVISIONS."""
        # auto_sp_provisions -> prefix "50"
        
        target_rule = next((r for r in self.rule_set.rules if r.id == "auto_sp_provisions"), None)
        self.assertIsNotNone(target_rule)
        self.assertIn("50", target_rule.filters.include_prefixes)

    def test_analytical_flux_static(self):
        """Test analytical categories for Cash Flow exist."""
        # auto_an_revenue should target TABLEAU DES FLUX DE TRESORERIE
        target_rule = next((r for r in self.rule_set.rules if r.id == "auto_an_revenue"), None)
        self.assertIsNotNone(target_rule)
        self.assertEqual(target_rule.target.sheet, "TABLEAU DES FLUX DE TRESORERIE")

if __name__ == '__main__':
    unittest.main()
