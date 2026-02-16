# -*- coding: utf-8 -*-
from __future__ import annotations

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from dsf_controls import ControlSuite
from dsf_reporting import generate_reports
from dsf_rule_config import load_rule_set
from dsf_rule_engine import Assignment, RuleEngineResult, UnmatchedAccount

TEST_ROOT = Path(__file__).parent
DATA_ROOT = TEST_ROOT / "data"


class ControlSuiteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.rule_set = load_rule_set(DATA_ROOT / "rules_sample.yaml")
        self.control_suite = ControlSuite(self.rule_set)

    def _sample_result(self) -> RuleEngineResult:
        assignments = [
            Assignment(sheet="BILAN PAYSAGE", cell="D14", rule_id="actif_immos", label="IMMOBILISATIONS", amount=Decimal("100000")),
            Assignment(sheet="BILAN PAYSAGE", cell="K33", rule_id="passif_capitaux", label="CAPITAUX PROPRES", amount=Decimal("100000")),
            Assignment(sheet="COMPTE DE RESULTAT", cell="D40", rule_id="resultat_net", label="RESULTAT NET", amount=Decimal("16000")),
        ]
        unmatched = [
            UnmatchedAccount(
                compte="801",
                label="COUTS",
                reason="Aucune règle applicable",
                classe="8",
                balance_side="debit",
                value="1000",
            )
        ]
        stats = {
            "total_assigned": "216000",
            "charges": "12000",
            "produits": "28000",
            "resultat_exercice": "16000",
            "bucket_totals": {"ACTIF": "100000", "PASSIF": "100000"},
            "rule_totals": {"actif_immos": "100000", "passif_capitaux": "100000"},
            "rule_usage": {
                "actif_immos": {"hits": 1, "applied": 1, "amount": "100000"},
                "passif_capitaux": {"hits": 1, "applied": 1, "amount": "100000"},
                "charges": {"hits": 0, "applied": 0, "amount": "0"},
            },
        }
        return RuleEngineResult(assignments=assignments, unmatched=unmatched, stats=stats)

    def test_control_suite_reports_balance_and_result(self) -> None:
        result = self._sample_result()
        report = self.control_suite.evaluate(result)
        names = {issue.name: issue for issue in report.issues}
        self.assertEqual(names["balance_equation"].status, "PASS")
        self.assertEqual(names["result_reconciliation"].status, "PASS")
        self.assertEqual(names["unmatched_accounts"].status, "WARN")

    def test_reporting_outputs(self) -> None:
        result = self._sample_result()
        report = self.control_suite.evaluate(result)
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            outputs = generate_reports(
                result,
                report,
                json_path=tmp_path / "report.json",
                html_path=tmp_path / "report.html",
                title="Test Report",
            )
            self.assertTrue(outputs["json"].exists())
            self.assertTrue(outputs["html"].exists())
            html_text = outputs["html"].read_text(encoding="utf-8")
            self.assertIn("Test Report", html_text)


if __name__ == "__main__":
    unittest.main()
