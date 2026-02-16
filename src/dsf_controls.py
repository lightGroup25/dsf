# -*- coding: utf-8 -*-
"""Automated DSF control suite (balance equality, result reconciliation, rubrics)."""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, List, Optional

from dsf_rule_config import DSFRuleSet
from dsf_rule_engine import Assignment, RuleEngineResult


@dataclass
class ControlIssue:
    name: str
    status: str  # PASS, WARN, FAIL
    message: str
    details: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ControlReport:
    issues: List[ControlIssue]
    unmatched_accounts: List[Dict[str, Any]]

    @property
    def summary(self) -> Dict[str, int]:
        counts = {"PASS": 0, "WARN": 0, "FAIL": 0}
        for issue in self.issues:
            if issue.status in counts:
                counts[issue.status] += 1
        return counts


class ControlSuite:
    def __init__(self, rule_set: DSFRuleSet, tolerance: Decimal = Decimal("1")):
        self.rule_set = rule_set
        self.tolerance = tolerance

    def evaluate(self, result: RuleEngineResult) -> ControlReport:
        issues: List[ControlIssue] = []
        issues.append(self._check_balance(result.stats))
        issues.append(self._check_result(result.stats, result.assignments))
        issues.extend(self._check_empty_rubrics(result.stats))
        issues.append(self._summarize_unmatched(result.unmatched))
        return ControlReport(issues=issues, unmatched_accounts=[account.__dict__ for account in result.unmatched])

    def _check_balance(self, stats: Dict[str, Any]) -> ControlIssue:
        bucket_totals = stats.get("bucket_totals", {}) or {}
        actif = _to_decimal(bucket_totals.get("ACTIF"))
        passif = _to_decimal(bucket_totals.get("PASSIF"))
        diff = abs(actif - passif)
        status = "PASS" if diff <= self.tolerance else "FAIL"
        message = "Total Actif = Total Passif" if status == "PASS" else "Écart entre Total Actif et Passif"
        return ControlIssue(
            name="balance_equation",
            status=status,
            message=message,
            details={"actif": str(actif), "passif": str(passif), "difference": str(diff)},
        )

    def _check_result(self, stats: Dict[str, Any], assignments: List[Assignment]) -> ControlIssue:
        result_stat = _to_decimal(stats.get("resultat_exercice"))
        net_assignments = [
            assignment
            for assignment in assignments
            if assignment.label and "RESULT" in assignment.label.upper()
        ]
        if not net_assignments:
            return ControlIssue(
                name="result_presence",
                status="WARN",
                message="Aucun poste 'Résultat' n'a été rempli dans le DSF",
                details={},
            )
        assigned_value = sum((assignment.amount for assignment in net_assignments), Decimal("0"))
        diff = abs(result_stat - assigned_value)
        status = "PASS" if diff <= self.tolerance else "FAIL"
        return ControlIssue(
            name="result_reconciliation",
            status=status,
            message="Résultat CR cohérent avec Bilan" if status == "PASS" else "Résultat CR ≠ DSF",
            details={
                "result_compte_resultat": str(result_stat),
                "result_poste_bilan": str(assigned_value),
                "difference": str(diff),
                "cells": [f"{a.sheet}!{a.cell}" for a in net_assignments],
            },
        )

    def _check_empty_rubrics(self, stats: Dict[str, Any]) -> List[ControlIssue]:
        issues: List[ControlIssue] = []
        usage = stats.get("rule_usage", {}) or {}
        empty_rules: List[str] = []
        for rule in self.rule_set.rules:
            rule_stats = usage.get(rule.id)
            if not rule_stats:
                continue
            hits = int(rule_stats.get("hits", 0))
            amount = _to_decimal(rule_stats.get("amount"))
            if hits > 0 and amount == Decimal("0"):
                empty_rules.append(rule.id)
        if empty_rules:
            issues.append(
                ControlIssue(
                    name="empty_rubrics",
                    status="WARN",
                    message=f"{len(empty_rules)} rubriques attendues n'ont reçu aucune écriture",
                    details={"rules": empty_rules},
                )
            )
        return issues

    def _summarize_unmatched(self, unmatched: List[Any]) -> ControlIssue:
        count = len(unmatched)
        status = "PASS" if count == 0 else "WARN"
        details = {"count": count}
        if count:
            details["sample"] = [account.__dict__ for account in unmatched[:5]]
        return ControlIssue(
            name="unmatched_accounts",
            status=status,
            message="Aucun compte ignoré" if status == "PASS" else f"{count} comptes ignorés",
            details=details,
        )


def _to_decimal(value: Optional[Any]) -> Decimal:
    if value is None:
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):  # pragma: no cover
        return Decimal("0")


__all__ = ["ControlSuite", "ControlIssue", "ControlReport"]
