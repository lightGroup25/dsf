# -*- coding: utf-8 -*-
"""Dynamic DSF allocation engine based on SYSCOHADA rules."""
from __future__ import annotations

import argparse
import json
import logging
from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from openpyxl.utils import column_index_from_string

from balance_normalizer import BalanceNormalizer, NormalizedBalanceRow
from dsf_inventory import DSFInventory, InventoryField
from dsf_rule_config import DSFRule, DSFRuleSet, load_rule_set

logger = logging.getLogger(__name__)


@dataclass
class Assignment:
    sheet: str
    cell: str
    rule_id: str
    label: Optional[str]
    amount: Decimal = Decimal("0")
    accounts: List[Dict[str, str]] = field(default_factory=list)

    def add(self, row: NormalizedBalanceRow, value: Decimal) -> None:
        self.amount += value
        self.accounts.append(
            {
                "compte": row.compte,
                "label": row.label,
                "value": format(value, "f"),
            }
        )

    def to_dict(self) -> Dict[str, object]:
        return {
            "sheet": self.sheet,
            "cell": self.cell,
            "rule_id": self.rule_id,
            "label": self.label,
            "amount": format(self.amount, "f"),
            "accounts": self.accounts,
        }


@dataclass
class UnmatchedAccount:
    compte: str
    label: str
    reason: str
    classe: str
    balance_side: str
    value: str


@dataclass
class RuleEngineResult:
    assignments: List[Assignment]
    unmatched: List[UnmatchedAccount]
    stats: Dict[str, object]

    def to_dict(self) -> Dict[str, object]:
        return {
            "assignments": [assignment.to_dict() for assignment in self.assignments],
            "unmatched": [dataclass_instance.__dict__ for dataclass_instance in self.unmatched],
            "stats": self.stats,
        }


class RuleEngine:
    def __init__(self, rule_set: DSFRuleSet, inventory: DSFInventory):
        self.rule_set = rule_set
        self.inventory = inventory
        self.rules = sorted(rule_set.rules, key=lambda rule: rule.priority)

    def apply(self, rows: Iterable[NormalizedBalanceRow]) -> RuleEngineResult:
        assignments: Dict[Tuple[str, str], Assignment] = {}
        unmatched: List[UnmatchedAccount] = []
        global_stats = defaultdict(Decimal)
        bucket_totals = defaultdict(Decimal)
        rule_totals = defaultdict(Decimal)
        rule_usage: Dict[str, Dict[str, object]] = {
            rule.id: {"hits": 0, "applied": 0, "amount": Decimal("0")}
            for rule in self.rules
        }

        for row in rows:
            success, reason = self._apply_row(
                row,
                assignments,
                global_stats,
                bucket_totals,
                rule_totals,
                rule_usage,
            )
            if not success:
                unmatched.append(
                    UnmatchedAccount(
                        compte=row.compte,
                        label=row.label,
                        reason=reason or "Aucune règle applicable",
                        classe=row.classe,
                        balance_side=row.metadata.get("balance_side", "unknown"),
                        value=format(row.solde_final, "f"),
                    )
                )

        stats_report = self._build_stats(global_stats, assignments, bucket_totals, rule_totals, rule_usage)
        return RuleEngineResult(list(assignments.values()), unmatched, stats_report)

    def _apply_row(
        self,
        row: NormalizedBalanceRow,
        assignments: Dict[Tuple[str, str], Assignment],
        stats: Dict[str, Decimal],
        bucket_totals: Dict[str, Decimal],
        rule_totals: Dict[str, Decimal],
        rule_usage: Dict[str, Dict[str, object]],
    ) -> Tuple[bool, Optional[str]]:
        if "CR_CHARGES" in row.dsf_buckets:
            stats["charges"] += row.debit_balance
        if "CR_PRODUITS" in row.dsf_buckets:
            stats["produits"] += row.credit_balance

        reason = "Aucune règle applicable"
        any_success = False
        
        for rule in self.rules:
            if not self._row_matches_rule(row, rule):
                continue
            usage = rule_usage[rule.id]
            usage["hits"] = int(usage["hits"]) + 1
            # reason = "Aucune cellule compatible dans le template" # Don't overwrite reason immediately
            
            field = self._select_target_field(rule, row)
            if not field:
                # If one rule fails target, we record reason but continue to try others
                # If no rule succeeds eventually, this reason (or last reason) is returned.
                reason = "Aucune cellule compatible (cible introuvable)"
                continue
                
            value = self._extract_value(row, rule)
            if value <= 0:
                # reason = "Solde inexploitable pour la règle"
                continue
                
            key = (field.sheet, field.cell)
            assignment = assignments.setdefault(
                key,
                Assignment(sheet=field.sheet, cell=field.cell, rule_id=rule.id, label=field.label),
            )
            assignment.add(row, value)
            usage["applied"] = int(usage["applied"]) + 1
            usage["amount"] = usage["amount"] + value
            rule_totals[rule.id] += value
            buckets = rule.buckets or row.dsf_buckets
            for bucket in buckets:
                bucket_totals[bucket] += value
            
            any_success = True
            # DO NOT RETURN. Continue to next rule.
            
        if any_success:
            return True, None
        return False, reason

    @staticmethod
    def _row_matches_rule(row: NormalizedBalanceRow, rule: DSFRule) -> bool:
        filters = rule.filters
        compte = row.compte
        if filters.include_classes and row.classe not in filters.include_classes:
            return False
        if filters.include_accounts and compte not in filters.include_accounts:
            return False
        if filters.exclude_accounts and compte in filters.exclude_accounts:
            return False
        if filters.include_prefixes and not any(compte.startswith(prefix) for prefix in filters.include_prefixes):
            return False
        if filters.include_ranges and not any(_compte_in_range(compte, r) for r in filters.include_ranges):
            return False
        if filters.require_sign == "debit" and row.debit_balance <= 0:
            return False
        if filters.require_sign == "credit" and row.credit_balance <= 0:
            return False
        return True

    def _select_target_field(self, rule: DSFRule, row: NormalizedBalanceRow) -> Optional[InventoryField]:
        target = rule.target
        allowed_cells = target.allowed_cells
        if allowed_cells:
            candidate_fields = [self.inventory.get_field(target.sheet, cell) for cell in allowed_cells]
            candidate_fields = [field for field in candidate_fields if field]
        else:
            candidate_fields = list(
                self.inventory.iter_fields(
                    target.sheet,
                    column_letter=target.column,
                    column_type=target.exercice_column_type,
                    row_min=target.row_start,
                    row_max=target.row_end,
                )
            )
        candidate_fields = [field for field in candidate_fields if field and not field.locked and not field.has_formula]
        if not candidate_fields:
            # Fallback for empty inventory sheets (e.g. Note 6, 7)
            # We construct a virtual field at the START of the target window.
            # This is "Blind Writing".
            # logger.warning(f"Force-creation of virtual field for rule {rule.id} on {target.sheet} {target.column}{target.row_start}")
            
            # Use the first row for the "virtual" field target.
            # However, if multiple accounts map to the same rule, they will all stack on this ONE cell.
            # Ideally we want to distribute them? No, usually rules map to a specific line.
            # Or if it's a "list" note, we might want to append?
            # For now, stacking on one cell is better than nothing (and usually corresponds to "Other" or "Total").
            
            # Using row_start.
            v_row = target.row_start
            v_col = target.column
            v_cell = f"{v_col}{v_row}"
            
            return InventoryField(
                sheet=target.sheet,
                cell=v_cell,
                row=v_row,
                column=column_index_from_string(v_col),
                column_letter=v_col,
                column_type=target.exercice_column_type or "unknown",
                label="Virtual Field (Forced)",
                section_hint=None,
                locked=False,
                has_formula=False,
                number_format=None,
                data_validation=[],
                merged=False,
                merge_range=None
            )

        scored = [
            (self._score_field(field, row), field)
            for field in candidate_fields
        ]
        scored.sort(key=lambda item: (-item[0], item[1].row))
        best_score, best_field = scored[0]
        if best_score <= 0:
            return best_field
        return best_field

    @staticmethod
    def _score_field(field: InventoryField, row: NormalizedBalanceRow) -> int:
        if not field.label:
            return 0
        tokens_field = _tokenize(field.label)
        tokens_row = _tokenize(row.label)
        score = 0
        for token in tokens_row:
            if token in tokens_field:
                score += 5
        if row.compte in field.label:
            score += 20
        if row.metadata.get("account_path"):
            for segment in row.metadata["account_path"]:
                normalized = segment.lower()
                if normalized in field.label.lower():
                    score += 3
        return score

    @staticmethod
    def _extract_value(row: NormalizedBalanceRow, rule: DSFRule) -> Decimal:
        sign = rule.filters.require_sign
        if sign == "debit":
            return row.debit_balance
        if sign == "credit":
            return row.credit_balance
        return row.solde_final

    @staticmethod
    def _build_stats(
        stats: Dict[str, Decimal],
        assignments: Dict[Tuple[str, str], Assignment],
        bucket_totals: Dict[str, Decimal],
        rule_totals: Dict[str, Decimal],
        rule_usage: Dict[str, Dict[str, object]],
    ) -> Dict[str, object]:
        total_assigned = sum((assignment.amount for assignment in assignments.values()), Decimal("0"))
        charges = stats.get("charges", Decimal("0"))
        produits = stats.get("produits", Decimal("0"))
        result = produits - charges
        return {
            "total_assigned": format(total_assigned, "f"),
            "charges": format(charges, "f"),
            "produits": format(produits, "f"),
            "resultat_exercice": format(result, "f"),
            "bucket_totals": {bucket: format(value, "f") for bucket, value in bucket_totals.items()},
            "rule_totals": {rule_id: format(value, "f") for rule_id, value in rule_totals.items()},
            "rule_usage": {
                rule_id: {
                    "hits": usage["hits"],
                    "applied": usage["applied"],
                    "amount": format(usage["amount"], "f"),
                }
                for rule_id, usage in rule_usage.items()
            },
        }


def _tokenize(text: str) -> List[str]:
    return [token for token in ''.join(ch if ch.isalnum() else ' ' for ch in text.lower()).split() if token]


def _compte_in_range(compte: str, pattern: str) -> bool:
    pattern = pattern.strip()
    if not pattern:
        return False
    if pattern.endswith("*"):
        return compte.startswith(pattern[:-1])
    if "-" in pattern:
        start, end = pattern.split("-", 1)
        length = len(start)
        prefix = compte[:length]
        try:
            value = int(prefix)
            return int(start) <= value <= int(end)
        except ValueError:
            return False
    try:
        return int(compte[: len(pattern)]) == int(pattern)
    except ValueError:
        return False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Apply DSF rule engine to a balance file")
    parser.add_argument("balance", help="Fichier de balance SYSCOHADA")
    parser.add_argument("--inventory", default="data/dsf_inventory.json", help="Inventaire JSON du template DSF")
    parser.add_argument("--rules", default="config/dsf_rule_prototypes.yaml", help="Fichier de règles YAML/JSON")
    parser.add_argument("--assignments", help="Chemin JSON de sortie pour les affectations")
    parser.add_argument("--report", help="Chemin JSON pour le rapport complet")
    parser.add_argument("--limit", type=int, help="Nombre maximum de lignes de balance à traiter")
    parser.add_argument("--verbose", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO)

    inventory = DSFInventory.from_json(Path(args.inventory))
    rule_set = load_rule_set(Path(args.rules))

    normalizer = BalanceNormalizer(Path(args.balance))
    rows = list(normalizer.iterate())
    if args.limit:
        rows = rows[: args.limit]

    engine = RuleEngine(rule_set, inventory)
    result = engine.apply(rows)

    if args.assignments:
        Path(args.assignments).write_text(
            json.dumps([assignment.to_dict() for assignment in result.assignments], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    if args.report:
        Path(args.report).write_text(
            json.dumps(result.to_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    logger.info("Affectations générées: %s", len(result.assignments))
    logger.info("Comptes non affectés: %s", len(result.unmatched))


if __name__ == "__main__":
    main()
