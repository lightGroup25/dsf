# -*- coding: utf-8 -*-
"""Data structures and helpers for DSF rule configuration."""
from __future__ import annotations

import argparse
import json
import logging
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional

try:  # Optional dependency for YAML configs
    import yaml  # type: ignore
except ImportError:  # pragma: no cover
    yaml = None

logger = logging.getLogger(__name__)


@dataclass
class AccountFilter:
    include_classes: List[str] = field(default_factory=list)
    include_ranges: List[str] = field(default_factory=list)
    include_prefixes: List[str] = field(default_factory=list)
    include_accounts: List[str] = field(default_factory=list)
    exclude_accounts: List[str] = field(default_factory=list)
    require_sign: Optional[str] = None  # "debit" | "credit"
    notes: Optional[str] = None

    @classmethod
    def from_dict(cls, data: Optional[Dict[str, Any]]) -> "AccountFilter":
        return cls(**(data or {}))


@dataclass
class DSFRuleTarget:
    sheet: str
    column: str
    exercice_column_type: str = "exercice_n"
    row_start: Optional[int] = None
    row_end: Optional[int] = None
    allowed_cells: List[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DSFRuleTarget":
        return cls(**(data or {}))


@dataclass
class DSFRule:
    id: str
    label: str
    section: str
    priority: int
    target: DSFRuleTarget
    filters: AccountFilter
    buckets: List[str] = field(default_factory=list)
    notes: Optional[str] = None

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DSFRule":
        payload = dict(data)
        target = DSFRuleTarget.from_dict(payload.pop("target", {}))
        filters = AccountFilter.from_dict(payload.pop("filters", {}))
        return cls(target=target, filters=filters, **payload)

    def to_dict(self) -> Dict[str, Any]:
        payload = asdict(self)
        payload["target"] = asdict(self.target)
        payload["filters"] = asdict(self.filters)
        return payload


@dataclass
class DSFRuleSet:
    version: str
    template_name: str
    generated_from_inventory: Optional[str]
    rules: List[DSFRule]
    metadata: Dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DSFRuleSet":
        rules = [DSFRule.from_dict(rule) for rule in data.get("rules", [])]
        return cls(
            version=data.get("version", "0.0.0"),
            template_name=data.get("template_name", ""),
            generated_from_inventory=data.get("generated_from_inventory"),
            rules=rules,
            metadata=data.get("metadata", {}),
        )

    def to_dict(self) -> Dict[str, Any]:
        payload = {
            "version": self.version,
            "template_name": self.template_name,
            "generated_from_inventory": self.generated_from_inventory,
            "metadata": self.metadata,
            "rules": [rule.to_dict() for rule in self.rules],
        }
        return payload

    def describe(self) -> Dict[str, Any]:
        summary = {
            "version": self.version,
            "template_name": self.template_name,
            "rules": len(self.rules),
            "sections": {},
        }
        for rule in self.rules:
            summary["sections"].setdefault(rule.section, 0)
            summary["sections"][rule.section] += 1
        return summary


def _load_structured_data(path: Path) -> Dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    suffix = path.suffix.lower()
    if suffix in {".yaml", ".yml"}:
        if yaml is None:  # pragma: no cover
            raise RuntimeError("PyYAML est requis pour charger des fichiers YAML")
        return yaml.safe_load(text)
    return json.loads(text)


def load_rule_set(path: Path) -> DSFRuleSet:
    data = _load_structured_data(path)
    return DSFRuleSet.from_dict(data)


def save_rule_set(rule_set: DSFRuleSet, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = rule_set.to_dict()
    suffix = path.suffix.lower()
    if suffix in {".yaml", ".yml"}:
        if yaml is None:  # pragma: no cover
            raise RuntimeError("PyYAML est requis pour écrire des fichiers YAML")
        path.write_text(yaml.safe_dump(payload, allow_unicode=True, sort_keys=False), encoding="utf-8")
    else:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("Configuration DSF sauvegardée dans %s", path)
    return path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Inspecter un fichier de règles DSF")
    parser.add_argument(
        "config",
        nargs="?",
        default="config/dsf_rule_prototypes.yaml",
        help="Chemin du fichier YAML/JSON",
    )
    parser.add_argument("--json", action="store_true", help="Affiche le contenu complet en JSON")
    parser.add_argument("--summary", action="store_true", help="Affiche un résumé synthétique")
    parser.add_argument("--verbose", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO)
    rule_set = load_rule_set(Path(args.config))

    if args.summary:
        print(json.dumps(rule_set.describe(), ensure_ascii=False, indent=2))
        return

    if args.json:
        print(json.dumps(rule_set.to_dict(), ensure_ascii=False, indent=2))
        return

    print(f"Version: {rule_set.version}")
    print(f"Template: {rule_set.template_name}")
    print(f"Règles chargées: {len(rule_set.rules)}")
    for rule in rule_set.rules:
        print(f"- {rule.id} ({rule.section}) -> {rule.target.sheet}:{rule.target.column}")


if __name__ == "__main__":
    main()
