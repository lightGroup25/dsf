# -*- coding: utf-8 -*-
"""Balance normalizer ready for DSF rule engine."""
from __future__ import annotations

import argparse
import json
import logging
from collections import defaultdict
from dataclasses import dataclass, asdict, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, Iterator, List, Optional

from balance_transformer import BalanceRow, BalanceStreamParser
from syscohada_db import SYSCOHADA_INDEX, SyscohadaAccount

logger = logging.getLogger(__name__)

NATURAL_SIDE_BY_CLASS = {
    "1": "credit",
    "2": "debit",
    "3": "debit",
    "4": "credit",
    "5": "debit",
    "6": "debit",
    "7": "credit",
    "8": "debit",
    "9": "credit",
}


@dataclass
class NormalizedBalanceRow:
    compte: str
    label: str
    classe: str
    natural_side: str
    solde_final: Decimal
    debit_balance: Decimal
    credit_balance: Decimal
    dsf_buckets: List[str]
    metadata: Dict[str, object] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, object]:
        payload = asdict(self)
        payload["solde_final"] = format(self.solde_final, "f")
        payload["debit_balance"] = format(self.debit_balance, "f")
        payload["credit_balance"] = format(self.credit_balance, "f")
        return payload


class BalanceNormalizer:
    def __init__(
        self,
        balance_file: Path,
        chunk_size: int = 500,
        column_overrides: Optional[Dict[str, object]] = None,
    ):
        self.balance_file = Path(balance_file)
        self.parser = BalanceStreamParser(
            self.balance_file,
            chunk_size=chunk_size,
            column_overrides=column_overrides,
        )
        self.col_map: Optional[Dict[str, int]] = None

    def __enter__(self) -> "BalanceNormalizer":  # pragma: no cover
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:  # pragma: no cover
        self.close()

    def prepare(self) -> Dict[str, int]:
        if not self.col_map:
            self.col_map = self.parser.load_balance_structure()
        return self.col_map

    def iterate(self) -> Iterator[NormalizedBalanceRow]:
        col_map = self.prepare()
        for raw_row in self.parser.iter_balance_rows(col_map):
            yield self.normalize_row(raw_row)
        self.close()

    def close(self) -> None:
        self.parser.close()

    def normalize_row(self, row: BalanceRow) -> NormalizedBalanceRow:
        account = self._lookup_account(row.compte_num)
        classe = str(account.classe) if account else row.compte_num[:1]
        natural_side = NATURAL_SIDE_BY_CLASS.get(classe, "debit")
        
        # If solde_final is 0 but we have debit/credit, calculate it
        solde_final = _to_decimal(row.solde_final)
        if solde_final == 0 and (row.debit or row.credit):
            # solde_final = debit - credit (debit is positive, credit is negative)
            solde_final = _to_decimal(row.debit) - _to_decimal(row.credit)
        
        debit_balance = solde_final if solde_final > 0 else Decimal("0")
        credit_balance = -solde_final if solde_final < 0 else Decimal("0")
        dsf_buckets = determine_buckets(classe, debit_balance, credit_balance)
        metadata = {
            "account_path": account.path if account else [],
            "account_label_source": "syscohada" if account else "balance",
            "raw": row.to_dict(),
            "balance_side": "debit" if debit_balance > 0 else "credit" if credit_balance > 0 else "zero",
        }
        label = account.intitule if account else row.compte_label
        return NormalizedBalanceRow(
            compte=row.compte_num,
            label=label,
            classe=classe,
            natural_side=natural_side,
            solde_final=solde_final,
            debit_balance=debit_balance,
            credit_balance=credit_balance,
            dsf_buckets=dsf_buckets,
            metadata=metadata,
        )

    @staticmethod
    def _lookup_account(compte: str) -> Optional[SyscohadaAccount]:
        if compte in SYSCOHADA_INDEX:
            return SYSCOHADA_INDEX[compte]
        for length in range(len(compte) - 1, 1, -1):
            prefix = compte[:length]
            if prefix in SYSCOHADA_INDEX:
                return SYSCOHADA_INDEX[prefix]
        return None


def determine_buckets(classe: str, debit_balance: Decimal, credit_balance: Decimal) -> List[str]:
    buckets: List[str] = []
    if classe in {"2", "3"} and debit_balance > 0:
        buckets.append("BILAN_ACTIF")
    if classe == "4":
        if debit_balance > 0:
            buckets.append("BILAN_ACTIF_CREANCES")
        if credit_balance > 0:
            buckets.append("BILAN_PASSIF_DETTES")
    if classe == "5":
        if debit_balance > 0:
            buckets.append("BILAN_ACTIF_TRESORERIE")
        if credit_balance > 0:
            buckets.append("BILAN_PASSIF_TRESORERIE")
    if classe == "1" and (credit_balance > 0 or debit_balance > 0):
        buckets.append("BILAN_PASSIF_CAPITAUX")
    if classe == "6" and debit_balance > 0:
        buckets.append("CR_CHARGES")
    if classe == "7" and credit_balance > 0:
        buckets.append("CR_PRODUITS")
    if classe == "8":
        buckets.append("ANALYTIQUE")
    if classe == "9":
        buckets.append("ENGAGEMENTS")
    return buckets or ["HORS_PÉRIMÈTRE"]


def _to_decimal(value: object) -> Decimal:
    if value is None:
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    try:
        text = str(value).strip().replace(" ", "")
        text = text.replace(",", ".") if text.count(",") == 1 and "." not in text else text
        return Decimal(text or "0")
    except (InvalidOperation, AttributeError):  # pragma: no cover
        logger.debug("Impossible de convertir %s en Decimal", value)
        return Decimal("0")


def summarize_entries(entries: Iterator[NormalizedBalanceRow]) -> Dict[str, object]:
    total = 0
    buckets = defaultdict(int)
    sums = defaultdict(Decimal)
    snapshot: List[Dict[str, object]] = []
    for entry in entries:
        total += 1
        for bucket in entry.dsf_buckets:
            buckets[bucket] += 1
            sums[bucket] += entry.debit_balance + entry.credit_balance
        if len(snapshot) < 50:
            snapshot.append(entry.to_dict())
    return {
        "total_rows": total,
        "bucket_counts": dict(buckets),
        "bucket_amounts": {k: format(v, "f") for k, v in sums.items()},
        "sample": snapshot,
    }


def parse_cli() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Normalize a SYSCOHADA balance file")
    parser.add_argument("balance", help="Chemin du fichier de balance")
    parser.add_argument("--chunk-size", type=int, default=500)
    parser.add_argument("--output", help="Fichier JSON pour sauvegarder les lignes normalisées")
    parser.add_argument("--summary", action="store_true", help="Affiche uniquement le résumé agrégé")
    parser.add_argument("--limit", type=int, help="Nombre maximal de lignes à exporter")
    parser.add_argument("--verbose", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_cli()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO)
    normalizer = BalanceNormalizer(Path(args.balance), chunk_size=args.chunk_size)
    entries: Optional[List[NormalizedBalanceRow]] = None
    if args.output or args.summary:
        entries = list(normalizer.iterate())

    if args.summary and entries is not None:
        summary = summarize_entries(iter(entries))
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return

    if args.output and entries is not None:
        limited = entries[: args.limit] if args.limit else entries
        payload = [entry.to_dict() for entry in limited]
        Path(args.output).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("%s lignes sauvegardées dans %s", len(limited), args.output)
        return

    stream = iter(entries) if entries is not None else normalizer.iterate()
    emitted = 0
    for entry in stream:
        print(json.dumps(entry.to_dict(), ensure_ascii=False))
        emitted += 1
        if args.limit and emitted >= args.limit:
            break


if __name__ == "__main__":
    main()
