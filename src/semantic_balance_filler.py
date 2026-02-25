# -*- coding: utf-8 -*-
"""
Semantic Balance Filler - Version Ultra-Optimisée
Traitement accéléré pour sheets 20 et 34 avec caches intelligents
"""
from __future__ import annotations

import json
import logging
import re
import unicodedata
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field, replace
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
from threading import Lock
from typing import Dict, List, Optional, Set, Tuple, Any
import time

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import coordinate_to_tuple, get_column_letter
from openpyxl.cell import Cell

from balance_normalizer import BalanceNormalizer, NormalizedBalanceRow
from dsf_inventory import DSFInventory
from column_formula_detector import ColumnFormulaDetector

logger = logging.getLogger(__name__)


@dataclass
class CellNeed:
    """Represents a need from a single cell"""
    sheet: str
    cell: str
    row_num: int
    col_num: int
    row_label: Optional[str] = None
    col_label: Optional[str] = None
    is_merged: bool = False
    merge_range: Optional[str] = None


@dataclass
class MatchedAccount:
    """Single account match for a cell"""
    compte: str
    label: str
    amount: Decimal
    side: str  # "debit" or "credit"
    similarity_score: float


@dataclass
class CellAssignment:
    """Final assignment for a cell"""
    sheet: str
    cell: str
    row_label: str
    col_label: str
    is_total: bool
    matched_accounts: List[MatchedAccount] = field(default_factory=list)
    total_amount: Decimal = Decimal("0")
    source_accounts: List[str] = field(default_factory=list)
    confidence: float = 0.0
    notes: str = ""


@dataclass(frozen=True)
class BusinessRule:
    sheet_groups: Tuple[str, ...]
    label_keywords: Tuple[str, ...]
    prefixes: Tuple[str, ...]
    classes: Tuple[str, ...] = ()
    aggregate: bool = False
    min_score: float = 0.60
    max_accounts: int = 12
    allow_n1_fallback: bool = False
    value_source_override: Optional[str] = None
    allow_stock_fallback_when_zero: bool = False
    allow_zero_fill: bool = False


@dataclass
class SheetMetadata:
    """Cache optimisé des métadonnées d'une feuille"""
    merged_masters: Set[str]
    merged_ranges: List[str]
    col_labels: Dict[int, str]
    max_row: int
    max_col: int
    row_labels: Dict[int, str] = field(default_factory=dict)
    row_norms: Dict[int, str] = field(default_factory=dict)
    row_tokens: Dict[int, Set[str]] = field(default_factory=dict)
    has_heavy_merges: bool = False
    row_patterns: Dict[int, Dict[str, Any]] = field(default_factory=dict)  # Patterns pré-calculés


@dataclass
class PrecomputedBalance:
    """Montants pré-calculés pour tous les comptes"""
    by_compte: Dict[str, NormalizedBalanceRow]
    amounts: Dict[Tuple[str, Optional[str]], Decimal]
    amounts_by_prefix: Dict[str, Dict[Optional[str], Decimal]]
    slices: Dict[str, Dict[str, Decimal]]


@dataclass
class HeavySheetCache:
    """Cache spécifique pour sheets 20 et 34"""
    row_patterns: Dict[int, Dict[str, Any]] = field(default_factory=dict)
    formula_results: Dict[str, Decimal] = field(default_factory=dict)
    computed_values: Dict[Tuple[int, int], Decimal] = field(default_factory=dict)
    total_rows: Set[int] = field(default_factory=set)
    subtotal_rows: Set[int] = field(default_factory=set)


# P3-M1 : Paires de termes sémantiquement opposés
FORBIDDEN_CROSS_TERMS: List[Tuple[Set[str], Set[str]]] = [
    ({"charge", "charges"}, {"produit", "produits"}),
    ({"actif"}, {"passif"}),
    ({"debiteur", "debit"}, {"crediteur", "credit"}),
    ({"fournisseur", "fournisseurs"}, {"client", "clients"}),
    ({"emploi", "emplois"}, {"ressource", "ressources"}),
    ({"amortissement", "amortissements"}, {"reintegration", "reintegrations", "reprise", "reprises"}),
]


class TimeoutError(Exception):
    pass


class SemanticBalanceFiller:
    def __init__(
        self,
        template_path: Path | str,
        balance_file: Path | str,
        inventory: DSFInventory,
        fuzzy_threshold: float = 0.70,
        enable_hybrid: bool = False,
        column_overrides: Optional[Dict[str, object]] = None,
        previous_balance_file: Optional[Path | str] = None,
        previous_column_overrides: Optional[Dict[str, object]] = None,
        allow_n1_fallback_without_prev: bool = False,
        preloaded_workbook: Optional[object] = None,
        skip_sheets: Optional[Set[str]] = None,
    ):
        self.template_path = Path(template_path)
        self.balance_file = Path(balance_file)
        self.previous_balance_file = Path(previous_balance_file) if previous_balance_file else None
        self.inventory = inventory
        self.fuzzy_threshold = fuzzy_threshold
        self.min_match_score = max(0.62, fuzzy_threshold)
        self.enable_hybrid = enable_hybrid
        self.column_overrides = column_overrides or {
            "compte": 1,
            "label": 4,
            "debit_columns": [24, 21, 17, 11],
            "credit_columns": [27, 26, 20, 14, 13],
        }
        self.previous_column_overrides = previous_column_overrides or self.column_overrides
        self.allow_n1_fallback_without_prev = allow_n1_fallback_without_prev
        
        # Workbooks et données
        self.wb = preloaded_workbook
        self.balance_accounts: Dict[str, NormalizedBalanceRow] = {}
        self.previous_balance_accounts: Dict[str, NormalizedBalanceRow] = {}
        self.assignments: List[CellAssignment] = []
        self.unmatched_cells: List[CellNeed] = []
        self.formula_detectors: Dict[str, ColumnFormulaDetector] = {}
        self.formulas_applied: int = 0
        self.formula_logs: List[Dict[str, str]] = []
        self.account_usage: Counter[str] = Counter()
        self.business_rules: List[BusinessRule] = self._build_business_rules()
        self.account_time_slices: Dict[str, Dict[str, Decimal]] = {}
        self.previous_account_time_slices: Dict[str, Dict[str, Decimal]] = {}
        self._skip_sheets: Set[str] = set(skip_sheets or [])
        
        # Caches principaux
        self._sheet_metadata: Dict[str, SheetMetadata] = {}
        self._balance_precomputed: Dict[Optional[str], PrecomputedBalance] = {}
        self._row_match_cache: Dict[Tuple[str, str, str, str], List[MatchedAccount]] = {}
        self._fuzzy_similarity_cache: Dict[Tuple[str, str], float] = {}
        self._need_cache: Dict[Tuple[str, str, str, str, str, str], Optional[CellAssignment]] = {}
        self._col_label_cache: Dict[Tuple[str, int], Optional[str]] = {}
        self._sheet_source_hints_cache: Dict[str, Dict[int, Tuple[str, int]]] = {}
        self._sheet_code_row_cache: Dict[str, Dict[str, int]] = {}
        
        # Caches spécialisés pour sheets lourdes
        self._heavy_sheet_cache: Dict[str, HeavySheetCache] = {}
        self._merged_bitmap: Dict[str, bool] = {}
        self._heavy_results_cache: Dict[Tuple[str, str, str], Optional[CellAssignment]] = {}
        self._prefix_amounts_cache: Dict[Tuple[str, ...], Dict[Optional[str], Decimal]] = {}
        
        # Compteurs de performance
        self._previous_assignments_count: Dict[str, int] = {}
        
        # Locks pour thread safety
        self._cache_lock = Lock()
        self._assignments_lock = Lock()
        self._unmatched_lock = Lock()
        self._merged_cache_lock = Lock()
        
        # Constantes de performance
        self.BATCH_SIZE = 25  # Réduit pour sheets lourdes
        self.HEAVY_BATCH_SIZE = 10
        self.PARALLEL_THRESHOLD = 5
        self.MAX_MERGED_CELLS = 50000
        self.CACHE_TIMEOUT = 30
        self.HEAVY_SHEETS = {"20", "34", "NOTE 3A", "NOTE 3C", "NOTE 34", "NOTE 4"}

    def load(self) -> None:
        """Load template and normalize balance with optimisations"""
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {self.template_path}")
        if not self.balance_file.exists():
            raise FileNotFoundError(f"Balance not found: {self.balance_file}")
        if self.previous_balance_file and not self.previous_balance_file.exists():
            raise FileNotFoundError(f"Previous balance not found: {self.previous_balance_file}")

        # Load template workbook unless we already have one
        if self.wb is None:
            self.wb = load_workbook(self.template_path, data_only=False)
        else:
            logger.info("Using preloaded workbook for semantic fill")
        
        # Cache toutes les métadonnées
        logger.info("Caching sheet metadata...")
        for sheet_name in self.wb.sheetnames:
            success = self._cache_sheet_metadata_with_timeout(sheet_name)
            if not success:
                logger.warning(f"  ⚠️ {sheet_name}: timeout - mode light activé")

        # Detect formulas
        logger.info("Detecting column formulas in template...")
        for sheet_name in self.wb.sheetnames:
            detector = ColumnFormulaDetector(self.wb[sheet_name])
            formulas = detector.detect()
            if formulas:
                self.formula_detectors[sheet_name] = detector
                logger.info(f"  {sheet_name}: {len(formulas)} formula columns detected")
        
        # Load balance accounts
        self.balance_accounts = self._load_balance_accounts(self.balance_file, self.column_overrides)
        logger.info(f"Loaded {len(self.balance_accounts)} accounts from balance")
        
        # Extract time slices
        self.account_time_slices = self._extract_balance_time_slices_from_file(
            self.balance_file,
            self.column_overrides,
        )
        logger.info("Loaded timeslices for %s accounts", len(self.account_time_slices))

        # Pré-calcule tous les montants pour N
        self._balance_precomputed['n'] = self._precompute_balance_amounts(
            self.balance_accounts, 
            self.account_time_slices,
            'n'
        )
        logger.info("Precomputed N balance amounts")

        # Handle previous balance
        if self.previous_balance_file:
            from_same_file = (
                self.previous_balance_file.resolve() == self.balance_file.resolve()
                and self.previous_column_overrides is not None
            )
            self.previous_balance_accounts = self._load_balance_accounts(
                self.previous_balance_file,
                self.previous_column_overrides,
            )
            self.previous_account_time_slices = self._extract_balance_time_slices_from_file(
                self.previous_balance_file,
                self.previous_column_overrides,
            )
            
            if self.previous_balance_accounts:
                self._balance_precomputed['n1'] = self._precompute_balance_amounts(
                    self.previous_balance_accounts,
                    self.previous_account_time_slices,
                    'n1'
                )
                logger.info(f"Precomputed N-1 balance amounts from {len(self.previous_balance_accounts)} accounts")
            
            if from_same_file:
                logger.info("N-1 chargé depuis les colonnes d'ouverture de la balance N")
            else:
                logger.info("Loaded N-1 balance: %s accounts", len(self.previous_balance_accounts))

        non_zero = sum(1 for row in self.balance_accounts.values() if row.debit_balance > 0 or row.credit_balance > 0)
        logger.info(f"  Accounts with non-zero amounts: {non_zero}/{len(self.balance_accounts)}")

    def _cache_sheet_metadata_with_timeout(self, sheet_name: str) -> bool:
        """Cache metadata with timeout protection"""
        import threading
        
        timeout_occurred = False
        
        def timeout_handler():
            nonlocal timeout_occurred
            timeout_occurred = True
        
        timer = threading.Timer(self.CACHE_TIMEOUT, timeout_handler)
        timer.daemon = True
        timer.start()
        
        try:
            self._cache_sheet_metadata_optimized(sheet_name)
            timer.cancel()
            return not timeout_occurred
        except Exception as e:
            timer.cancel()
            if timeout_occurred:
                logger.warning(f"⚠️ Cache timeout pour {sheet_name} - mode light")
                ws = self.wb[sheet_name]
                self._sheet_metadata[sheet_name] = SheetMetadata(
                    merged_masters=set(),
                    merged_ranges=[],
                    col_labels={},
                    row_labels={},
                    row_norms={},
                    row_tokens={},
                    max_row=ws.max_row,
                    max_col=ws.max_column,
                    has_heavy_merges=True
                )
                return False
            raise e

    def _cache_sheet_metadata_optimized(self, sheet_name: str) -> None:
        """Cache optimisé pour toutes les feuilles"""
        ws = self.wb[sheet_name]
        is_heavy = sheet_name in self.HEAVY_SHEETS or "NOTE" in sheet_name.upper()
        
        # Compter les merged cells
        total_merged_cells = 0
        merged_ranges_list = []
        for merged_range in ws.merged_cells.ranges:
            cells_count = len(list(merged_range.cells))
            total_merged_cells += cells_count
            merged_ranges_list.append(merged_range.coord)
        
        # Mode light si trop de fusions
        if total_merged_cells > self.MAX_MERGED_CELLS:
            logger.warning(f"  ⚠️ {sheet_name}: {total_merged_cells} merged cells - mode light")
            merged_masters = set()
            for merged_range in ws.merged_cells.ranges:
                merged_masters.add(merged_range.start_cell.coordinate)
            
            # Cache limité
            col_labels = {}
            max_cols = min(ws.max_column, 30)
            for col_num in range(1, max_cols + 1):
                col_labels[col_num] = self._extract_col_label(ws, col_num)
            
            row_labels = {}
            row_norms = {}
            row_tokens = {}
            row_patterns = {}
            max_rows = min(ws.max_row, 300)
            
            for row_num in range(11, max_rows + 1):
                label = self._extract_row_label(ws, row_num)
                if label:
                    row_labels[row_num] = label
                    norm = self._normalize_text(label)
                    row_norms[row_num] = norm
                    row_tokens[row_num] = self._tokenize(label)
                    
                    # Pré-calcul des patterns
                    if is_heavy:
                        row_patterns[row_num] = {
                            'is_total': 'total' in norm or 'somme' in norm,
                            'is_subtotal': 'sous total' in norm,
                            'prefixes': self._preferred_prefixes_for_label(label)
                        }
            
            self._sheet_metadata[sheet_name] = SheetMetadata(
                merged_masters=merged_masters,
                merged_ranges=merged_ranges_list,
                col_labels=col_labels,
                max_row=ws.max_row,
                max_col=ws.max_column,
                row_labels=row_labels,
                row_norms=row_norms,
                row_tokens=row_tokens,
                row_patterns=row_patterns,
                has_heavy_merges=True
            )
            
            # Initialiser cache heavy sheet
            if is_heavy:
                self._init_heavy_sheet_cache(sheet_name, row_patterns)
            
            return
        
        # Mode normal
        merged_masters = set()
        for merged_range in ws.merged_cells.ranges:
            merged_masters.add(merged_range.start_cell.coordinate)
        
        # Cache column labels
        col_labels = {}
        for col_num in range(1, ws.max_column + 1):
            col_labels[col_num] = self._extract_col_label(ws, col_num)
        
        # Cache row labels avec patterns
        row_labels = {}
        row_norms = {}
        row_tokens = {}
        row_patterns = {}
        
        for row_num in range(11, ws.max_row + 1):
            label = self._extract_row_label(ws, row_num)
            if label:
                row_labels[row_num] = label
                norm = self._normalize_text(label)
                row_norms[row_num] = norm
                row_tokens[row_num] = self._tokenize(label)
                
                # Pré-calcul des patterns pour sheets lourdes
                if is_heavy:
                    row_patterns[row_num] = {
                        'is_total': 'total' in norm or 'somme' in norm,
                        'is_subtotal': 'sous total' in norm,
                        'prefixes': self._preferred_prefixes_for_label(label)
                    }
        
        self._sheet_metadata[sheet_name] = SheetMetadata(
            merged_masters=merged_masters,
            merged_ranges=merged_ranges_list,
            col_labels=col_labels,
            max_row=ws.max_row,
            max_col=ws.max_column,
            row_labels=row_labels,
            row_norms=row_norms,
            row_tokens=row_tokens,
            row_patterns=row_patterns,
            has_heavy_merges=False
        )
        
        # Initialiser cache heavy sheet
        if is_heavy:
            self._init_heavy_sheet_cache(sheet_name, row_patterns)
        
        logger.info(f"  ✓ {sheet_name}: {len(merged_masters)} masters, {len(col_labels)} cols, {len(row_labels)} rows")

    def _init_heavy_sheet_cache(self, sheet_name: str, row_patterns: Dict[int, Dict]) -> None:
        """Initialise le cache pour sheets lourdes"""
        heavy_cache = HeavySheetCache(row_patterns=row_patterns)
        
        # Pré-remplir les ensembles de totaux
        for row_num, pattern in row_patterns.items():
            if pattern.get('is_total'):
                heavy_cache.total_rows.add(row_num)
            if pattern.get('is_subtotal'):
                heavy_cache.subtotal_rows.add(row_num)
        
        self._heavy_sheet_cache[sheet_name] = heavy_cache

    @lru_cache(maxsize=20000)
    def _is_merged_child_cached(self, sheet_name: str, cell_coord: str) -> bool:
        """Vérification avec cache LRU"""
        metadata = self._sheet_metadata.get(sheet_name)
        if not metadata:
            return False
        
        if cell_coord in metadata.merged_masters:
            return False
        
        # Vérification rapide bitmap pour sheets lourdes
        if sheet_name in self.HEAVY_SHEETS:
            cache_key = f"{sheet_name}:{cell_coord}"
            if cache_key in self._merged_bitmap:
                return self._merged_bitmap[cache_key]
        
        ws = self.wb[sheet_name]
        for merged_range in ws.merged_cells.ranges:
            if cell_coord in merged_range:
                if sheet_name in self.HEAVY_SHEETS:
                    self._merged_bitmap[cache_key] = True
                return True
        
        if sheet_name in self.HEAVY_SHEETS:
            self._merged_bitmap[cache_key] = False
        return False

    def _precompute_balance_amounts(
        self,
        accounts: Dict[str, NormalizedBalanceRow],
        slices: Dict[str, Dict[str, Decimal]],
        year_hint: str
    ) -> PrecomputedBalance:
        """Pré-calcule tous les montants"""
        amounts = {}
        amounts_by_prefix = {}
        
        for compte, row in accounts.items():
            # Calculer pour debit, credit, et any
            for col_type in ['debit', 'credit', None]:
                amount = self._resolve_amount_fast(
                    row,
                    col_type,
                    year_hint,
                    'final',
                    slices
                )
                amounts[(compte, col_type or 'any')] = amount
                
                # Aussi pour opening et movement
                if 'opening_signed' in slices.get(compte, {}):
                    opening = self._resolve_amount_fast(
                        row, col_type, year_hint, 'opening', slices
                    )
                    amounts[(compte, f'{col_type or "any"}_opening')] = opening
                
                if 'movement_debit' in slices.get(compte, {}):
                    movement = self._resolve_amount_fast(
                        row, col_type, year_hint, 'movement_debit', slices
                    )
                    amounts[(compte, f'{col_type or "any"}_movement')] = movement
                
                if 'movement_credit' in slices.get(compte, {}):
                    movement_credit = self._resolve_amount_fast(
                        row, col_type, year_hint, 'movement_credit', slices
                    )
                    amounts[(compte, f'{col_type or "any"}_movement_credit')] = movement_credit
            
            # Pré-calcul par préfixe
            prefix = compte[:2] if len(compte) >= 2 else compte
            if prefix not in amounts_by_prefix:
                amounts_by_prefix[prefix] = {}
            amounts_by_prefix[prefix][col_type or 'any'] = amounts_by_prefix[prefix].get(col_type or 'any', Decimal('0')) + amounts[(compte, col_type or 'any')]
        
        return PrecomputedBalance(
            by_compte=accounts,
            amounts=amounts,
            amounts_by_prefix=amounts_by_prefix,
            slices=slices
        )

    def _resolve_amount_fast(
        self,
        normalized_row: NormalizedBalanceRow,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        slices: Dict[str, Dict[str, Decimal]]
    ) -> Decimal:
        """Version rapide de _resolve_amount"""
        compte = normalized_row.compte
        account_slices = slices.get(compte, {})
        
        if value_source == 'opening':
            signed = account_slices.get('opening_signed', Decimal('0'))
        elif value_source == 'movement_debit':
            signed = account_slices.get('movement_debit', Decimal('0'))
        elif value_source == 'movement_credit':
            signed = -account_slices.get('movement_credit', Decimal('0'))
        elif value_source == 'variation':
            closing = account_slices.get('closing_signed', normalized_row.solde_final)
            opening = account_slices.get('opening_signed', Decimal('0'))
            signed = closing - opening
        elif value_source == 'flux':
            signed = account_slices.get('flux', Decimal('0'))
        else:  # final
            signed = account_slices.get('closing_signed', normalized_row.solde_final)
        
        if col_type == 'debit':
            return signed if signed > 0 else Decimal('0')
        elif col_type == 'credit':
            return -signed if signed < 0 else Decimal('0')
        else:
            return signed.copy_abs() if signed != 0 else Decimal('0')

    def fill(self) -> int:
        """Main filling logic optimisée"""
        if self.wb is None or not self.balance_accounts:
            raise RuntimeError("Filler not loaded. Call load() first.")

        assignments_count = 0
        EXCLUDED_SHEETS = {
            "ENTETE", "ENTÊTE", "Fiche R1", "R1", "Fiche R2", "R2",
            "Fiche R3", "R3", "PAGE DE GARDE", "INFORMATIONS GENERALES",
            "INFORMATIONS GÉNÉRALES", "SOMMAIRE"
        }
        excluded_sheets = EXCLUDED_SHEETS | self._skip_sheets

        for sheet_name in self.wb.sheetnames:
            if sheet_name in excluded_sheets:
                if sheet_name in self._skip_sheets:
                    logger.info(f"Skipping sheet (fast filler prefilled): {sheet_name}")
                else:
                    logger.info(f"Skipping sheet: {sheet_name}")
                continue

            metadata = self._sheet_metadata.get(sheet_name)
            if not metadata:
                logger.warning(f"No metadata for sheet {sheet_name}, skipping")
                continue

            logger.info(f"Processing sheet: {sheet_name} (rows: {metadata.max_row}, cols: {metadata.max_col})")
            
            # Réinitialiser le compteur pour cette sheet
            self._previous_assignments_count[sheet_name] = len(self.assignments)

            # Traitement adaptatif selon la charge
            is_heavy = sheet_name in self.HEAVY_SHEETS or "NOTE" in sheet_name.upper()
            
            if is_heavy:
                count = self._process_heavy_sheet(sheet_name, metadata)
            elif metadata.has_heavy_merges or metadata.max_row > self.PARALLEL_THRESHOLD:
                count = self._process_sheet_parallel(sheet_name, metadata)
            else:
                count = self._process_sheet_sequential(sheet_name, metadata)
            
            assignments_count += count
            logger.info(f"  → {count} assignments for {sheet_name}")

        logger.info("Applying column formulas...")
        self._apply_column_formulas()
        logger.info(f"Applied {self.formulas_applied} column formulas")
        
        return assignments_count

    def _process_heavy_sheet(self, sheet_name: str, metadata: SheetMetadata) -> int:
        """Traitement spécialisé pour sheets lourdes (20, 34, etc.)"""
        ws = self.wb[sheet_name]
        has_formula_detector = sheet_name in self.formula_detectors
        formula_detector = self.formula_detectors.get(sheet_name) if has_formula_detector else None
        
        is_bilan = self._canonical_sheet_name(sheet_name) == "BILAN PAYSAGE"
        heavy_cache = self._heavy_sheet_cache.get(sheet_name, HeavySheetCache())
        
        # Filtrer les lignes rapidement
        valid_rows = []
        for row_num in range(11, metadata.max_row + 1):
            if row_num not in metadata.row_labels:
                continue
            
            row_norm = metadata.row_norms[row_num]
            
            # Skip rapide pour lignes non-numériques
            if self._is_non_numeric_row(row_norm):
                continue
            
            valid_rows.append(row_num)
        
        logger.info(f"  {sheet_name}: {len(valid_rows)} numeric rows to process")
        
        # Traiter par petits lots
        assignments_count = 0
        batch_size = self.HEAVY_BATCH_SIZE
        
        for i in range(0, len(valid_rows), batch_size):
            batch_rows = valid_rows[i:i + batch_size]
            batch_assignments = self._process_heavy_batch(
                sheet_name, ws, metadata, batch_rows, 
                formula_detector, is_bilan, heavy_cache
            )
            
            # Écrire les résultats
            for assignment, need in batch_assignments:
                with self._assignments_lock:
                    self.assignments.append(assignment)
                    self._write_cell(ws, need, assignment)
                    self.account_usage.update(assignment.source_accounts)
                    assignments_count += 1
        
        return assignments_count

    def _process_heavy_batch(self, sheet_name: str, ws, metadata: SheetMetadata, 
                            rows: List[int], formula_detector, is_bilan: bool,
                            heavy_cache: HeavySheetCache) -> List[Tuple[CellAssignment, CellNeed]]:
        """Traite un lot de lignes pour sheet lourde"""
        results = []
        tasks = []
        
        for row_num in rows:
            row_label = metadata.row_labels[row_num]
            row_norm = metadata.row_norms[row_num]
            row_pattern = heavy_cache.row_patterns.get(row_num, {})
            
            if is_bilan:
                row_label_left = row_label
                row_label_right = metadata.row_labels.get(row_num)
            else:
                row_label_default = row_label
            
            for col_num in range(1, metadata.max_col + 1):
                cell_coord = f"{get_column_letter(col_num)}{row_num}"
                
                # Vérification merged cells
                if self._is_merged_child_cached(sheet_name, cell_coord):
                    continue
                
                if formula_detector and formula_detector.has_formula(get_column_letter(col_num)):
                    continue
                
                col_label = metadata.col_labels.get(col_num)
                if not col_label:
                    continue
                
                if is_bilan:
                    current_row_label = row_label_right if col_num >= 9 else row_label_left
                else:
                    current_row_label = row_label_default
                
                if not current_row_label:
                    continue
                
                field = self.inventory.get_field(sheet_name, cell_coord) if self.inventory else None
                
                if not self._is_fillable_cell(ws.cell(row=row_num, column=col_num)):
                    continue
                
                tasks.append((
                    sheet_name, row_num, col_num, cell_coord,
                    current_row_label, col_label, field, row_norm, row_pattern
                ))
        
        # Traiter les tâches en parallèle (peu de workers)
        if tasks:
            with ThreadPoolExecutor(max_workers=2) as executor:
                futures = [executor.submit(self._process_heavy_cell_task, task) for task in tasks]
                
                for future in as_completed(futures):
                    result = future.result(timeout=5)
                    if result:
                        results.append(result)
        
        return results

    def _process_heavy_cell_task(self, task: tuple) -> Optional[Tuple[CellAssignment, CellNeed]]:
        """Traite une cellule pour sheet lourde"""
        try:
            sheet_name, row_num, col_num, cell_coord, row_label, col_label, field, row_norm, row_pattern = task
            
            # Cache key
            cache_key = (row_norm, col_label or '', str(field.id if field else 'no_field'))
            
            if cache_key in self._heavy_results_cache:
                cached = self._heavy_results_cache[cache_key]
                if cached:
                    need = CellNeed(
                        sheet=sheet_name, cell=cell_coord, row_num=row_num,
                        col_num=col_num, row_label=row_label, col_label=col_label
                    )
                    return (cached, need)
                return None
            
            need = CellNeed(
                sheet=sheet_name,
                cell=cell_coord,
                row_num=row_num,
                col_num=col_num,
                row_label=row_label,
                col_label=col_label,
                is_merged=False,
            )
            
            # Assignment avec traitement allégé
            assignment = self._satisfy_cell_need_heavy(need, field, row_norm, row_pattern)
            
            if assignment:
                self._heavy_results_cache[cache_key] = assignment
                return (assignment, need)
            
            self._heavy_results_cache[cache_key] = None
            with self._unmatched_lock:
                self.unmatched_cells.append(need)
            return None
            
        except Exception as e:
            logger.debug(f"Error in heavy cell task: {e}")
            return None

    def _satisfy_cell_need_heavy(
        self,
        need: CellNeed,
        field,
        row_norm: str,
        row_pattern: Dict[str, Any]
    ) -> Optional[CellAssignment]:
        """Version allégée pour sheets lourdes"""
        if not need.row_label:
            return None

        is_total = row_pattern.get('is_total', False) or "total" in row_norm
        preferred_prefixes = row_pattern.get('prefixes')
        
        col_type, year_hint, value_source = self._derive_cell_contract(need, field)
        
        # Règles spécifiques par sheet
        if need.sheet in ["20", "NOTE 3A"]:
            forced_value = self._compute_note3a_targeted_value(need, col_type, year_hint, value_source)
            if forced_value is not None:
                return self._create_derived_assignment(need, forced_value, col_type, year_hint, "note3a_targeted")
        
        if need.sheet in ["34", "NOTE 34"]:
            forced_ebe = self._compute_note34_ebe_from_cr(need, year_hint)
            if forced_ebe is not None:
                return self._create_derived_assignment(need, forced_ebe, col_type, year_hint, "note34_ebe")
            
            forced_flux = self._compute_note34_flux_value(need, year_hint)
            if forced_flux is not None:
                return self._create_derived_assignment(need, forced_flux, col_type, year_hint, "note34_flux")
        
        if need.sheet == "NOTE 3C":
            forced_amort = self._compute_note3c_amortization_value(need, col_type, year_hint)
            if forced_amort is not None:
                return self._create_derived_assignment(need, forced_amort, col_type, year_hint, "note3c_amortization")
        
        if need.sheet == "NOTE 4":
            forced_net = self._compute_note4_total_net_depreciation(need)
            if forced_net is not None:
                return self._create_derived_assignment(need, forced_net, col_type, year_hint, "note4_net_depreciation")
        
        # Utiliser la balance pré-calculée
        precomputed = self._balance_precomputed.get('n' if year_hint != 'n1' else 'n1')
        if not precomputed:
            return None
        
        # Règles métier
        rule = self._find_business_rule(need.sheet, need.row_label)
        if rule:
            matches = self._match_with_business_rules_fast(
                row_norm, need, col_type, year_hint, value_source,
                self._expected_classes_for_sheet(need.sheet, need.cell),
                precomputed
            )
            if matches:
                total_amount = sum(m.amount for m in matches)
                return CellAssignment(
                    sheet=need.sheet,
                    cell=need.cell,
                    row_label=need.row_label,
                    col_label=need.col_label,
                    is_total=is_total,
                    matched_accounts=matches,
                    total_amount=total_amount,
                    source_accounts=[m.compte for m in matches if not m.compte.startswith("__ZERO__")],
                    confidence=matches[0].similarity_score,
                    notes=f"source={value_source}; year={year_hint or 'n'}; heavy_optimized",
                )
        
        # Fallback zéro pour les totaux
        if is_total:
            zero_match = MatchedAccount(
                compte=f"__ZERO__:{need.sheet}",
                label="deterministic_zero",
                amount=Decimal("0"),
                side=col_type or "debit",
                similarity_score=0.60,
            )
            return CellAssignment(
                sheet=need.sheet,
                cell=need.cell,
                row_label=need.row_label,
                col_label=need.col_label,
                is_total=is_total,
                matched_accounts=[zero_match],
                total_amount=Decimal("0"),
                source_accounts=[],
                confidence=0.60,
                notes=f"source={value_source}; year={year_hint or 'n'}; heavy_zero_fallback",
            )
        
        return None

    def _is_non_numeric_row(self, row_norm: str) -> bool:
        """Détection rapide des lignes non-numériques"""
        skip_patterns = {
            'commentaire', 'indiquer', 'detailler', 'descriptif',
            'nature de la creance', 'echeance', 'justifier', 'commenter',
            'cocher la case', 'si possible', 'pour les banques',
            'modes d amortissement', 'designation entite', 'numero d identification',
            'ligue', 'nature du produit', 'plafonne a', 'cocher',
            'avantages', 'delai restant', 'toute variation doit etre commentee',
            'doit etre commentee'
        }
        return any(pattern in row_norm for pattern in skip_patterns)

    def _process_sheet_sequential(self, sheet_name: str, metadata: SheetMetadata) -> int:
        """Traitement séquentiel pour petites feuilles"""
        ws = self.wb[sheet_name]
        has_formula_detector = sheet_name in self.formula_detectors
        formula_detector = self.formula_detectors.get(sheet_name) if has_formula_detector else None
        
        is_bilan = self._canonical_sheet_name(sheet_name) == "BILAN PAYSAGE"
        assignments_count = 0

        for row_num in range(11, metadata.max_row + 1):
            if row_num not in metadata.row_labels:
                continue

            row_label = metadata.row_labels[row_num]
            row_norm = metadata.row_norms[row_num]
            
            if is_bilan:
                row_label_left = row_label
                row_label_right = metadata.row_labels.get(row_num)
            else:
                row_label_default = row_label

            for col_num in range(1, metadata.max_col + 1):
                cell_coord = f"{get_column_letter(col_num)}{row_num}"
                
                if self._is_merged_child_cached(sheet_name, cell_coord):
                    continue

                if formula_detector and formula_detector.has_formula(get_column_letter(col_num)):
                    continue

                col_label = metadata.col_labels.get(col_num)
                if not col_label:
                    continue

                if is_bilan:
                    current_row_label = row_label_right if col_num >= 9 else row_label_left
                else:
                    current_row_label = row_label_default

                if not current_row_label:
                    continue

                field = self.inventory.get_field(sheet_name, cell_coord) if self.inventory else None

                if not self._is_fillable_cell(ws.cell(row=row_num, column=col_num)):
                    continue

                need = CellNeed(
                    sheet=sheet_name,
                    cell=cell_coord,
                    row_num=row_num,
                    col_num=col_num,
                    row_label=current_row_label,
                    col_label=col_label,
                    is_merged=False,
                )

                assignment = self._satisfy_cell_need_fast(need, field, row_norm)
                if assignment:
                    self.assignments.append(assignment)
                    self._write_cell(ws, need, assignment)
                    self.account_usage.update(assignment.source_accounts)
                    assignments_count += 1
                else:
                    self.unmatched_cells.append(need)

        return assignments_count

    def _process_sheet_parallel(self, sheet_name: str, metadata: SheetMetadata) -> int:
        """Traitement parallèle pour grandes feuilles"""
        ws = self.wb[sheet_name]
        has_formula_detector = sheet_name in self.formula_detectors
        formula_detector = self.formula_detectors.get(sheet_name) if has_formula_detector else None
        
        is_bilan = self._canonical_sheet_name(sheet_name) == "BILAN PAYSAGE"
        
        # Préparer les tâches
        tasks = []
        for row_num in range(11, metadata.max_row + 1):
            if row_num not in metadata.row_labels:
                continue

            row_label = metadata.row_labels[row_num]
            row_norm = metadata.row_norms[row_num]
            
            if is_bilan:
                row_label_left = row_label
                row_label_right = metadata.row_labels.get(row_num)
            else:
                row_label_default = row_label

            for col_num in range(1, metadata.max_col + 1):
                cell_coord = f"{get_column_letter(col_num)}{row_num}"
                
                if self._is_merged_child_cached(sheet_name, cell_coord):
                    continue

                if formula_detector and formula_detector.has_formula(get_column_letter(col_num)):
                    continue

                col_label = metadata.col_labels.get(col_num)
                if not col_label:
                    continue

                if is_bilan:
                    current_row_label = row_label_right if col_num >= 9 else row_label_left
                else:
                    current_row_label = row_label_default

                if not current_row_label:
                    continue

                field = self.inventory.get_field(sheet_name, cell_coord) if self.inventory else None

                if not self._is_fillable_cell(ws.cell(row=row_num, column=col_num)):
                    continue

                tasks.append((
                    sheet_name, row_num, col_num, cell_coord,
                    current_row_label, col_label, field, row_norm
                ))

        # Traiter par lots
        assignments_count = 0
        with ThreadPoolExecutor(max_workers=4) as executor:
            for i in range(0, len(tasks), self.BATCH_SIZE):
                batch = tasks[i:i + self.BATCH_SIZE]
                futures = [executor.submit(self._process_cell_task, task) for task in batch]
                
                for future in as_completed(futures):
                    result = future.result()
                    if result:
                        assignment, need = result
                        with self._assignments_lock:
                            self.assignments.append(assignment)
                            self._write_cell(ws, need, assignment)
                            self.account_usage.update(assignment.source_accounts)
                            assignments_count += 1

        return assignments_count

    def _process_cell_task(self, task: tuple) -> Optional[Tuple[CellAssignment, CellNeed]]:
        """Traite une cellule individuelle"""
        sheet_name, row_num, col_num, cell_coord, row_label, col_label, field, row_norm = task
        
        need = CellNeed(
            sheet=sheet_name,
            cell=cell_coord,
            row_num=row_num,
            col_num=col_num,
            row_label=row_label,
            col_label=col_label,
            is_merged=False,
        )
        
        assignment = self._satisfy_cell_need_fast(need, field, row_norm)
        if assignment:
            return (assignment, need)
        else:
            with self._unmatched_lock:
                self.unmatched_cells.append(need)
            return None

    def _satisfy_cell_need_fast(
        self,
        need: CellNeed,
        field,
        row_norm: str
    ) -> Optional[CellAssignment]:
        """Version optimisée de _satisfy_cell_need"""
        if not need.row_label:
            return None

        is_total = (
            "total" in row_norm
            or "somme" in row_norm
            or self._looks_like_section_heading(need.row_label)
        )
        
        col_type, year_hint, value_source = self._derive_cell_contract(need, field)
        
        # Cache key
        cache_key = (
            self._sheet_group(need.sheet),
            row_norm,
            (need.col_label or "").lower(),
            col_type or "",
            year_hint or "",
            value_source
        )
        
        with self._cache_lock:
            if cache_key in self._need_cache:
                cached = self._need_cache[cache_key]
                if cached is None:
                    return None
                return replace(cached, cell=need.cell)

        # Valeurs dérivées
        forced_percentage = self._compute_percentage_value(need, col_type, year_hint, value_source)
        if forced_percentage is not None:
            assignment = self._create_derived_assignment(need, forced_percentage, col_type, year_hint, "percentage")
            with self._cache_lock:
                self._need_cache[cache_key] = assignment
            return assignment

        forced_movement = self._compute_movement_value(need, col_type, year_hint, value_source)
        if forced_movement is not None:
            assignment = self._create_derived_assignment(need, forced_movement, col_type, year_hint, "movement")
            with self._cache_lock:
                self._need_cache[cache_key] = assignment
            return assignment

        forced_note3c = self._compute_note3c_amortization_value(need, col_type, year_hint)
        if forced_note3c is not None:
            assignment = self._create_derived_assignment(need, forced_note3c, col_type, year_hint, "note3c_amortization")
            with self._cache_lock:
                self._need_cache[cache_key] = assignment
            return assignment

        forced_note3a = self._compute_note3a_targeted_value(need, col_type, year_hint, value_source)
        if forced_note3a is not None:
            assignment = self._create_derived_assignment(need, forced_note3a, col_type, year_hint, "note3a_targeted")
            with self._cache_lock:
                self._need_cache[cache_key] = assignment
            return assignment

        forced_bilan = self._compute_bilan_targeted_value(need, col_type, year_hint, value_source)
        if forced_bilan is not None:
            assignment = self._create_derived_assignment(need, forced_bilan, col_type, year_hint, "bilan_targeted")
            with self._cache_lock:
                self._need_cache[cache_key] = assignment
            return assignment

        forced_cr = self._compute_cr_targeted_value(need, col_type, year_hint, value_source)
        if forced_cr is not None:
            assignment = self._create_derived_assignment(need, forced_cr, col_type, year_hint, "cr_targeted")
            with self._cache_lock:
                self._need_cache[cache_key] = assignment
            return assignment

        # Matching avec balances
        expected_classes = self._expected_classes_for_sheet(need.sheet, need.cell)
        preferred_prefixes = self._preferred_prefixes_for_label(need.row_label)
        
        balance_key = 'n1' if year_hint == 'n1' else 'n'
        precomputed = self._balance_precomputed.get(balance_key)
        
        if not precomputed and year_hint == 'n1' and self.allow_n1_fallback_without_prev:
            precomputed = self._balance_precomputed.get('n')
        
        if not precomputed:
            with self._cache_lock:
                self._need_cache[cache_key] = None
            return None

        # Matching
        matches = self._match_explicit_account_codes_fast(
            row_norm, need, col_type, year_hint, value_source,
            expected_classes, precomputed
        )
        
        if not matches:
            matches = self._fuzzy_match_accounts_fast(
                row_norm, need, col_type, year_hint, value_source,
                expected_classes, preferred_prefixes, precomputed
            )
        
        if not matches:
            matches = self._match_with_business_rules_fast(
                row_norm, need, col_type, year_hint, value_source,
                expected_classes, precomputed
            )

        if not matches:
            derived = self._build_derived_assignment(need, col_type, year_hint, value_source)
            with self._cache_lock:
                self._need_cache[cache_key] = derived
            return derived

        total_amount = sum(m.amount for m in matches)
        if total_amount <= 0 and not self._is_zero_placeholder_match(matches):
            with self._cache_lock:
                self._need_cache[cache_key] = None
            return None

        source_accounts = [m.compte for m in matches if not m.compte.startswith("__ZERO__")]
        notes = f"source={value_source}; year={year_hint or 'n'}"
        
        assignment = CellAssignment(
            sheet=need.sheet,
            cell=need.cell,
            row_label=need.row_label,
            col_label=need.col_label,
            is_total=is_total,
            matched_accounts=matches,
            total_amount=total_amount,
            source_accounts=source_accounts,
            confidence=matches[0].similarity_score if matches else 0.0,
            notes=notes,
        )

        with self._cache_lock:
            self._need_cache[cache_key] = assignment
        
        return assignment

    def _compute_percentage_value(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[Decimal]:
        """Calcule les valeurs de pourcentage"""
        col_norm = self._normalize_text(need.col_label or "")
        row_norm = self._normalize_text(need.row_label or "")
        
        is_percentage = (
            "%" in (need.col_label or "") or
            "ratio" in col_norm or
            "taux" in col_norm or
            "pourcentage" in col_norm
        )
        
        if not is_percentage:
            return None
        
        ws = self.wb[need.sheet]
        
        numerator_row = None
        denominator_row = None
        
        for row_idx in range(11, need.row_num):
            current_label = self._normalize_text(self._extract_row_label(ws, row_idx) or "")
            if not current_label:
                continue
            
            if "total" in current_label and any(x in row_norm for x in ["marge", "taux"]):
                denominator_row = row_idx
            elif any(x in current_label for x in ["resultat", "benefice", "marge"]) and "taux" in row_norm:
                numerator_row = row_idx
        
        if numerator_row and denominator_row:
            numerator = self._as_decimal_strict(ws.cell(row=numerator_row, column=need.col_num).value)
            denominator = self._as_decimal_strict(ws.cell(row=denominator_row, column=need.col_num).value)
            
            if numerator is not None and denominator is not None and denominator != 0:
                percentage = (numerator / denominator) * 100
                return percentage
        
        return Decimal("0")

    def _compute_movement_value(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[Decimal]:
        """Calcule les valeurs de mouvement"""
        col_norm = self._normalize_text(need.col_label or "")
        
        is_movement = (
            "variation" in col_norm or
            "mouvement" in col_norm or
            "flux" in col_norm or
            "ecart" in col_norm
        )
        
        if not is_movement:
            return None
        
        balance_key = 'n1' if year_hint == 'n1' else 'n'
        precomputed = self._balance_precomputed.get(balance_key)
        
        if not precomputed:
            return None
        
        if "variation" in col_norm or "ecart" in col_norm:
            source = "variation"
        elif "flux" in col_norm:
            source = "flux"
        elif any(x in col_norm for x in ["augmentation", "acquisition", "entree"]):
            source = "movement_debit"
        elif any(x in col_norm for x in ["diminution", "cession", "sortie"]):
            source = "movement_credit"
        else:
            source = "variation"
        
        prefixes = self._preferred_prefixes_for_label(need.row_label or "")
        if not prefixes:
            return None
        
        total = Decimal("0")
        for compte, row in precomputed.by_compte.items():
            if any(compte.startswith(p) for p in prefixes):
                amount_key = (compte, col_type or 'any')
                amount = precomputed.amounts.get(amount_key, Decimal('0'))
                total += amount
        
        return total

    def _match_explicit_account_codes_fast(
        self,
        row_norm: str,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        expected_classes: Optional[Set[str]],
        precomputed: PrecomputedBalance
    ) -> List[MatchedAccount]:
        """Matching explicite optimisé"""
        account_codes: Set[str] = set(re.findall(r"\bcompte\s*(\d{2,6})\b", row_norm))
        if re.fullmatch(r"\d{2,6}", row_norm):
            account_codes.add(row_norm)
        
        if not account_codes:
            return []

        candidates: Dict[str, MatchedAccount] = {}
        for code in sorted(account_codes, key=len, reverse=True):
            for compte, normalized_row in precomputed.by_compte.items():
                if not compte.startswith(code):
                    continue
                if expected_classes and normalized_row.classe not in expected_classes:
                    continue
                
                amount_key = (compte, col_type or 'any')
                amount = precomputed.amounts.get(amount_key, Decimal('0'))
                
                if amount <= 0:
                    continue
                
                score = 0.92 if len(code) >= 3 else 0.82
                candidates[compte] = MatchedAccount(
                    compte=compte,
                    label=normalized_row.label,
                    amount=amount,
                    side=col_type or normalized_row.natural_side,
                    similarity_score=score,
                )

        if not candidates:
            return []
        
        matches = sorted(candidates.values(), key=lambda m: (-float(m.amount), m.compte))
        matches = self._drop_hierarchical_parents(matches)
        return matches[:12]

    def _fuzzy_match_accounts_fast(
        self,
        row_norm: str,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        expected_classes: Optional[Set[str]],
        preferred_prefixes: Optional[Set[str]],
        precomputed: PrecomputedBalance
    ) -> List[MatchedAccount]:
        """Fuzzy matching ultra-optimisé"""
        sheet_group = self._sheet_group(need.sheet)
        row_tokens = self._tokenize(need.row_label)
        
        cache_key = (sheet_group, row_norm, col_type or "", year_hint or "")
        
        with self._cache_lock:
            if cache_key in self._row_match_cache:
                cached = self._row_match_cache[cache_key]
                updated = []
                for match in cached:
                    amount_key = (match.compte, col_type or 'any')
                    amount = precomputed.amounts.get(amount_key, Decimal('0'))
                    if amount > 0:
                        updated.append(replace(match, amount=amount))
                if updated:
                    return updated

        matches: List[Tuple[float, MatchedAccount]] = []
        
        if "total" in row_norm or "somme" in row_norm:
            significance_threshold = Decimal("100")
        else:
            significance_threshold = Decimal("1000")
        
        for compte, normalized_row in precomputed.by_compte.items():
            if expected_classes and normalized_row.classe not in expected_classes:
                continue
            
            amount_key = (compte, col_type or 'any')
            amount = precomputed.amounts.get(amount_key, Decimal('0'))
            
            if amount <= 0 or abs(amount) < significance_threshold:
                continue
            
            score = self._score_account_match_fast(
                row_norm,
                row_tokens,
                normalized_row,
                compte
            )
            
            score -= min(0.18, self.account_usage.get(compte, 0) * 0.02)
            
            if preferred_prefixes and any(compte.startswith(p) for p in preferred_prefixes):
                score += 0.08
            
            if score < self.min_match_score:
                continue
            
            matches.append((
                score,
                MatchedAccount(
                    compte=compte,
                    label=normalized_row.label,
                    amount=amount,
                    side=col_type or normalized_row.natural_side,
                    similarity_score=score,
                )
            ))

        matches.sort(key=lambda x: (-x[0], x[1].compte))
        
        if not matches:
            return []
        
        if len(matches) > 1 and "total" not in row_norm:
            gap = matches[0][0] - matches[1][0]
            if gap < 0.07 and matches[0][0] < 0.86:
                return []
        
        result = [matches[0][1]]
        
        with self._cache_lock:
            self._row_match_cache[cache_key] = result
        
        return result

    def _score_account_match_fast(
        self,
        row_norm: str,
        row_tokens: Set[str],
        normalized_row: NormalizedBalanceRow,
        compte: str
    ) -> float:
        """Scoring optimisé"""
        acc_norm = self._get_account_label_norm(compte, normalized_row.label)
        account_tokens = self._get_account_label_tokens(compte, normalized_row.label)
        
        fuzzy = self._fuzzy_similarity_cached(row_norm, acc_norm)
        overlap = self._token_overlap(row_tokens, account_tokens)
        
        if "total" in row_norm or "somme" in row_norm:
            fuzzy_weight = 0.40
            overlap_weight = 0.60
        else:
            fuzzy_weight = 0.58
            overlap_weight = 0.42
        
        score = (fuzzy_weight * fuzzy) + (overlap_weight * overlap)
        
        if "incorporelle" in row_norm and "corporelle" in acc_norm and "incorporelle" not in acc_norm:
            score -= 0.20
        if "corporelle" in row_norm and "incorporelle" in acc_norm and "corporelle" not in acc_norm:
            score -= 0.20
        
        penalty = self._forbidden_cross_penalty(row_norm, acc_norm)
        if penalty < 0:
            score += penalty
        
        return score

    @lru_cache(maxsize=10000)
    def _fuzzy_similarity_cached(self, str1: str, str2: str) -> float:
        """Fuzzy similarity avec cache"""
        return SequenceMatcher(None, str1, str2).ratio()

    def _match_with_business_rules_fast(
        self,
        row_norm: str,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        expected_classes: Optional[Set[str]],
        precomputed: PrecomputedBalance
    ) -> List[MatchedAccount]:
        """Règles métier optimisées"""
        rule = self._find_business_rule(need.sheet, need.row_label)
        if not rule:
            return []

        prefixes = set(rule.prefixes)
        classes = set(rule.classes)
        row_tokens = self._tokenize(need.row_label)
        
        effective_classes: Optional[Set[str]] = None
        if expected_classes and classes:
            effective_classes = set(expected_classes) & classes
        elif expected_classes:
            effective_classes = set(expected_classes)
        elif classes:
            effective_classes = classes
        
        effective_value_source = rule.value_source_override or value_source
        current_group = self._sheet_group(need.sheet)

        candidates: List[Tuple[float, MatchedAccount]] = []
        
        if rule.aggregate:
            for prefix in prefixes:
                if prefix in precomputed.amounts_by_prefix:
                    amount = precomputed.amounts_by_prefix[prefix].get(col_type or 'any', Decimal('0'))
                    if amount > 0:
                        for compte, normalized_row in precomputed.by_compte.items():
                            if compte.startswith(prefix):
                                score = 0.65
                                candidates.append((
                                    score,
                                    MatchedAccount(
                                        compte=f"__AGGREGATE__:{prefix}",
                                        label=f"Agrégé {prefix}",
                                        amount=amount,
                                        side=col_type or normalized_row.natural_side,
                                        similarity_score=score,
                                    )
                                ))
                                break
            if candidates:
                return [c[1] for c in sorted(candidates, key=lambda x: -x[0])[:rule.max_accounts]]
        
        for compte, normalized_row in precomputed.by_compte.items():
            if effective_classes and normalized_row.classe not in effective_classes:
                continue
            if prefixes and not any(compte.startswith(prefix) for prefix in prefixes):
                continue
            
            amount_key = (compte, col_type or 'any')
            amount = precomputed.amounts.get(amount_key, Decimal('0'))
            
            if amount <= 0 and rule.allow_stock_fallback_when_zero:
                fallback_key = (compte, f"{col_type or 'any'}_opening")
                amount = precomputed.amounts.get(fallback_key, Decimal('0'))
            
            if amount <= 0 and year_hint == 'n1' and rule.allow_n1_fallback:
                fallback_key = (compte, col_type or 'any')
                amount = self._balance_precomputed.get('n', {}).amounts.get(fallback_key, Decimal('0'))
            
            if amount <= 0:
                continue

            score = self._score_account_match_fast(
                row_norm,
                row_tokens,
                normalized_row,
                compte
            )
            
            if prefixes and any(compte.startswith(prefix) for prefix in prefixes):
                score += 0.08
            
            score = min(1.0, score)
            min_required = max(0.50, rule.min_score - 0.10) if rule.aggregate else rule.min_score
            
            if rule.aggregate and prefixes and current_group in {"FLUX", "NOTE13", "NOTE3", "NOTE4", "NOTE28", "NOTE34"}:
                min_required = min(min_required, 0.0)
            
            if score < min_required:
                continue

            if current_group in {"FLUX", "NOTE13", "NOTE3", "NOTE4", "NOTE28", "NOTE34"} and rule.aggregate and prefixes:
                score = max(score, 0.55)

            candidates.append((
                score,
                MatchedAccount(
                    compte=compte,
                    label=normalized_row.label,
                    amount=amount,
                    side=col_type or normalized_row.natural_side,
                    similarity_score=score,
                )
            ))

        if not candidates and rule.allow_zero_fill:
            pseudo = MatchedAccount(
                compte=f"__ZERO__:{'|'.join(rule.prefixes) if rule.prefixes else 'rule'}",
                label="deterministic_zero",
                amount=Decimal("0"),
                side=col_type or "debit",
                similarity_score=max(0.55, rule.min_score),
            )
            return [pseudo]

        if not candidates:
            return []

        if rule.aggregate:
            candidates.sort(key=lambda item: (-float(item[1].amount), -item[0], item[1].compte))
            selected = [m for _, m in candidates]
            if any(k in row_norm for k in ("depreciation", "provision")):
                selected = self._drop_hierarchical_parents(selected)
            return selected[:rule.max_accounts]

        candidates.sort(key=lambda item: (-item[0], item[1].compte))
        if len(candidates) > 1:
            gap = candidates[0][0] - candidates[1][0]
            if gap < 0.05 and candidates[0][0] < 0.90:
                return []
        
        return [candidates[0][1]]

    def _create_derived_assignment(
        self,
        need: CellNeed,
        amount: Decimal,
        col_type: Optional[str],
        year_hint: Optional[str],
        source_type: str
    ) -> CellAssignment:
        """Crée un assignment dérivé"""
        derived_match = MatchedAccount(
            compte=f"__DERIVED__:{source_type.upper()}",
            label=source_type,
            amount=amount,
            side=col_type or "debit",
            similarity_score=0.90,
        )
        
        return CellAssignment(
            sheet=need.sheet,
            cell=need.cell,
            row_label=need.row_label or "",
            col_label=need.col_label or "",
            is_total=("total" in self._normalize_text(need.row_label or "")),
            matched_accounts=[derived_match],
            total_amount=amount,
            source_accounts=[],
            confidence=0.90,
            notes=f"source=final; year={year_hint or 'n'}; derived={source_type}",
        )

    @lru_cache(maxsize=10000)
    def _get_account_label_norm(self, compte: Optional[str], label: str) -> str:
        """Cache pour labels normalisés"""
        return self._normalize_text(label)

    @lru_cache(maxsize=10000)
    def _get_account_label_tokens(self, compte: Optional[str], label: str) -> frozenset:
        """Cache pour tokens"""
        return frozenset(self._tokenize(label))

    @property
    def normalized_rows(self) -> list:
        return list(self.balance_accounts.values())

    @property
    def previous_normalized_rows(self) -> list:
        return list(self.previous_balance_accounts.values())

    def _extract_row_label_for_cell(self, ws, row_num: int, col_num: int) -> Optional[str]:
        if self._canonical_sheet_name(ws.title) != "BILAN PAYSAGE":
            return self._extract_row_label(ws, row_num)

        primary_col = 9 if col_num >= 9 else 2
        primary = ws.cell(row=row_num, column=primary_col).value
        col_a = ws.cell(row=row_num, column=1).value

        p_text = primary.strip() if isinstance(primary, str) else ""
        if p_text and not self._is_structural_text(self._normalize_text(p_text)):
            return p_text

        a_text = col_a.strip() if isinstance(col_a, str) else ""
        if a_text and not re.fullmatch(r"[A-Z0-9]{1,4}", a_text):
            return a_text

        if p_text:
            return p_text
        if a_text:
            return a_text
        return None

    def _is_inventory_writable(self, sheet: str, cell_ref: str, col_label: Optional[str]) -> bool:
        if not self.inventory:
            return True
        sheet_group = self._sheet_group(sheet)
        col_letter = re.match(r"[A-Z]+", cell_ref).group(0) if re.match(r"[A-Z]+", cell_ref) else ""
        field = getattr(self.inventory, "get_field_with_fallback", self.inventory.get_field)(
            sheet, cell_ref
        )
        if not field:
            if sheet_group in {"CF1", "CF2"}:
                return col_letter not in {"A", "B", "C"}
            if sheet_group == "NOTE34":
                return col_letter in {"B", "C"}
            sheet_upper = self._canonical_sheet_name(sheet)
            if "NOTE" in sheet_upper and col_letter and col_letter >= "D":
                return True
            return False
        if field.has_formula:
            return False

        if field.locked:
            if sheet_group in {"CF1", "CF2"}:
                return field.column_letter not in {"A", "B", "C"}
            if sheet_group == "NOTE34":
                return field.column_letter in {"B", "C"}
            sheet_upper = self._canonical_sheet_name(sheet)
            if "NOTE" not in sheet_upper and "SYNTHESE" not in sheet_upper and "SYNTH" not in sheet_upper:
                return False

        required_type = self._infer_required_column_type(col_label)
        if required_type and field.column_type and field.column_type != required_type:
            return False
        return True

    def _infer_required_column_type(self, col_label: Optional[str]) -> Optional[str]:
        if not col_label:
            return None
        label = col_label.lower()
        if "n-1" in label or "n - 1" in label or "n- 1" in label or label == "exercice_n1":
            return "exercice_n1"
        if any(k in label for k in ("precedent", "precedant", "annee prec", "cloture prec")):
            return "exercice_n1"
        if label == "exercice_n" or (("exercice" in label or "exerc" in label) and "n-1" not in label and "n - 1" not in label and "precedent" not in label):
            return "exercice_n"
        return None

    def _is_supported_numeric_target(
        self,
        field,
        row_label: str,
        col_label: str,
        col_num: Optional[int] = None,
        sheet_name: Optional[str] = None,
    ) -> bool:
        if not row_label or not col_label:
            return False

        col_letter = getattr(field, "column_letter", "") or (get_column_letter(col_num) if col_num else "")

        sheet = (getattr(field, "sheet", "") if field else "") or (sheet_name or "")
        sheet_group = self._sheet_group(sheet)
        if col_letter in {"A", "B", "C"} and sheet_group != "NOTE34":
            return False
        row_norm_full = self._normalize_text(row_label)
        if sheet.upper() == "TABLEAU DES FLUX DE TRESORERIE":
            if col_letter not in {"E", "F"}:
                return False
            flux_forbidden = (
                "flux de tresorerie provenant",
                "somme",
                "variation de la tresorerie nette",
                "tresorerie nette au 31 decembre",
                "controle tresorerie",
            )
            if any(k in row_norm_full for k in flux_forbidden):
                return False
            if row_norm_full.startswith("1 a l exclusion") or row_norm_full.startswith("1 a l"):
                return False

        if "NOTE" in sheet.upper():
            note_forbidden = (
                "commentaire",
                "toute variation doit etre commentee",
                "doit etre commentee",
                "detailler",
                "descriptif",
                "indiquer",
                "nature de la creance",
                "echeance",
                "justifier",
                "commenter",
                "cocher la case",
                "si possible",
                "creances du groupe",
                "pour les banques",
                "modes d amortissement",
            )
            if any(k in row_norm_full for k in note_forbidden):
                return False
            if self._looks_like_section_heading(row_label) and "total" not in row_norm_full and "sous total" not in row_norm_full:
                return False

        if sheet.upper().strip() == "NOTE 13":
            if col_letter != "F":
                return False
            note13_forbidden = ("commentaire", "indiquer", "avantages", "delai restant")
            if any(k in row_norm_full for k in note13_forbidden):
                return False

        if sheet_group in {"CF1", "CF2"}:
            cf_forbidden = ("rubriques", "intitules", "periodes", "mois", "declaration annuelle")
            if any(k in row_norm_full for k in cf_forbidden):
                return False
            col_norm = self._normalize_text(col_label)
            if "ligne" in col_norm and "montant" not in col_norm:
                return False

        if sheet_group == "C2_NOTE25":
            if col_letter not in {"J", "K"}:
                return False
            if any(k in row_norm_full for k in ("designation entite", "numero d identification", "ligue", "nature du produit")):
                return False
            if row_norm_full.startswith("le montant du droit d accises") or row_norm_full.startswith("plafonne a"):
                return False

        if sheet_group == "NOTE34":
            if col_letter not in {"B", "C"}:
                return False
            note34_forbidden = (
                "nature des indications",
                "analyse de l activite",
                "solde intermediaires de gestion",
                "analyse de la rentabilite",
                "analyse de la structure financiere",
                "analyse de la variation de la tresorerie",
                "analyse de la variation de l endettement",
                "rentabilite",
                "fonds de roulement",
                "besoin de financement",
                "tresorerie nette",
                "controle tresorerie",
                "variation de la tresorerie nette",
                "endettement financiere net",
                "determin(at|ation) de la capacite",
            )
            if any(k in row_norm_full for k in note34_forbidden if "determin" not in k):
                return False
            if row_norm_full.startswith("detrmination de la capacite") or row_norm_full.startswith("determination de la capacite"):
                return False
            if row_norm_full.startswith("a resultat d exploitation") or row_norm_full.startswith("dettes financieres "):
                return False
            col_clean_note34 = self._normalize_text(col_label)
            if col_clean_note34 not in {"n", "n 1"}:
                return False

        row_clean = self._normalize_text(row_label)
        col_clean = self._normalize_text(col_label)
        row_structural = self._is_structural_text(row_clean)
        if sheet_group in {"CF1", "CF2"} and re.fullmatch(r"\d{2,6}", row_clean):
            row_structural = False
        col_structural = self._is_structural_text(col_clean)
        if sheet_group == "NOTE34" and col_clean in {"n", "n 1"}:
            col_structural = False
        if sheet_group == "NOTE34" and len(row_clean) >= 3:
            row_structural = False
        if row_structural or col_structural:
            return False

        combined = f"{row_clean} {col_clean}"
        forbidden_keywords = {
            "reference", "ref", "libelle", "rubrique", "signe", "code", "notes", "note",
            "designation", "identification", "duree", "unite", "tableau",
        }
        if any(kw in combined for kw in forbidden_keywords):
            return False

        if "%" in col_label or any(kw in col_clean for kw in ("ratio", "taux", "pourcentage")):
            return True

        numeric_hints = (
            "exercice", "montant", "solde", "net", "brut", "debit", "credit", "cloture",
            "ouverture", "variation", "flux", "acquisition", "cession", "virement", "apport",
        )
        if any(kw in combined for kw in numeric_hints):
            return True
        if sheet_group in {"CF1", "CF2"} and any(
            kw in combined for kw in ("base", "impot", "tva", "acompte", "retenue", "precompte", "versement")
        ):
            return True
        if sheet_group == "NOTE34":
            note34_hints = (
                "chiffre d affaires",
                "marge commerciale",
                "valeur ajoute",
                "ebe",
                "resultat",
                "autofinancement",
                "revenus financiers",
                "frais financiers",
                "gains de change",
                "pertes de change",
                "participation",
                "impot sur les resultats",
                "capitaux propres",
                "dettes financieres",
                "actif circulant",
                "passif circulant",
                "tresorerie",
                "endettement",
            )
            if any(k in row_clean for k in note34_hints):
                return True

        if field:
            fmt = str(getattr(field, "number_format", "") or "").replace(" ", "")
            if any(ch in fmt for ch in ("0", "#")):
                return True
            validations = getattr(field, "data_validation", []) or []
            if any((v.get("type") in ("decimal", "whole")) for v in validations if isinstance(v, dict)):
                return True
        return False

    def _extract_row_label(self, ws, row_num: int) -> Optional[str]:
        col_a = ws.cell(row=row_num, column=1).value
        col_b = ws.cell(row=row_num, column=2).value

        b_text = col_b.strip() if isinstance(col_b, str) else ""
        if b_text and not self._is_structural_text(self._normalize_text(b_text)):
            return b_text

        a_text = col_a.strip() if isinstance(col_a, str) else ""
        if a_text and not re.fullmatch(r"[A-Z0-9]{1,4}", a_text):
            return a_text

        if b_text:
            return b_text
        if a_text:
            return a_text
        return None

    def _extract_col_label(self, ws, col_num: int) -> Optional[str]:
        parts: List[str] = []
        seen: Set[str] = set()

        for row in range(1, min(21, (ws.max_row or 21) + 1)):
            cell = ws.cell(row=row, column=col_num)
            val = cell.value
            if val is None and ws.merged_cells:
                coord = f"{get_column_letter(col_num)}{row}"
                for merged in ws.merged_cells.ranges:
                    if coord in merged:
                        val = ws[merged.start_cell.coordinate].value
                        break
            if not isinstance(val, str):
                continue
            val = " ".join(val.strip().split())
            if not val:
                continue
            val_norm = self._normalize_text(val)
            if any(
                noise in val_norm
                for noise in ("designation entite", "numero d identification", "duree", "au 31 decembre")
            ):
                continue
            if len(val_norm) > 120:
                continue
            if val not in seen:
                seen.add(val)
                parts.append(val)

        if not parts:
            return None

        priority = [p for p in parts if any(k in self._normalize_text(p) for k in (
            "exercice", "montant", "solde", "net", "brut", "debit", "credit", "ouverture", "cloture",
            "n-1", "n 1", "precedent", "n-2", "variation", "flux", "pourcentage", "taux", "ratio"
        ))]
        if priority:
            return " | ".join(priority[:2])
        return " | ".join(parts[:2])

    def _is_fillable_cell(self, cell) -> bool:
        if cell.value is not None:
            return False
        if cell.data_type == "f":
            return False
        if cell.data_type == "e":
            return False
        return True

    def _extract_column_type(self, col_label: Optional[str]) -> Optional[str]:
        if not col_label:
            return None

        col_lower = col_label.lower().strip()

        if re.search(r"\b(d[ée]bit|debit)\b", col_lower):
            return "debit"
        if re.search(r"\b(cr[ée]dit|credit)\b", col_lower):
            return "credit"

        if re.search(r"\b(db|dr|d:)\b", col_lower):
            return "debit"
        if re.search(r"\b(cr|c:)\b", col_lower):
            return "credit"

        return None

    def _infer_year_hint(self, col_label: Optional[str]) -> Optional[str]:
        if not col_label:
            return None
        label = self._normalize_text(col_label)
        compact = label.replace(" ", "")
        col_lower = col_label.lower()
        if "n-1" in col_lower or "n - 1" in col_lower or "n- 1" in col_lower:
            return "n1"
        if "n1" in compact or "n1" in label or (col_lower.strip() == "exercice_n1"):
            return "n1"
        if any(k in label for k in (
            "precedent", "precedant", "annee prec", "exercice prec", "cloture prec",
            "anterieur", "antérieur", "ex n-1", "exercice n-1", "solde n-1", "montant n-1",
            "net n-1", "brut n-1", "n 1", "cloture n 1", "ouverture n 1"
        )):
            return "n1"
        if "n-2" in col_lower or "n - 2" in col_lower or "n2" in compact:
            return "n2"
        if col_lower.strip() in ("exercice_n", "n"):
            return "n"
        if "exercice" in label or "n " in label or label.endswith(" n"):
            return "n"
        if "31/12" in col_label or "31 12" in compact:
            if "n-1" not in col_lower and "precedent" not in label:
                return "n"
        if ("net" in label or "montant" in label or "solde" in label) and "n-1" not in col_lower and "precedent" not in label:
            return "n"
        return None

    def _infer_year_hint_from_inventory(self, field) -> Optional[str]:
        if not field:
            return None
        column_type = (getattr(field, "column_type", "") or "").lower().strip()
        if "n-1" in column_type or "n1" in column_type:
            return "n1"
        if "n-2" in column_type or "n2" in column_type:
            return "n2"
        if column_type in {"exercice_n", "n"}:
            return "n"
        return None

    def _derive_cell_contract(self, need: CellNeed, field=None) -> Tuple[Optional[str], Optional[str], str]:
        col_label = need.col_label or ""
        col_type = self._extract_column_type(col_label)
        year_hint = self._infer_year_hint(col_label) or self._infer_year_hint_from_inventory(field)
        value_source = self._infer_value_source_from_column(
            need.sheet,
            need.col_num,
            col_label,
            row_label=need.row_label or "",
        )
        return col_type, year_hint, value_source

    def _source_scores_from_label(self, normalized_label: str) -> Dict[str, int]:
        scores = {
            "opening": 0,
            "final": 0,
            "movement_debit": 0,
            "movement_credit": 0,
            "variation": 0,
            "flux": 0,
        }
        label = normalized_label
        if not label:
            return scores

        opening_kw = ("ouverture", "initial", "debut", "report")
        final_kw = ("cloture", "final", "cumul", "net", "solde")
        plus_kw = ("augmentation", "acquisition", "apport", "creation", "dotation", "reevaluation", "virement")
        minus_kw = ("diminution", "cession", "sortie", "reduction", "remboursement", "prelevement")
        variation_kw = ("variation", "ecart")
        flux_kw = ("flux", "tresorerie")

        for kw in opening_kw:
            if kw in label:
                scores["opening"] += 2
        for kw in final_kw:
            if kw in label:
                scores["final"] += 2
        for kw in plus_kw:
            if kw in label:
                scores["movement_debit"] += 2
        for kw in minus_kw:
            if kw in label:
                scores["movement_credit"] += 2
        for kw in variation_kw:
            if kw in label:
                scores["variation"] += 3
        for kw in flux_kw:
            if kw in label:
                scores["flux"] += 3

        if "mouvement" in label:
            scores["movement_debit"] += 1
            scores["movement_credit"] += 1
        return scores

    def _build_sheet_source_hints(self, sheet_name: str) -> Dict[int, Tuple[str, int]]:
        if sheet_name in self._sheet_source_hints_cache:
            return self._sheet_source_hints_cache[sheet_name]
        ws = self.wb[sheet_name]
        hints: Dict[int, Tuple[str, int]] = {}
        for col_num in range(1, ws.max_column + 1):
            label = self._extract_col_label_cached(ws, col_num)
            if not label:
                continue
            scores = self._source_scores_from_label(self._normalize_text(label))
            source, score = max(scores.items(), key=lambda item: item[1])
            if score > 0:
                hints[col_num] = (source, score)
        self._sheet_source_hints_cache[sheet_name] = hints
        return hints

    def _nearest_strong_sources(self, sheet_name: str, col_num: int, min_score: int = 2) -> Tuple[Optional[str], Optional[str]]:
        hints = self._build_sheet_source_hints(sheet_name)
        left: Optional[str] = None
        right: Optional[str] = None
        left_distance = 10**6
        right_distance = 10**6
        for other_col, (source, score) in hints.items():
            if score < min_score or other_col == col_num:
                continue
            if other_col < col_num:
                dist = col_num - other_col
                if dist < left_distance:
                    left = source
                    left_distance = dist
            else:
                dist = other_col - col_num
                if dist < right_distance:
                    right = source
                    right_distance = dist
        return left, right

    def _infer_value_source(self, col_label: Optional[str]) -> str:
        if not col_label:
            return "final"
        label = self._normalize_text(col_label)
        if any(k in label for k in ("acquisition", "apport", "creation", "augmentation")):
            return "movement_debit"
        if any(k in label for k in ("cession", "sortie", "diminution", "reduction", "scission", "hors service")):
            return "movement_credit"
        if "variation" in label or "ecart" in label:
            return "variation"
        if "flux" in label:
            return "flux"
        if any(k in label for k in ("ouverture", "initial", "n-1", "n1", "exercice_n1")):
            return "opening"
        return "final"

    def _infer_value_source_from_column(
        self,
        sheet_name: str,
        col_num: int,
        col_label: Optional[str],
        row_label: str = "",
    ) -> str:
        base = self._infer_value_source(col_label)
        label_norm = self._normalize_text(col_label or "")
        scores = self._source_scores_from_label(label_norm)
        source, score = max(scores.items(), key=lambda item: item[1])

        if score >= 2:
            chosen = source
        else:
            chosen = base
            left, right = self._nearest_strong_sources(sheet_name, col_num, min_score=2)
            if left == "opening" and right == "final":
                if any(k in label_norm for k in ("diminution", "cession", "sortie", "reduction", "remboursement", "prelevement")):
                    chosen = "movement_credit"
                elif any(k in label_norm for k in ("reevaluation", "dotation", "augmentation", "acquisition", "apport", "virement")):
                    chosen = "movement_debit"

        if "NOTE" in sheet_name.upper():
            row_norm = self._normalize_text(row_label)
            if chosen == "final" and any(k in row_norm for k in ("dotation", "augmentation", "acquisition", "apport", "reevaluation")):
                chosen = "movement_debit"
            if chosen == "final" and any(k in row_norm for k in ("diminution", "cession", "sortie", "remboursement", "prelevement")):
                chosen = "movement_credit"
        return chosen

    def _expected_classes_for_sheet(self, sheet_name: str, cell_ref: str) -> Optional[Set[str]]:
        upper = self._canonical_sheet_name(sheet_name)
        col_letter = re.match(r"[A-Z]+", cell_ref).group(0) if re.match(r"[A-Z]+", cell_ref) else ""
        if upper == "BILAN PAYSAGE":
            if col_letter in {"D", "E", "F", "G"}:
                return {"2", "3", "4", "5"}
            if col_letter in {"K", "L"}:
                return {"1", "4", "5"}
        if upper == "NOTE 34":
            return {"1", "2", "3", "4", "5", "6", "7", "8"}
        if upper.startswith("CF1"):
            return {"1", "4", "5", "6", "7", "8", "9"}
        if upper.startswith("CF2"):
            return {"4", "5", "6", "7"}
        if upper == "C2 NOTE 25":
            return {"4", "5", "6", "7"}
        if upper == "COMPTE DE RESULTAT":
            return {"6", "7"}
        if "NOTE 3" in upper:
            return {"2"}
        if "NOTE 4" in upper:
            return {"2", "5"}
        return None

    def _sheet_cell_domain_filter(self, sheet_name: str, cell_ref: str, row: NormalizedBalanceRow) -> bool:
        upper = self._canonical_sheet_name(sheet_name)
        col_letter = re.match(r"[A-Z]+", cell_ref).group(0) if re.match(r"[A-Z]+", cell_ref) else ""
        if upper == "BILAN PAYSAGE":
            if col_letter in {"D", "E", "F", "G"} and row.classe not in {"2", "3", "4", "5"}:
                return False
            if col_letter in {"K", "L"} and row.classe not in {"1", "4", "5"}:
                return False
        if upper.startswith("CF1") and row.classe not in {"1", "4", "5", "6", "7", "8", "9"}:
            return False
        if upper.startswith("CF2") and row.classe not in {"4", "5", "6", "7"}:
            return False
        if upper == "C2 NOTE 25" and row.classe not in {"4", "5", "6", "7"}:
            return False
        if upper == "COMPTE DE RESULTAT" and row.classe not in {"6", "7"}:
            return False
        return True

    def _canonical_sheet_name(self, sheet_name: str) -> str:
        return " ".join((sheet_name or "").replace("-", " ").upper().split())

    def _sheet_group(self, sheet_name: str) -> str:
        upper = self._canonical_sheet_name(sheet_name)
        if upper == "BILAN PAYSAGE":
            return "BILAN"
        if upper == "COMPTE DE RESULTAT":
            return "CR"
        if upper == "TABLEAU DES FLUX DE TRESORERIE":
            return "FLUX"
        if upper.startswith("CF1"):
            return "CF1"
        if upper.startswith("CF2"):
            return "CF2"
        if upper == "C2 NOTE 25":
            return "C2_NOTE25"
        if upper == "NOTE 13":
            return "NOTE13"
        if upper == "NOTE 28":
            return "NOTE28"
        if upper == "NOTE 34":
            return "NOTE34"
        if upper.startswith("NOTE 3"):
            return "NOTE3"
        if upper.startswith("NOTE 4"):
            return "NOTE4"
        return "OTHER"

    def _build_business_rules(self) -> List[BusinessRule]:
        return [
            BusinessRule(("BILAN", "NOTE3"), ("frais de developpement",), ("201",), ("2",), False, 0.62),
            BusinessRule(("BILAN", "NOTE3"), ("brevet", "logiciel"), ("202",), ("2",), False, 0.62),
            BusinessRule(("BILAN", "NOTE3"), ("fonds commercial",), ("203",), ("2",), False, 0.62),
            BusinessRule(("BILAN", "NOTE3"), ("autres immobilisations incorporelles",), ("204",), ("2",), False, 0.60),
            BusinessRule(("BILAN", "NOTE3"), ("immobilisations incorporelles",), ("20",), ("2",), True, 0.56, 18),
            BusinessRule(("BILAN", "NOTE3"), ("terrains",), ("21", "22"), ("2",), True, 0.56, 10),
            BusinessRule(("BILAN", "NOTE3"), ("batiments",), ("22",), ("2",), True, 0.56, 10),
            BusinessRule(("BILAN", "NOTE3"), ("amenagements", "agencements", "installations"), ("23",), ("2",), True, 0.56, 12),
            BusinessRule(("BILAN", "NOTE3"), ("materiel de transport",), ("245", "24"), ("2",), True, 0.56, 12),
            BusinessRule(("BILAN", "NOTE3"), ("materiel", "mobilier"), ("24", "25"), ("2",), True, 0.56, 16),
            BusinessRule(("BILAN", "NOTE3"), ("avances", "acomptes"), ("23",), ("2",), True, 0.56, 10),
            BusinessRule(("BILAN", "NOTE3", "NOTE4"), ("titres de participation",), ("26",), ("2",), True, 0.56, 12),
            BusinessRule(("NOTE3",), ("sous total", "immobilisations incorporelles"), ("20",), ("2",), True, 0.52, 20),
            BusinessRule(("NOTE3",), ("sous total", "immobilisations corporelles"), ("21", "22", "23", "24", "25"), ("2",), True, 0.52, 24),
            BusinessRule(("NOTE3",), ("immobilisation corporelles",), ("21", "22", "23", "24", "25"), ("2",), True, 0.52, 24),
            BusinessRule(("NOTE3", "NOTE4"), ("total general",), ("20", "21", "22", "23", "24", "25", "26", "27"), ("2",), True, 0.50, 30),
            BusinessRule(("BILAN", "NOTE4"), ("autres immobilisations financieres",), ("27",), ("2",), True, 0.56, 12),
            BusinessRule(("BILAN", "NOTE4"), ("prets et creances",), ("27", "41", "42", "44", "46"), ("2", "4"), True, 0.54, 16),
            BusinessRule(("NOTE4",), ("titres de participation",), ("26",), ("2",), True, 0.56, 12, False, "final"),
            BusinessRule(("NOTE4",), ("prets et creances",), ("27", "278"), ("2",), True, 0.54, 16, False, "final"),
            BusinessRule(("NOTE4",), ("pret au personnel",), ("274",), ("2",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE4",), ("creances sur l etat",), ("276",), ("2",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE4",), ("titres immobilises",), ("273",), ("2",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE4",), ("depots et cautionnements",), ("275",), ("2",), True, 0.56, 12, False, "final"),
            BusinessRule(("NOTE4",), ("interets courus",), ("276",), ("2",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE4",), ("total brut",), ("26", "27"), ("2",), True, 0.52, 24, False, "final"),
            BusinessRule(("NOTE4",), ("depreciations titres de participation",), ("296",), ("2",), True, 0.56, 12, False, "final"),
            BusinessRule(("NOTE4",), ("depreciations autres immobilisations",), ("297", "29"), ("2",), True, 0.54, 14, False, "final", False, True),
            BusinessRule(("BILAN",), ("stocks", "encours"), ("3",), ("3",), True, 0.56, 20),
            BusinessRule(("BILAN",), ("clients",), ("41",), ("4",), True, 0.56, 16),
            BusinessRule(("BILAN",), ("creances", "emplois assimiles"), ("41", "42", "43", "44", "46", "47"), ("4",), True, 0.54, 24),
            BusinessRule(("BILAN",), ("autres creances",), ("46", "47", "48"), ("4",), True, 0.54, 16),
            BusinessRule(("BILAN",), ("titres de placement",), ("50",), ("5",), True, 0.56, 12),
            BusinessRule(("BILAN",), ("valeurs a encaisser",), ("51",), ("5",), True, 0.56, 12),
            BusinessRule(("BILAN",), ("banques", "caisse"), ("52", "53", "54", "56", "57"), ("5",), True, 0.54, 20),
            BusinessRule(("BILAN",), ("total actif immobilise",), ("20", "21", "22", "23", "24", "25", "26", "27"), ("2",), True, 0.52, 24),
            BusinessRule(("BILAN",), ("total actif circulant",), ("3", "4"), ("3", "4"), True, 0.52, 26),
            BusinessRule(("BILAN",), ("total tresorerie actif",), ("5",), ("5",), True, 0.52, 20),
            BusinessRule(("BILAN",), ("total general",), ("2", "3", "4", "5"), ("2", "3", "4", "5"), True, 0.50, 30),
            BusinessRule(("CR",), ("vente de marchandises",), ("70",), ("7",), True, 0.58, 16, True),
            BusinessRule(("CR",), ("vente de produits fabriques",), ("70", "71"), ("7",), True, 0.58, 16, True),
            BusinessRule(("CR",), ("travaux", "services vendus"), ("70", "71"), ("7",), True, 0.58, 16, True),
            BusinessRule(("CR",), ("produits accessoires"), ("70", "71", "72"), ("7",), True, 0.56, 16, True),
            BusinessRule(("CR",), ("chiffre d affaires"), ("70", "71", "72"), ("7",), True, 0.54, 24, True),
            BusinessRule(("CR",), ("produits stockes"), ("73",), ("7",), True, 0.56, 12, True, None, False, True),
            BusinessRule(("CR",), ("production immobilisee"), ("72",), ("7",), True, 0.56, 12, True, None, False, True),
            BusinessRule(("CR",), ("subventions d exploitation"), ("74",), ("7",), True, 0.56, 12, True, None, False, True),
            BusinessRule(("CR",), ("transfert de charges d exploitation"), ("78",), ("7",), True, 0.56, 12, True, None, False, True),
            BusinessRule(("CR",), ("achat de marchandises"), ("60",), ("6",), True, 0.58, 16, True),
            BusinessRule(("CR",), ("achats de matieres premieres"), ("60", "61"), ("6",), True, 0.56, 18, True),
            BusinessRule(("CR",), ("variation de stocks de matieres premieres"), ("603",), ("6",), True, 0.56, 10, True, None, False, True),
            BusinessRule(("CR",), ("autres achats"), ("60", "61", "62"), ("6",), True, 0.56, 18, True),
            BusinessRule(("CR",), ("variation de stock d autres approvisionnement"), ("603",), ("6",), True, 0.56, 10, True, None, False, True),
            BusinessRule(("CR",), ("services exterieurs"), ("62",), ("6",), True, 0.58, 14, True),
            BusinessRule(("CR",), ("transports"), ("62",), ("6",), True, 0.58, 14, True),
            BusinessRule(("CR",), ("impots et taxes"), ("63",), ("6",), True, 0.58, 14, True),
            BusinessRule(("CR",), ("charges de personnel"), ("64",), ("6",), True, 0.58, 14, True),
            BusinessRule(("CR",), ("reprises d amortissements", "depreciations"), ("78",), ("7",), True, 0.56, 14, True, None, False, True),
            BusinessRule(("CR",), ("dotations aux amortissements",), ("68", "69"), ("6",), True, 0.56, 18, True),
            BusinessRule(("CR",), ("revenus financiers",), ("75", "76", "77"), ("7",), True, 0.56, 18, True),
            BusinessRule(("CR",), ("reprises de provisions", "financieres"), ("78",), ("7",), True, 0.56, 12, True, None, False, True),
            BusinessRule(("CR",), ("frais financiers",), ("67",), ("6",), True, 0.56, 14, True),
            BusinessRule(("CR",), ("dotations aux provisions", "financieres"), ("686", "687", "69"), ("6",), True, 0.56, 12, True, None, False, True),
            BusinessRule(("CR",), ("autres produits hao"), ("77",), ("7",), True, 0.56, 12, True, None, False, True),
            BusinessRule(("CR",), ("autres produits",), ("75", "76", "77"), ("7",), True, 0.56, 18, True),
            BusinessRule(("CR",), ("produits et cessions d immobilisations",), ("754",), ("7",), True, 0.56, 10, True, None, False, True),
            BusinessRule(("CR",), ("valeurs comptables de cessiions d immobilisations",), ("654",), ("6",), True, 0.56, 10, True, None, False, True),
            BusinessRule(("CR",), ("participation des travailleurs",), ("69",), ("6",), True, 0.56, 10, True, None, False, True),
            BusinessRule(("CR",), ("impot sur le resultat",), ("89", "69"), ("6", "8"), True, 0.56, 12, True, None, False, True),
            BusinessRule(("CR",), ("autres charges hao"), ("67", "68", "69"), ("6",), True, 0.56, 14, True, None, False, True),
            BusinessRule(("CR",), ("autres charges",), ("65", "66", "67", "68"), ("6",), True, 0.54, 20, True),
            BusinessRule(("FLUX",), ("capacite d autofinancement globale",), ("6", "7"), ("6", "7"), True, 0.56, 26, True),
            BusinessRule(("FLUX",), ("actif circulant hao",), ("4",), ("4",), True, 0.54, 18, False, "variation"),
            BusinessRule(("FLUX",), ("variation des stocks",), ("3",), ("3",), True, 0.56, 18, False, "variation"),
            BusinessRule(("FLUX",), ("variation des creances",), ("41", "42", "43", "44", "46", "47"), ("4",), True, 0.54, 22, False, "variation"),
            BusinessRule(("FLUX",), ("variation du passif circulant",), ("40", "42", "43", "44", "47"), ("4",), True, 0.54, 22, False, "variation"),
            BusinessRule(("FLUX",), ("variation du bf",), ("3", "4"), ("3", "4"), True, 0.52, 26, False, "variation"),
            BusinessRule(("FLUX",), ("decaissements", "acquisitions", "immobilisation incorporelles"), ("20",), ("2",), True, 0.56, 14, False, "movement_debit", False, True),
            BusinessRule(("FLUX",), ("decaissements", "acquisitions", "immobilisation corporelles"), ("21", "22", "23", "24", "25"), ("2",), True, 0.56, 20, False, "movement_debit"),
            BusinessRule(("FLUX",), ("decaissements", "acquisitions", "immobilisation financieres"), ("26", "27"), ("2",), True, 0.56, 16, False, "movement_debit"),
            BusinessRule(("FLUX",), ("encaissement", "cessions", "immobilisations incorporelles"), ("20", "21", "22", "23", "24", "25"), ("2",), True, 0.56, 20, False, "movement_credit", False, True),
            BusinessRule(("FLUX",), ("encaissement", "cessions", "immobilisations financieres"), ("26", "27"), ("2",), True, 0.56, 16, False, "movement_credit", False, True),
            BusinessRule(("FLUX",), ("augmentation de capital",), ("10",), ("1",), True, 0.56, 12, False, "movement_credit", True),
            BusinessRule(("FLUX",), ("subventions d investissement recues",), ("15",), ("1",), True, 0.56, 12, False, "movement_credit", False, True),
            BusinessRule(("FLUX",), ("prelevement sur le capital",), ("10",), ("1",), True, 0.56, 12, False, "movement_debit", True),
            BusinessRule(("FLUX",), ("dividendes verses",), ("13",), ("1",), True, 0.56, 12, False, "movement_debit"),
            BusinessRule(("FLUX",), ("tresorerie provenant du financement par les capitaux etrangers",), ("16", "17", "47"), ("1", "4"), True, 0.52, 18, False, "final"),
            BusinessRule(("FLUX",), ("emprunts",), ("16", "17"), ("1",), True, 0.56, 14, False, "movement_credit", True),
            BusinessRule(("FLUX",), ("autres dettes financieres",), ("16", "17", "47"), ("1", "4"), True, 0.54, 18, False, "movement_credit", True),
            BusinessRule(("FLUX",), ("remboursemement des emprunts",), ("16", "17"), ("1",), True, 0.56, 14, False, "movement_debit", True),
            BusinessRule(("CF1",), ("compte 89",), ("89",), ("8",), True, 0.56, 20, False, "final"),
            BusinessRule(("CF1",), ("impot sur le resultat",), ("89", "44", "69"), ("4", "6", "8"), True, 0.54, 20, False, "final"),
            BusinessRule(("CF1",), ("minimum de perception",), ("44", "89"), ("4", "8"), True, 0.54, 16, False, "final"),
            BusinessRule(("CF1",), ("ircm",), ("44", "47", "75"), ("4", "7"), True, 0.54, 14, False, "final", False, True),
            BusinessRule(("CF1",), ("retenues a la source",), ("44",), ("4",), True, 0.54, 12, False, "final", False, True),
            BusinessRule(("CF1",), ("acomptes",), ("44",), ("4",), True, 0.54, 12, False, "final", False, True),
            BusinessRule(("CF1",), ("centimes additionnels communaux",), ("44",), ("4",), True, 0.56, 10, False, "final", False, True),
            BusinessRule(("CF2",), ("tva brute",), ("44",), ("4",), True, 0.54, 14, False, "movement_credit", False, True),
            BusinessRule(("CF2",), ("tva deductible",), ("44",), ("4",), True, 0.54, 14, False, "movement_debit", False, True),
            BusinessRule(("CF2",), ("tva nette",), ("44",), ("4",), True, 0.54, 16, False, "final", False, True),
            BusinessRule(("CF2",), ("credit de tva",), ("44",), ("4",), True, 0.54, 14, False, "final", False, True),
            BusinessRule(("NOTE34",), ("chiffre d affaires",), ("70", "71", "72"), ("7",), True, 0.54, 20),
            BusinessRule(("NOTE34",), ("revenus financiers",), ("75", "76", "77"), ("7",), True, 0.56, 16, True),
            BusinessRule(("NOTE34",), ("gains de change",), ("776", "777"), ("7",), True, 0.56, 10, False, None, False, True),
            BusinessRule(("NOTE34",), ("frais financiers",), ("67",), ("6",), True, 0.56, 14, True),
            BusinessRule(("NOTE34",), ("pertes de change",), ("676", "677"), ("6",), True, 0.56, 10, False, None, False, True),
            BusinessRule(("NOTE34",), ("participation",), ("69",), ("6",), True, 0.56, 10, False, None, False, True),
            BusinessRule(("NOTE34",), ("impot sur les resultats",), ("89", "69"), ("6", "8"), True, 0.54, 14, True),
            BusinessRule(("NOTE34",), ("valeurs comptables des cessions",), ("654",), ("6",), True, 0.56, 8, False, None, False, True),
            BusinessRule(("NOTE34",), ("produits des cessions courantes d immobilisation",), ("754",), ("7",), True, 0.56, 8, False, None, False, True),
            BusinessRule(("NOTE34",), ("distribution de dividendes",), ("13",), ("1",), True, 0.56, 8, False, "movement_debit", False, True),
            BusinessRule(("NOTE34",), ("capitaux propres",), ("10", "11", "12", "13", "14", "15"), ("1",), True, 0.52, 20),
            BusinessRule(("NOTE34",), ("dettes financieres",), ("16", "17"), ("1",), True, 0.54, 14),
            BusinessRule(("NOTE34",), ("actif immobilise",), ("2",), ("2",), True, 0.54, 20),
            BusinessRule(("NOTE34",), ("actif circulant d exploitation",), ("3", "41", "42", "43", "44", "46", "47"), ("3", "4"), True, 0.52, 24),
            BusinessRule(("NOTE34",), ("passif circulant d exploitation",), ("40", "42", "43", "44", "47"), ("4",), True, 0.52, 20),
            BusinessRule(("NOTE34",), ("actif circulant hao",), ("4",), ("4",), True, 0.54, 14),
            BusinessRule(("NOTE34",), ("passif circulant hao",), ("47",), ("4",), True, 0.54, 14),
            BusinessRule(("NOTE34",), ("tresorerie actif",), ("52", "53", "54", "56", "57"), ("5",), True, 0.54, 14),
            BusinessRule(("NOTE34",), ("tresorerie passif",), ("52", "56"), ("5",), True, 0.54, 10),
            BusinessRule(("NOTE34",), ("endettement financiere net",), ("16", "17", "52", "56"), ("1", "5"), True, 0.52, 18),
            BusinessRule(("NOTE13",), ("apporteurs capital non appele",), ("1013", "101300"), ("1",), True, 0.56, 12, False, "movement_debit", True),
            BusinessRule(("NOTE13",), ("total",), ("1013", "101300"), ("1",), True, 0.52, 14, False, "movement_debit", True),
            BusinessRule(("NOTE28",), ("provisions reglementees",), ("14",), ("1",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE28",), ("provisions financieres",), ("198", "15"), ("1",), True, 0.56, 10, False, "final", False, True),
            BusinessRule(("NOTE28",), ("depreciation des immobilisations",), ("29",), ("2",), True, 0.56, 12, False, "final"),
            BusinessRule(("NOTE28",), ("total dotations",), ("14", "15", "198", "29"), ("1", "2"), True, 0.52, 24, False, "final", False, True),
            BusinessRule(("NOTE28",), ("depreciations des stocks",), ("39",), ("3",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE28",), ("depreciations actif circulant hao",), ("499",), ("4",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE28",), ("depreciations fournisseurs",), ("490",), ("4",), True, 0.56, 10, False, "final"),
            BusinessRule(("NOTE28",), ("depreciations clients",), ("491",), ("4",), True, 0.56, 10, False, "final"),
            BusinessRule(("NOTE28",), ("depreciations autres creances",), ("496", "497"), ("4",), True, 0.56, 12, False, "final"),
            BusinessRule(("NOTE28",), ("depreciations titres de placement",), ("59",), ("5",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE28",), ("depreciations valeurs a encaisser",), ("59",), ("5",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE28",), ("depreciations disponibilite",), ("59",), ("5",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE28",), ("court termes exploitation",), ("499",), ("4",), True, 0.56, 8, False, "final", False, True),
            BusinessRule(("NOTE28",), ("caractere financier",), ("59", "499"), ("4", "5"), True, 0.56, 10, False, "final", False, True),
            BusinessRule(("NOTE28",), ("total charges pour depreciations"), ("39", "49", "59", "499"), ("3", "4", "5"), True, 0.52, 28, False, "final", False, True),
            BusinessRule(("NOTE28",), ("total provisions et depreciations"), ("14", "15", "198", "29", "39", "49", "59", "499"), ("1", "2", "3", "4", "5"), True, 0.50, 32, False, "final", False, True),
        ]

    def _find_business_rule(self, sheet_name: str, row_label: str) -> Optional[BusinessRule]:
        group = self._sheet_group(sheet_name)
        row_norm = self._normalize_text(row_label)
        if len(row_norm) < 4:
            return None
        candidates: List[Tuple[Tuple[int, int, int, int], BusinessRule]] = []
        for idx, rule in enumerate(self.business_rules):
            if group not in rule.sheet_groups:
                continue
            if all(keyword in row_norm for keyword in rule.label_keywords):
                specificity = (
                    len(rule.label_keywords),
                    sum(len(k) for k in rule.label_keywords),
                    len(rule.prefixes),
                    -idx,
                )
                candidates.append((specificity, rule))
        if not candidates:
            return None
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def _drop_hierarchical_parents(self, matches: List[MatchedAccount]) -> List[MatchedAccount]:
        if not matches:
            return matches
        comptes = [m.compte for m in matches]
        filtered: List[MatchedAccount] = []
        for match in matches:
            is_parent_of_selected = any(
                other.startswith(match.compte) and len(other) > len(match.compte)
                for other in comptes
            )
            if is_parent_of_selected:
                continue
            filtered.append(match)
        return filtered or matches

    def _preferred_prefixes_for_label(self, row_label: str) -> Optional[Set[str]]:
        label = self._normalize_text(row_label)
        prefix_map = {
            "immobilisations incorporelles": {"20"},
            "immobilisations corporelles": {"21", "22", "23"},
            "titres de participation": {"24"},
            "stocks": {"3"},
            "creances": {"41"},
            "fournisseurs": {"40"},
            "personnel": {"42"},
            "impots": {"44"},
            "banques": {"52"},
            "caisse": {"53"},
            "chiffre d affaires": {"70"},
            "ventes": {"70"},
            "achat": {"60"},
        }
        explicit_code = re.search(r"\bcompte\s*(\d{2,6})\b", label)
        if explicit_code:
            return {explicit_code.group(1)}
        if re.fullmatch(r"\d{2,6}", label):
            return {label}
        for key, prefixes in prefix_map.items():
            if key in label:
                return set(prefixes)
        return None

    def _looks_like_section_heading(self, row_label: str) -> bool:
        text = row_label.strip()
        if len(text) < 8:
            return False
        alpha = [c for c in text if c.isalpha()]
        if not alpha:
            return False
        upper_ratio = sum(1 for c in alpha if c.isupper()) / len(alpha)
        return upper_ratio > 0.85

    def _is_structural_text(self, text: str) -> bool:
        if not text:
            return True
        if text in {"ref", "note", "libelles", "libelle", "rubriques", "rubrique"}:
            return True
        if re.fullmatch(r"[a-z0-9]{1,4}", text):
            return True
        return False

    def _normalize_text(self, value: str) -> str:
        if not value:
            return ""
        text = unicodedata.normalize("NFKD", str(value))
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = text.lower()
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return " ".join(text.split())

    def _tokenize(self, value: str) -> Set[str]:
        stop_words = {
            "de", "la", "le", "les", "des", "du", "et", "au", "aux", "a", "an", "the",
            "pour", "sur", "par", "en", "d", "l",
        }
        tokens = {tok for tok in self._normalize_text(value).split() if len(tok) > 2 and tok not in stop_words}
        return tokens

    def _token_overlap(self, left: Set[str], right: Set[str]) -> float:
        if not left or not right:
            return 0.0
        inter = len(left & right)
        union = len(left | right)
        return inter / union if union else 0.0

    @staticmethod
    def _forbidden_cross_penalty(row_norm: str, acc_norm: str) -> float:
        row_tokens = set(row_norm.split())
        acc_tokens = set(acc_norm.split())
        for left_terms, right_terms in FORBIDDEN_CROSS_TERMS:
            row_has_left  = bool(left_terms  & row_tokens)
            acc_has_right = bool(right_terms & acc_tokens)
            row_has_right = bool(right_terms & row_tokens)
            acc_has_left  = bool(left_terms  & acc_tokens)
            if (row_has_left and acc_has_right) or (row_has_right and acc_has_left):
                return -0.35
        return 0.0

    def _resolve_amount(
        self,
        normalized_row: NormalizedBalanceRow,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        *,
        slices_map: Optional[Dict[str, Dict[str, Decimal]]] = None,
        use_previous_year_dataset: bool = False,
    ) -> Decimal:
        """Version originale conservée pour compatibilité"""
        raw = normalized_row.metadata.get("raw", {}) if isinstance(normalized_row.metadata, dict) else {}
        slices_source = slices_map if slices_map is not None else self.account_time_slices
        slices = slices_source.get(normalized_row.compte, {})

        opening_debit = slices.get("opening_debit", Decimal("0"))
        opening_credit = slices.get("opening_credit", Decimal("0"))
        movement_debit = slices.get("movement_debit", self._as_decimal(raw.get("debit")))
        movement_credit = slices.get("movement_credit", self._as_decimal(raw.get("credit")))
        closing_debit = slices.get("closing_debit", normalized_row.debit_balance)
        closing_credit = slices.get("closing_credit", normalized_row.credit_balance)
        opening_signed = slices.get("opening_signed", self._as_decimal(raw.get("solde_initial")))
        closing_signed = slices.get("closing_signed", normalized_row.solde_final)
        flux_value = slices.get("flux", self._as_decimal(raw.get("flux_tresorerie")))

        source = value_source

        if source == "opening":
            signed = opening_signed
        elif source == "movement_debit":
            signed = movement_debit
        elif source == "movement_credit":
            signed = -movement_credit
        elif source == "variation":
            signed = closing_signed - opening_signed
        elif source == "flux":
            signed = flux_value
        else:
            signed = closing_signed

        if source in {"movement_debit", "movement_credit"}:
            if col_type == "debit":
                return movement_debit if source == "movement_debit" else Decimal("0")
            if col_type == "credit":
                return movement_credit if source == "movement_credit" else Decimal("0")
            return movement_debit if source == "movement_debit" else movement_credit

        if source == "opening":
            if col_type == "debit":
                return opening_debit
            if col_type == "credit":
                return opening_credit
            return signed.copy_abs() if signed != 0 else Decimal("0")

        if source == "final":
            if col_type == "debit":
                return closing_debit
            if col_type == "credit":
                return closing_credit
            return signed.copy_abs() if signed != 0 else Decimal("0")

        if col_type == "debit":
            return signed if signed > 0 else Decimal("0")
        if col_type == "credit":
            return -signed if signed < 0 else Decimal("0")

        if signed == 0:
            return Decimal("0")
        return signed.copy_abs()

    def _select_balance_dataset_for_year(
        self,
        year_hint: Optional[str],
    ) -> Tuple[Dict[str, NormalizedBalanceRow], Dict[str, Dict[str, Decimal]], bool]:
        if year_hint == "n1":
            if self.previous_balance_accounts:
                return self.previous_balance_accounts, self.previous_account_time_slices, True
            if self.allow_n1_fallback_without_prev:
                return self.balance_accounts, self.account_time_slices, False
            return {}, {}, True
        return self.balance_accounts, self.account_time_slices, False

    def _load_balance_accounts(
        self,
        balance_file: Path,
        column_overrides: Optional[Dict[str, object]],
    ) -> Dict[str, NormalizedBalanceRow]:
        accounts: Dict[str, NormalizedBalanceRow] = {}
        normalizer = BalanceNormalizer(balance_file, column_overrides=column_overrides)
        for normalized_row in normalizer.iterate():
            accounts[normalized_row.compte] = normalized_row
        return accounts

    def _extract_balance_time_slices_from_file(
        self,
        balance_file: Path,
        column_overrides: Optional[Dict[str, object]],
    ) -> Dict[str, Dict[str, Decimal]]:
        wb = load_workbook(balance_file, data_only=True)
        try:
            ws = wb.active
            overrides = column_overrides or {}
            compte_col = int(overrides.get("compte", 1))
            label_col = int(overrides.get("label", 4))

            opening_debit_cols = self._resolve_override_columns("opening_debit_columns", [11], overrides)
            opening_credit_cols = self._resolve_override_columns("opening_credit_columns", [14, 13], overrides)
            movement_debit_cols = self._resolve_override_columns("movement_debit_columns", [17], overrides)
            movement_credit_cols = self._resolve_override_columns("movement_credit_columns", [20], overrides)
            closing_debit_cols = self._resolve_override_columns("closing_debit_columns", [24, 21], overrides)
            closing_credit_cols = self._resolve_override_columns("closing_credit_columns", [27, 26], overrides)

            slices: Dict[str, Dict[str, Decimal]] = {}
            for row_idx in range(13, ws.max_row + 1):
                compte_raw = ws.cell(row=row_idx, column=compte_col).value
                compte = str(compte_raw or "").strip()
                if not compte or not compte[0].isdigit():
                    continue

                label_raw = ws.cell(row=row_idx, column=label_col).value
                label_text = str(label_raw or "").strip()
                if not label_text:
                    continue

                opening_debit = self._first_numeric_cell(ws, row_idx, opening_debit_cols)
                opening_credit = self._first_numeric_cell(ws, row_idx, opening_credit_cols)
                movement_debit = self._first_numeric_cell(ws, row_idx, movement_debit_cols)
                movement_credit = self._first_numeric_cell(ws, row_idx, movement_credit_cols)
                closing_debit = self._first_numeric_cell(ws, row_idx, closing_debit_cols)
                closing_credit = self._first_numeric_cell(ws, row_idx, closing_credit_cols)

                opening_signed = opening_debit - opening_credit
                closing_signed = closing_debit - closing_credit
                flux_val = self._first_numeric_cell(
                    ws,
                    row_idx,
                    self._resolve_override_columns("flux_columns", [], overrides),
                )

                slices[compte] = {
                    "opening_debit": opening_debit,
                    "opening_credit": opening_credit,
                    "movement_debit": movement_debit,
                    "movement_credit": movement_credit,
                    "closing_debit": closing_debit,
                    "closing_credit": closing_credit,
                    "opening_signed": opening_signed,
                    "closing_signed": closing_signed,
                    "flux": flux_val,
                }
            return slices
        finally:
            wb.close()

    def _resolve_override_columns(
        self,
        key: str,
        default: List[int],
        overrides: Optional[Dict[str, object]] = None,
    ) -> List[int]:
        source = overrides if overrides is not None else self.column_overrides
        value = source.get(key, default)
        if value is None:
            return []
        if isinstance(value, int):
            return [value]
        if isinstance(value, (list, tuple, set)):
            return [int(v) for v in value if v]
        try:
            return [int(value)]
        except (TypeError, ValueError):
            return list(default)

    def _first_numeric_cell(self, ws, row_idx: int, columns: List[int]) -> Decimal:
        if not columns:
            return Decimal("0")
        for col_idx in columns:
            try:
                value = ws.cell(row=row_idx, column=col_idx).value
            except Exception:
                continue
            dec = self._as_decimal(value)
            if dec != 0:
                return dec
        return Decimal("0")

    def _extract_col_label_cached(self, ws, col_num: int) -> Optional[str]:
        key = (ws.title, col_num)
        if key in self._col_label_cache:
            return self._col_label_cache[key]
        label = self._extract_col_label(ws, col_num)
        self._col_label_cache[key] = label
        return label

    def _is_zero_placeholder_match(self, matches: List[MatchedAccount]) -> bool:
        return bool(matches) and all(m.compte.startswith("__ZERO__") for m in matches)

    def _match_explicit_account_codes(
        self,
        row_label: str,
        sheet_name: str,
        cell_ref: str,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        expected_classes: Optional[Set[str]],
        is_total: bool,
    ) -> List[MatchedAccount]:
        return []

    def _fuzzy_match_accounts(
        self,
        row_label: str,
        col_type: Optional[str],
        is_total: bool,
        year_hint: Optional[str] = None,
        value_source: str = "final",
        expected_classes: Optional[Set[str]] = None,
        preferred_prefixes: Optional[Set[str]] = None,
        sheet_name: Optional[str] = None,
        cell_ref: Optional[str] = None,
    ) -> List[MatchedAccount]:
        return []

    def _match_with_business_rules(
        self,
        row_label: str,
        sheet_name: str,
        cell_ref: str,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        expected_classes: Optional[Set[str]],
    ) -> List[MatchedAccount]:
        return []

    def _sum_balance_by_prefixes(
        self,
        prefixes: Tuple[str, ...],
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
        expected_classes: Optional[Set[str]] = None,
        drop_hierarchy: bool = False,
    ) -> Decimal:
        balance_accounts, slices_map, use_previous_dataset = self._select_balance_dataset_for_year(year_hint)
        selected: Dict[str, Decimal] = {}
        for compte, normalized_row in balance_accounts.items():
            if expected_classes and normalized_row.classe not in expected_classes:
                continue
            if not any(compte.startswith(prefix) for prefix in prefixes):
                continue
            amount = self._resolve_amount(
                normalized_row,
                col_type,
                year_hint,
                value_source,
                slices_map=slices_map,
                use_previous_year_dataset=use_previous_dataset,
            )
            if amount > 0:
                selected[compte] = amount

        if not selected:
            return Decimal("0")

        if drop_hierarchy:
            comptes = list(selected.keys())
            filtered: Dict[str, Decimal] = {}
            for compte, amount in selected.items():
                is_parent = any(other.startswith(compte) and len(other) > len(compte) for other in comptes)
                if is_parent:
                    continue
                filtered[compte] = amount
            if filtered:
                selected = filtered

        return sum(selected.values(), Decimal("0"))

    def _sum_rows_same_column(self, ws, row_start: int, row_end: int, col_num: int) -> Optional[Decimal]:
        total = Decimal("0")
        used_any = False
        for row_idx in range(row_start, row_end + 1):
            val = self._as_decimal_strict(ws.cell(row=row_idx, column=col_num).value)
            if val is None:
                continue
            used_any = True
            total += val
        if not used_any:
            return None
        return total if total >= 0 else Decimal("0")

    def _compute_note3c_amortization_value(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
    ) -> Optional[Decimal]:
        if self.wb is None:
            return None
        if self._canonical_sheet_name(need.sheet) != "NOTE 3C":
            return None

        col_norm = self._normalize_text(need.col_label or "")
        if "cumul des amortissements" not in col_norm:
            return None

        row_norm = self._normalize_text(need.row_label or "")
        ws = self.wb[need.sheet]

        if "sous total" in row_norm and "immobilisations incorporelles" in row_norm:
            subtotal = self._sum_rows_same_column(ws, 14, 17, need.col_num)
            if subtotal is not None:
                return subtotal
            return self._sum_balance_by_prefixes(
                ("281",),
                col_type,
                year_hint,
                "final",
                expected_classes={"2"},
                drop_hierarchy=True,
            )
        if "sous total" in row_norm and "immobilisations corporelles" in row_norm:
            subtotal = self._sum_rows_same_column(ws, 19, 25, need.col_num)
            if subtotal is not None:
                return subtotal
            return self._sum_balance_by_prefixes(
                ("282", "283", "284", "285", "286", "287", "288", "289", "29"),
                col_type,
                year_hint,
                "final",
                expected_classes={"2"},
                drop_hierarchy=True,
            )
        if "total general" in row_norm:
            v_incorp = self._as_decimal_strict(ws.cell(row=18, column=need.col_num).value)
            v_corp = self._as_decimal_strict(ws.cell(row=26, column=need.col_num).value)
            if v_incorp is not None or v_corp is not None:
                return max(Decimal("0"), (v_incorp or Decimal("0")) + (v_corp or Decimal("0")))
            return self._sum_balance_by_prefixes(
                ("281", "282", "283", "284", "285", "286", "287", "288", "289", "29"),
                col_type,
                year_hint,
                "final",
                expected_classes={"2"},
                drop_hierarchy=True,
            )

        row_prefixes: Optional[Tuple[str, ...]] = None
        if "frais de developpement" in row_norm:
            row_prefixes = ("2811", "2810")
        elif "brevet" in row_norm or "logiciel" in row_norm:
            row_prefixes = ("2813",)
        elif "fonds commercial" in row_norm or "fond commercial" in row_norm:
            row_prefixes = ("2815",)
        elif "autres immobilisations incorporelles" in row_norm:
            row_prefixes = ("2816", "2817", "2818", "2819")
        elif "terrains hors immeubles de placement" in row_norm:
            row_prefixes = ("2821",)
        elif "terrains" in row_norm and "immeubles de placement" in row_norm:
            row_prefixes = ("2822",)
        elif "batiments hors immeubles de placement" in row_norm:
            row_prefixes = ("2831",)
        elif "batiments" in row_norm and "immeubles de placement" in row_norm:
            row_prefixes = ("2832",)
        elif "amenagements" in row_norm or "agencements" in row_norm or "installations" in row_norm:
            row_prefixes = ("2834", "2835", "2838", "2839")
        elif "materiel de transport" in row_norm:
            row_prefixes = ("2845",)
        elif "materiel" in row_norm or "mobilier" in row_norm or "actif biolog" in row_norm:
            row_prefixes = ("284", "286", "287")

        if not row_prefixes:
            return None

        return self._sum_balance_by_prefixes(
            row_prefixes,
            col_type,
            year_hint,
            "final",
            expected_classes={"2"},
            drop_hierarchy=True,
        )

    def _compute_note3_targeted_value(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[Decimal]:
        if self._sheet_group(need.sheet) != "NOTE3":
            return None
        row_norm = self._normalize_text(need.row_label or "")
        col_norm = self._normalize_text(need.col_label or "")

        brut_prefixes: Optional[Tuple[str, ...]] = None
        amort_prefixes: Optional[Tuple[str, ...]] = None

        if "sous total" in row_norm:
            return None
        if "frais de developpement" in row_norm:
            brut_prefixes = ("201",)
            amort_prefixes = ("2811", "2810", "280")
        elif "brevet" in row_norm or "logiciel" in row_norm:
            brut_prefixes = ("202",)
            amort_prefixes = ("2813",)
        elif "fonds commercial" in row_norm or "fond commercial" in row_norm:
            brut_prefixes = ("203",)
            amort_prefixes = ("2815",)
        elif "autres immobilisations incorporelles" in row_norm:
            brut_prefixes = ("204",)
            amort_prefixes = ("2816", "2817", "2818", "2819")
        elif "titres de participation" in row_norm:
            brut_prefixes = ("26",)
            amort_prefixes = ("296",)

        if not brut_prefixes:
            return None

        if "cumul des amortissements" in col_norm:
            return self._sum_balance_by_prefixes(
                amort_prefixes or ("281",),
                col_type,
                year_hint,
                "final",
                expected_classes={"2"},
            )
        return self._sum_balance_by_prefixes(
            brut_prefixes,
            col_type,
            year_hint,
            value_source,
            expected_classes={"2"},
        )

    def _compute_note3_incorp_subtotal(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[Decimal]:
        if self.wb is None:
            return None
        if self._sheet_group(need.sheet) != "NOTE3":
            return None
        row_norm = self._normalize_text(need.row_label or "")
        if "sous total" not in row_norm or "immobilisations incorporelles" not in row_norm:
            return None

        ws = self.wb[need.sheet]
        detail_keywords = (
            "frais de developpement",
            "brevet",
            "logiciel",
            "fonds commercial",
            "fond commercial",
            "autres immobilisations incorporelles",
        )

        total = Decimal("0")
        used_any = False
        for row_idx in range(11, need.row_num):
            candidate_label = self._normalize_text(self._extract_row_label(ws, row_idx) or "")
            if not candidate_label:
                continue
            if "sous total" in candidate_label or "total general" in candidate_label:
                continue
            if not any(keyword in candidate_label for keyword in detail_keywords):
                continue
            parsed = self._as_decimal_strict(ws.cell(row=row_idx, column=need.col_num).value)
            if parsed is None:
                continue
            used_any = True
            total += parsed

        if used_any:
            return total if total >= 0 else Decimal("0")

        col_norm = self._normalize_text(need.col_label or "")
        if "cumul des amortissements" in col_norm:
            return self._sum_balance_by_prefixes(
                ("281",),
                col_type,
                year_hint,
                "final",
                expected_classes={"2"},
            )
        return self._sum_balance_by_prefixes(
            ("20",),
            col_type,
            year_hint,
            value_source,
            expected_classes={"2"},
        )

    def _compute_note3a_targeted_value(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[Decimal]:
        if self.wb is None:
            return None
        if self._canonical_sheet_name(need.sheet) != "NOTE 3A":
            return None

        ws = self.wb[need.sheet]
        row_norm = self._normalize_text(need.row_label or "")

        if not row_norm or "commentaire" in row_norm:
            return None
        if self._looks_like_section_heading(need.row_label or "") and "total" not in row_norm:
            return None

        if row_norm == "immobilisations incorporelles":
            subtotal = self._sum_rows_same_column(ws, 15, 18, need.col_num)
            if subtotal is not None:
                return subtotal
        if row_norm == "immobilisation corporelles":
            subtotal = self._sum_rows_same_column(ws, 20, 27, need.col_num)
            if subtotal is not None:
                return subtotal
        if "total general" in row_norm:
            v1 = self._as_decimal_strict(ws.cell(row=28, column=need.col_num).value)
            v2 = self._as_decimal_strict(ws.cell(row=29, column=need.col_num).value)
            v3 = self._as_decimal_strict(ws.cell(row=31, column=need.col_num).value)
            v4 = self._as_decimal_strict(ws.cell(row=32, column=need.col_num).value)
            if any(v is not None for v in (v1, v2, v3, v4)):
                return max(Decimal("0"), (v1 or Decimal("0")) + (v2 or Decimal("0")) + (v3 or Decimal("0")) + (v4 or Decimal("0")))

        row_prefixes: Optional[Tuple[str, ...]] = None
        if "frais de developpement" in row_norm:
            row_prefixes = ("201",)
        elif "brevets" in row_norm or "logiciels" in row_norm or "brevet" in row_norm or "logiciel" in row_norm:
            row_prefixes = ("213",)
        elif "fonds commercial" in row_norm or "fond commercial" in row_norm:
            row_prefixes = ("215",)
        elif "autres immobilisations incorporelles" in row_norm:
            row_prefixes = ("219",)
        elif "terrains hors immeubles de placement" in row_norm:
            row_prefixes = ("223", "221", "222")
        elif "terrains" in row_norm and "immeubles de placement" in row_norm:
            row_prefixes = ("224", "225")
        elif "batiments hors immeubles de placement" in row_norm:
            row_prefixes = ("231", "232")
        elif "batiments" in row_norm and "immeubles de placement" in row_norm:
            row_prefixes = ("233",)
        elif "amenagements" in row_norm or "agencements" in row_norm or "installations" in row_norm:
            row_prefixes = ("234", "235", "238", "239")
        elif "materiel de transport" in row_norm:
            row_prefixes = ("245",)
        elif "materiel" in row_norm or "mobilier" in row_norm or "actif biolog" in row_norm:
            row_prefixes = ("241", "243", "244", "248", "249")
        elif "avances et acomptes versees sur immobilisations" in row_norm:
            row_prefixes = ("252", "25")
        elif "titres de participation" in row_norm:
            row_prefixes = ("26",)
        elif "autres immobilisations financieres" in row_norm:
            row_prefixes = ("27", "275", "278")

        if not row_prefixes:
            return None

        return self._sum_balance_by_prefixes(
            row_prefixes,
            col_type,
            year_hint,
            value_source,
            expected_classes={"2"},
            drop_hierarchy=True,
        )

    def _compute_cf_targeted_value(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[Decimal]:
        sheet_group = self._sheet_group(need.sheet)
        if sheet_group not in {"CF1", "CF2", "C2_NOTE25"}:
            return None
        if self.wb is None:
            return Decimal("0")

        ws = self.wb[need.sheet]
        row_norm = self._normalize_text(need.row_label or "")
        col_norm = self._normalize_text(need.col_label or "")
        upper_sheet = self._canonical_sheet_name(need.sheet)

        if upper_sheet == "CF1 QUATER" and "totaux" in row_norm:
            subtotal = self._sum_rows_same_column(ws, 11, 22, need.col_num)
            return subtotal if subtotal is not None else Decimal("0")
        if upper_sheet == "CF2 BIS" and "total lignes" in row_norm:
            subtotal = self._sum_rows_same_column(ws, 12, 23, need.col_num)
            return subtotal if subtotal is not None else Decimal("0")
        if upper_sheet == "C2 NOTE 25":
            if "sous total a" in row_norm:
                subtotal = self._sum_rows_same_column(ws, 11, 16, need.col_num)
                return subtotal if subtotal is not None else Decimal("0")
            if "sous total b" in row_norm:
                subtotal = self._sum_rows_same_column(ws, 21, 38, need.col_num)
                return subtotal if subtotal is not None else Decimal("0")
            if row_norm == "total a b" or row_norm == "total a b " or row_norm.startswith("total a b"):
                v1 = self._as_decimal_strict(ws.cell(row=17, column=need.col_num).value)
                v2 = self._as_decimal_strict(ws.cell(row=39, column=need.col_num).value)
                if v1 is not None or v2 is not None:
                    return (v1 or Decimal("0")) + (v2 or Decimal("0"))
                return Decimal("0")

        if upper_sheet == "CF2 TER":
            if "tva nette a payer" in row_norm:
                v5 = self._as_decimal_strict(ws.cell(row=14, column=need.col_num).value) or Decimal("0")
                v6 = self._as_decimal_strict(ws.cell(row=15, column=need.col_num).value) or Decimal("0")
                v7 = self._as_decimal_strict(ws.cell(row=16, column=need.col_num).value) or Decimal("0")
                return max(Decimal("0"), v5 - v6 - v7)
            if "credit de tva net a reporter" in row_norm:
                v5 = self._as_decimal_strict(ws.cell(row=14, column=need.col_num).value) or Decimal("0")
                v6 = self._as_decimal_strict(ws.cell(row=15, column=need.col_num).value) or Decimal("0")
                v7 = self._as_decimal_strict(ws.cell(row=16, column=need.col_num).value) or Decimal("0")
                return max(Decimal("0"), v6 + v7 - v5)

        prefixes: Optional[Tuple[str, ...]] = None
        source_override = value_source
        if "impot sur le resultat" in row_norm or "impots sur les benefices" in row_norm:
            prefixes = ("89", "69", "44")
            source_override = "final"
        elif "minimum de perception" in row_norm:
            prefixes = ("89", "44")
            source_override = "final"
        elif "ircm" in row_norm:
            prefixes = ("44", "47", "75")
            source_override = "final"
        elif "retenue" in row_norm or "acompte" in row_norm or "acompte" in row_norm:
            prefixes = ("44",)
            source_override = "final"
        elif "tva brute" in row_norm:
            prefixes = ("44",)
            source_override = "movement_credit"
        elif "tva deductible" in row_norm:
            prefixes = ("44",)
            source_override = "movement_debit"
        elif "tva nette" in row_norm or "credit de tva" in row_norm:
            prefixes = ("44",)
            source_override = "final"

        if prefixes:
            return self._sum_balance_by_prefixes(
                prefixes,
                col_type,
                year_hint,
                source_override,
                expected_classes={"1", "4", "5", "6", "7", "8"},
                drop_hierarchy=True,
            )

        if "montant" in col_norm or "base" in col_norm or "impot" in col_norm or "tva" in col_norm or "droits" in col_norm:
            return Decimal("0")
        return Decimal("0")

    def _compute_bilan_targeted_value(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[Decimal]:
        if self._canonical_sheet_name(need.sheet) != "BILAN PAYSAGE":
            return None
        if self.wb is None:
            return None

        col_letter_match = re.match(r"[A-Z]+", need.cell)
        col_letter = col_letter_match.group(0) if col_letter_match else ""
        if col_letter not in {"D", "G", "K", "L"}:
            return None

        row_norm = self._normalize_text(need.row_label or "")
        if not row_norm:
            return None
        ws = self.wb[need.sheet]

        def sum_same_col(rows: Tuple[int, ...]) -> Optional[Decimal]:
            total = Decimal("0")
            used = False
            for r in rows:
                parsed = self._as_decimal_strict(ws.cell(row=r, column=need.col_num).value)
                if parsed is None:
                    continue
                used = True
                total += parsed
            return total if used else None

        if col_letter in {"D", "G"}:
            if "frais de developpement" in row_norm:
                return self._sum_balance_by_prefixes(("201",), col_type, year_hint, value_source, {"2"}, True)
            if "brevet" in row_norm or "logiciel" in row_norm:
                return self._sum_balance_by_prefixes(("213", "202"), col_type, year_hint, value_source, {"2"}, True)
            if "fond commercial" in row_norm or "fonds commercial" in row_norm:
                return self._sum_balance_by_prefixes(("215", "203"), col_type, year_hint, value_source, {"2"}, True)
            if "autres immobilisations incorporelles" in row_norm or "autres immobilisations incorporelles" in row_norm:
                return self._sum_balance_by_prefixes(("219", "204"), col_type, year_hint, value_source, {"2"}, True)
            if row_norm.startswith("terrains"):
                return self._sum_balance_by_prefixes(("22",), col_type, year_hint, value_source, {"2"}, True)
            if row_norm.startswith("batiments"):
                return self._sum_balance_by_prefixes(("231", "232", "233"), col_type, year_hint, value_source, {"2"}, True)
            if "immobilisation financieres" in row_norm and "autres" not in row_norm and "titres" not in row_norm:
                subtotal = sum_same_col((25, 26))
                if subtotal is not None:
                    return subtotal
                return self._sum_balance_by_prefixes(("26", "27"), col_type, year_hint, value_source, {"2"}, True)
            if "total actif immobilise" in row_norm:
                subtotal = sum_same_col((12, 17, 24))
                if subtotal is not None:
                    return subtotal
            if "actif circulant hao" in row_norm:
                return self._sum_balance_by_prefixes(("38",), col_type, year_hint, value_source, {"3"}, True)
            if "stocks et encours" in row_norm:
                return self._sum_balance_by_prefixes(("3",), col_type, year_hint, value_source, {"3"}, True)
            if "fournisseurs avances versees" in row_norm:
                return self._sum_balance_by_prefixes(("409",), col_type, year_hint, value_source, {"4"}, True)
            if "creances et emplois assimiles" in row_norm:
                subtotal = sum_same_col((31, 32, 33))
                if subtotal is not None:
                    return subtotal
                return self._sum_balance_by_prefixes(("41", "42", "43", "44", "45", "46", "47", "48"), col_type, year_hint, value_source, {"4"}, True)
            if "autres creances" in row_norm:
                return self._sum_balance_by_prefixes(("42", "43", "44", "45", "46", "47", "48"), col_type, year_hint, value_source, {"4"}, True)
            if "total actif circulant" in row_norm:
                subtotal = sum_same_col((28, 29, 30))
                if subtotal is not None:
                    return subtotal
            if "titres de placement" in row_norm:
                return self._sum_balance_by_prefixes(("50",), col_type, year_hint, value_source, {"5"}, True)
            if "valeurs a encaisser" in row_norm:
                return self._sum_balance_by_prefixes(("51",), col_type, year_hint, value_source, {"5"}, True)
            if "banques cheques postaux caisse" in row_norm:
                return self._sum_balance_by_prefixes(("52", "53", "54", "56", "57", "58"), col_type, year_hint, value_source, {"5"}, True)
            if "total tresorerie actif" in row_norm:
                subtotal = sum_same_col((35, 36, 37))
                if subtotal is not None:
                    return subtotal
            if row_norm == "total general":
                subtotal = sum_same_col((27, 34, 38, 39))
                if subtotal is not None:
                    return subtotal
            return None

        if "apporteurs capital non appele" in row_norm:
            val = self._sum_balance_by_prefixes(("1013",), col_type, year_hint, value_source, {"1"}, True)
            return -val if val else Decimal("0")
        if "primes liees au capital social" in row_norm:
            return self._sum_balance_by_prefixes(("11",), col_type, year_hint, value_source, {"1"}, True)
        if "ecarts de reevaluations" in row_norm:
            return self._sum_balance_by_prefixes(("12",), col_type, year_hint, value_source, {"1"}, True)
        if "reserves indisponibles" in row_norm:
            return self._sum_balance_by_prefixes(("112", "113"), col_type, year_hint, value_source, {"1"}, True)
        if "reserves libres" in row_norm:
            return self._sum_balance_by_prefixes(("114", "118"), col_type, year_hint, value_source, {"1"}, True)
        if "resultat net de l exercice" in row_norm:
            cr_sheet = "COMPTE DE RESULTAT"
            if cr_sheet in self.wb.sheetnames:
                ws_cr = self.wb[cr_sheet]
                col_cr = self._resolve_flux_value_column(ws_cr, year_hint)
                if col_cr is not None:
                    code_rows = self._sheet_code_row_map(cr_sheet)
                    row_xi = code_rows.get("XI")
                    if row_xi:
                        parsed = self._as_decimal_strict(ws_cr.cell(row=row_xi, column=col_cr).value)
                        if parsed is not None:
                            return parsed
            return Decimal("0")
        if "subventions d investissement" in row_norm:
            return self._sum_balance_by_prefixes(("15",), col_type, year_hint, value_source, {"1"}, True)
        if "povisions reglementees" in row_norm or "provisions reglementees" in row_norm:
            return self._sum_balance_by_prefixes(("14",), col_type, year_hint, value_source, {"1"}, True)
        if "total capitaux propres et ressources assimilees" in row_norm:
            subtotal = sum_same_col((12, 13, 14, 15, 16, 17, 18, 19, 20, 21))
            if subtotal is not None:
                return subtotal
            return None
        if "emprunts et dettes financieres diverses" in row_norm:
            return self._sum_balance_by_prefixes(("16", "17"), col_type, year_hint, value_source, {"1"}, True)
        if "dettes de location acquisition" in row_norm:
            return self._sum_balance_by_prefixes(("17",), col_type, year_hint, value_source, {"1"}, True)
        if "total dettes financieres et ressources assimilees" in row_norm:
            subtotal = sum_same_col((23, 24, 25))
            if subtotal is not None:
                return subtotal
            return None
        if "total ressources stables" in row_norm:
            subtotal = sum_same_col((22, 26))
            if subtotal is not None:
                return subtotal
            return None
        if "dettes circulantes hao" in row_norm:
            return self._sum_balance_by_prefixes(("47",), col_type, year_hint, value_source, {"4"}, True)
        if "fournisseurs d exploitation" in row_norm:
            return self._sum_balance_by_prefixes(("40",), col_type, year_hint, value_source, {"4"}, True)
        if "dettes fiscales et sociales" in row_norm:
            return self._sum_balance_by_prefixes(("42", "43", "44"), col_type, year_hint, value_source, {"4"}, True)
        if "autres dettes" in row_norm:
            return self._sum_balance_by_prefixes(("45", "46", "47", "48", "49"), col_type, year_hint, value_source, {"4"}, True)
        if "provisions pour risques a court terme" in row_norm:
            return self._sum_balance_by_prefixes(("15",), col_type, year_hint, value_source, {"1"}, True)
        if "total pasif circulant" in row_norm or "total passif circulant" in row_norm:
            subtotal = sum_same_col((28, 29, 30, 31, 32, 33))
            if subtotal is not None:
                return subtotal
            return None
        if "banques credits d escompte" in row_norm:
            return self._sum_balance_by_prefixes(("52", "56"), col_type, year_hint, value_source, {"5"}, True)
        if "banques etablissements financiers et credits de tresorerie" in row_norm:
            return self._sum_balance_by_prefixes(("52", "56"), col_type, year_hint, value_source, {"5"}, True)
        if "total tresorerie pasif" in row_norm or "total tresorerie passif" in row_norm:
            subtotal = sum_same_col((35, 36, 37))
            if subtotal is not None:
                return subtotal
            return None
        if row_norm == "total general":
            subtotal = sum_same_col((27, 34, 38, 39))
            if subtotal is not None:
                return subtotal
            return None

        return None

    def _compute_cr_targeted_value(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[Decimal]:
        if self._sheet_group(need.sheet) != "CR":
            return None
        row_norm = self._normalize_text(need.row_label or "")
        if not row_norm:
            return None

        prefixes: Optional[Tuple[str, ...]] = None
        if "variation de stock de marchandises" in row_norm:
            prefixes = ("603",)
        elif "vente de produits fabriques" in row_norm:
            prefixes = ("71",)
        elif row_norm == "transports":
            prefixes = ("62",)
        elif "impots et taxes" in row_norm:
            prefixes = ("63",)
        elif "charges de personnel" in row_norm:
            prefixes = ("64",)
        elif "revenus financiers et assimiles" in row_norm:
            prefixes = ("75", "76", "77")
        elif "transfert de charges financieres" in row_norm:
            prefixes = ("78",)

        if not prefixes:
            return None

        return self._sum_balance_by_prefixes(
            prefixes,
            col_type,
            year_hint,
            value_source,
            expected_classes={"6", "7"},
            drop_hierarchy=True,
        )

    def _compute_cr_formula_value(self, need: CellNeed) -> Optional[Decimal]:
        if self.wb is None:
            return None
        if self._sheet_group(need.sheet) != "CR":
            return None

        ws = self.wb[need.sheet]
        code_raw = ws.cell(row=need.row_num, column=1).value
        code = str(code_raw or "").strip().upper()
        if not re.fullmatch(r"[A-Z]{1,3}", code):
            return None

        formulas: Dict[str, Tuple[Tuple[str, ...], Tuple[str, ...]]] = {
            "XA": (("TA", "RB"), ("RA",)),
            "XB": (("TA", "TB", "TC", "TD"), ()),
            "XC": (("XA", "TE", "TF", "TG", "TH", "TI"), ("RC", "RD", "RE", "RF", "RG", "RH", "RI", "RJ")),
            "XD": (("XC",), ("RK",)),
            "XE": (("XD", "TJ"), ("RL",)),
            "XF": (("TK", "TL", "TM"), ("RM", "RN")),
            "XG": (("XE", "XF"), ()),
            "XH": (("TN", "TO"), ("RO", "RP")),
            "XI": (("XG", "XH"), ("RQ", "RS")),
        }
        if code not in formulas:
            return None

        code_rows = self._sheet_code_row_map(need.sheet)
        plus_codes, minus_codes = formulas[code]

        used_any = False
        total = Decimal("0")

        for ref_code in plus_codes:
            ref_row = code_rows.get(ref_code)
            if not ref_row:
                continue
            val = self._as_decimal_strict(ws.cell(row=ref_row, column=need.col_num).value)
            if val is None:
                continue
            used_any = True
            total += val

        for ref_code in minus_codes:
            ref_row = code_rows.get(ref_code)
            if not ref_row:
                continue
            val = self._as_decimal_strict(ws.cell(row=ref_row, column=need.col_num).value)
            if val is None:
                continue
            used_any = True
            total -= val

        if not used_any:
            return None
        if total < 0:
            return Decimal("0")
        return total

    def _sheet_code_row_map(self, sheet_name: str) -> Dict[str, int]:
        if sheet_name in self._sheet_code_row_cache:
            return self._sheet_code_row_cache[sheet_name]
        ws = self.wb[sheet_name]
        mapping: Dict[str, int] = {}
        for row_idx in range(1, ws.max_row + 1):
            raw = ws.cell(row=row_idx, column=1).value
            code = str(raw or "").strip().upper()
            if re.fullmatch(r"[A-Z]{1,3}", code):
                mapping[code] = row_idx
        self._sheet_code_row_cache[sheet_name] = mapping
        return mapping

    def _compute_note34_ebe_from_cr(self, need: CellNeed, year_hint: Optional[str]) -> Optional[Decimal]:
        if self.wb is None:
            return None
        if self._sheet_group(need.sheet) != "NOTE34":
            return None
        if self._normalize_text(need.row_label or "") != "ebe":
            return None
        cr_sheet = "COMPTE DE RESULTAT"
        if cr_sheet not in self.wb.sheetnames:
            return None
        ws_cr = self.wb[cr_sheet]
        cr_col = self._resolve_flux_value_column(ws_cr, year_hint)
        if cr_col is None:
            return None
        cr_rows = self._sheet_code_row_map(cr_sheet)
        xd_row = cr_rows.get("XD")
        if not xd_row:
            return None
        parsed = self._as_decimal_strict(ws_cr.cell(row=xd_row, column=cr_col).value)
        if parsed is not None:
            return parsed
        return Decimal("0")

    def _compute_note4_total_net_depreciation(self, need: CellNeed) -> Optional[Decimal]:
        if self.wb is None:
            return None
        if self._sheet_group(need.sheet) != "NOTE4":
            return None
        row_norm = self._normalize_text(need.row_label or "")
        if "total net de depreciation" not in row_norm:
            return None

        ws = self.wb[need.sheet]
        brut = None
        dep_part = None
        dep_other = None

        for row_idx in range(11, ws.max_row + 1):
            current_label = self._normalize_text(self._extract_row_label(ws, row_idx) or "")
            if not current_label:
                continue
            value = self._as_decimal(ws.cell(row=row_idx, column=need.col_num).value)
            if value < 0:
                continue
            if "total brut" in current_label:
                brut = value
            elif "depreciations titres de participation" in current_label:
                dep_part = value
            elif "depreciations autres immobilisations" in current_label:
                dep_other = value

        if brut is None:
            return None
        dep_total = (dep_part or Decimal("0")) + (dep_other or Decimal("0"))
        net = brut - dep_total
        if net < 0:
            return Decimal("0")
        return net

    def _compute_note34_flux_value(self, need: CellNeed, year_hint: Optional[str]) -> Optional[Decimal]:
        if self.wb is None:
            return None
        if self._sheet_group(need.sheet) != "NOTE34":
            return None
        row_norm = self._normalize_text(need.row_label or "")
        if "flux de tresorerie des activites operationnelles" in row_norm:
            flux_kind = "operations"
        elif "flux de tresorerie des activites d investissement" in row_norm:
            flux_kind = "investment"
        elif "flux de tresorerie des activites de financement" in row_norm:
            flux_kind = "financing"
        else:
            return None

        flux_sheet_name = "TABLEAU DES FLUX DE TRESORERIE"
        if flux_sheet_name not in self.wb.sheetnames:
            return None
        ws = self.wb[flux_sheet_name]

        target_col = self._resolve_flux_value_column(ws, year_hint)
        if target_col is None:
            return Decimal("0")
        source_row = self._find_flux_row_index(ws, flux_kind)
        if source_row is None:
            return Decimal("0")

        raw_value = ws.cell(row=source_row, column=target_col).value
        parsed = self._as_decimal_strict(raw_value)
        if parsed is not None:
            return parsed

        return Decimal("0")

    def _resolve_flux_value_column(self, ws, year_hint: Optional[str]) -> Optional[int]:
        wanted = "n1" if year_hint == "n1" else "n"
        for col_num in range(1, ws.max_column + 1):
            col_label = self._extract_col_label_cached(ws, col_num)
            if not col_label:
                continue
            hint = self._infer_year_hint(col_label)
            if hint == wanted:
                return col_num
        return None

    def _find_flux_row_index(self, ws, flux_kind: str) -> Optional[int]:
        best_row = None
        for row_idx in range(11, ws.max_row + 1):
            label = self._normalize_text(self._extract_row_label(ws, row_idx) or "")
            if not label:
                continue
            if flux_kind == "operations":
                if "somme fa a fe" in label and "activites operaionnelles" in label:
                    return row_idx
                if "flux de tresorerie provenant des activites operaionnelles" in label:
                    best_row = row_idx
            elif flux_kind == "investment":
                if "somme ff a fj" in label and "activites d investissements" in label:
                    return row_idx
                if "flux de tresorerie provenant des activites d investissements" in label:
                    best_row = row_idx
            elif flux_kind == "financing":
                if "flux de tresorerie provenant des activites de financement" in label:
                    return row_idx
        return best_row

    def _build_derived_assignment(
        self,
        need: CellNeed,
        col_type: Optional[str],
        year_hint: Optional[str],
        value_source: str,
    ) -> Optional[CellAssignment]:
        note3c_value = self._compute_note3c_amortization_value(need, col_type, year_hint)
        if note3c_value is not None:
            return self._create_derived_assignment(need, note3c_value, col_type, year_hint, "note3c_amortization")

        note3_targeted = self._compute_note3_targeted_value(need, col_type, year_hint, value_source)
        if note3_targeted is not None:
            return self._create_derived_assignment(need, note3_targeted, col_type, year_hint, "note3_targeted")

        note3_subtotal = self._compute_note3_incorp_subtotal(need, col_type, year_hint, value_source)
        if note3_subtotal is not None:
            return self._create_derived_assignment(need, note3_subtotal, col_type, year_hint, "note3_incorp_subtotal")

        cf_value = self._compute_cf_targeted_value(need, col_type, year_hint, value_source)
        if cf_value is not None:
            return self._create_derived_assignment(need, cf_value, col_type, year_hint, "cf_targeted")

        derived_total = self._compute_cr_formula_value(need)
        if derived_total is not None:
            return self._create_derived_assignment(need, derived_total, col_type, year_hint, "cr_formula")

        note34_ebe = self._compute_note34_ebe_from_cr(need, year_hint)
        if note34_ebe is not None:
            return self._create_derived_assignment(need, note34_ebe, col_type, year_hint, "note34_ebe")

        derived_total = self._compute_note4_total_net_depreciation(need)
        if derived_total is not None:
            return self._create_derived_assignment(need, derived_total, col_type, year_hint, "note4_net_depreciation")

        note34_flux = self._compute_note34_flux_value(need, year_hint)
        if note34_flux is not None:
            return self._create_derived_assignment(need, note34_flux, col_type, year_hint, "note34_flux")

        return None

    def _write_cell(self, ws, need: CellNeed, assignment: CellAssignment) -> bool:
        try:
            actual_cell = ws[need.cell]

            for merged_range in ws.merged_cells.ranges:
                if need.cell in merged_range:
                    actual_cell = ws[merged_range.start_cell.coordinate]
                    break

            actual_cell.value = float(assignment.total_amount)
            return True
            
        except Exception as e:
            logger.debug(f"Failed to write {need.cell}: {e}")
            return False

    def _apply_column_formulas(self) -> None:
        strict_no_n1 = not bool(self.previous_balance_accounts) and not self.allow_n1_fallback_without_prev
        for sheet_name, detector in self.formula_detectors.items():
            ws = self.wb[sheet_name]
            formula_columns = detector.get_formula_columns()
            formula_seeds = self._collect_formula_seeds(ws)
            year_hints = self._sheet_year_hints_by_column(ws)

            if not formula_columns and not formula_seeds:
                continue

            if formula_columns:
                logger.info(f"  Sheet {sheet_name}: applying formulas to columns {', '.join(formula_columns)}")

            data_start_row = detector.header_row_detected + 1 if detector.header_row_detected else 11

            for row_idx in range(data_start_row, ws.max_row + 1):
                first_cell = ws.cell(row=row_idx, column=1)
                if not first_cell.value:
                    continue

                if isinstance(first_cell.value, str):
                    text_lower = first_cell.value.lower()
                    if any(kw in text_lower for kw in ['total', 'sous-total', 'en-tête', 'header', 'libellé']):
                        continue

                for col_letter in formula_columns:
                    if strict_no_n1 and year_hints.get(col_letter) == "n1":
                        continue
                    cell = ws[f"{col_letter}{row_idx}"]
                    if cell.data_type == "f":
                        continue
                    if cell.value is None or isinstance(cell.value, (int, float)):
                        formula = detector.apply_formula_to_cell(col_letter, row_idx)
                        if formula:
                            if strict_no_n1 and self._formula_uses_n1_column(str(formula), year_hints):
                                continue
                            try:
                                cell.value = formula
                                self.formulas_applied += 1
                                self.formula_logs.append(
                                    {
                                        "sheet": sheet_name,
                                        "cell": cell.coordinate,
                                        "formula": str(formula),
                                        "source": "header",
                                    }
                                )
                                logger.debug(f"    {cell.coordinate}: {formula}")
                            except Exception as e:
                                logger.warning(f"Failed to apply formula to {cell.coordinate}: {e}")

                for col_letter, (seed_cell, seed_formula) in formula_seeds.items():
                    if strict_no_n1 and year_hints.get(col_letter) == "n1":
                        continue
                    target_cell = ws[f"{col_letter}{row_idx}"]
                    if target_cell.data_type == "f":
                        continue
                    if target_cell.value is None or isinstance(target_cell.value, (int, float)):
                        try:
                            translated = Translator(seed_formula, origin=seed_cell).translate_formula(target_cell.coordinate)
                            if strict_no_n1 and self._formula_uses_n1_column(str(translated), year_hints):
                                continue
                            target_cell.value = translated
                            self.formulas_applied += 1
                            self.formula_logs.append(
                                {
                                    "sheet": sheet_name,
                                    "cell": target_cell.coordinate,
                                    "formula": str(translated),
                                    "source": f"seed:{seed_cell}",
                                }
                            )
                            logger.debug(f"    {target_cell.coordinate}: {translated}")
                        except Exception as e:
                            logger.warning(f"Failed to translate formula to {target_cell.coordinate}: {e}")

    def _sheet_year_hints_by_column(self, ws) -> Dict[str, str]:
        hints: Dict[str, str] = {}
        for col_num in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_num)
            col_label = self._extract_col_label_cached(ws, col_num)
            hint = self._infer_year_hint(col_label)
            if hint:
                hints[col_letter] = hint
        return hints

    def _formula_uses_n1_column(self, formula: str, year_hints: Dict[str, str]) -> bool:
        if not formula or not isinstance(formula, str):
            return False
        refs = re.findall(r"\$?([A-Z]{1,3})\$?\d+", formula.upper())
        for col in refs:
            if year_hints.get(col) == "n1":
                return True
        return False

    def _collect_formula_seeds(self, ws) -> Dict[str, tuple[str, str]]:
        seeds: Dict[str, tuple[str, str]] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f" and isinstance(cell.value, str):
                    col_letter = cell.column_letter
                    if col_letter not in seeds:
                        seeds[col_letter] = (cell.coordinate, cell.value)
        return seeds

    def write_formula_log(self, path: Path | str) -> Path:
        output = Path(path)
        output.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "total": len(self.formula_logs),
            "items": self.formula_logs,
        }
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return output

    def save(self, output_path: Path | str) -> Path:
        if self.wb is None:
            raise RuntimeError("Workbook not loaded")
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        try:
            self.wb.save(output)
            logger.info(f"Saved to {output}")
            return output
        except PermissionError as exc:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            fallback = output.with_name(f"{output.stem}_{timestamp}{output.suffix}")
            logger.warning(
                "Impossible d'ecrire %s (%s). Sauvegarde de secours vers %s.",
                output,
                exc,
                fallback,
            )
            self.wb.save(fallback)
            logger.info(f"Saved to {fallback}")
            return fallback

    def get_stats(self) -> Dict[str, object]:
        total_assigned = sum(a.total_amount for a in self.assignments)
        return {
            "assignments": len(self.assignments),
            "unmatched_cells": len(self.unmatched_cells),
            "total_amount": float(total_assigned),
            "average_confidence": (
                sum(a.confidence for a in self.assignments) / len(self.assignments)
                if self.assignments
                else 0.0
            ),
            "high_confidence": len([a for a in self.assignments if a.confidence > 0.8]),
        }

    def _as_decimal(self, value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        if isinstance(value, (int, float)):
            return Decimal(str(value))
        try:
            txt = str(value).strip().replace(" ", "")
            txt = txt.replace(",", ".") if txt.count(",") == 1 and "." not in txt else txt
            return Decimal(txt or "0")
        except (InvalidOperation, ValueError, TypeError):
            return Decimal("0")

    def _as_decimal_strict(self, value: object) -> Optional[Decimal]:
        if value is None:
            return None
        if isinstance(value, Decimal):
            return value
        if isinstance(value, (int, float)):
            return Decimal(str(value))
        txt = str(value).strip().replace(" ", "")
        if not txt:
            return None
        txt = txt.replace(",", ".") if txt.count(",") == 1 and "." not in txt else txt
        if not re.fullmatch(r"[-+]?\d+(\.\d+)?", txt):
            return None
        try:
            return Decimal(txt)
        except (InvalidOperation, ValueError, TypeError):
            return None


__all__ = ["SemanticBalanceFiller", "CellAssignment", "CellNeed"]