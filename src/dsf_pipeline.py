# -*- coding: utf-8 -*-
"""DSF production pipeline chaining normalization, rules, controls, and reporting."""

from __future__ import annotations

import gc
import logging
import os
import time
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Callable, Dict, Optional, Set

from balance_normalizer import BalanceNormalizer
from dsf_controls import ControlReport, ControlSuite
from dsf_general_info import DSF_InfosGenerales, get_gulfcam_config
from dsf_general_prefill import DSFGeneralPrefiller
from smart_general_filler import SmartGeneralFiller
from dsf_inventory import DSFInventory
from dsf_reporting import generate_reports
from dsf_rule_config import DSFRuleSet, load_rule_set
from dsf_rule_engine import RuleEngine, RuleEngineResult
from dsf_secure_writer import ProtectedDSFWriter
from semantic_balance_filler import CellAssignment, SemanticBalanceFiller
from semantic_validators import SemanticFillerValidator
from openpyxl.utils import coordinate_to_tuple, range_boundaries
from dsf_calculation_applier import DSFCalculationApplier, apply_calculations_to_dsf
# P2 — Nouveaux modules
from dsf_fiscal_controls import run_fiscal_controls, FiscalReport
from dsf_notes_filler import fill_notes_from_balance
from dsf_cash_flow_filler import fill_cash_flow
from dsf_bilan_composite_filler import fill_bilan_composite_cells  # P3-B1 : Remplissage composite BILAN
# P3 — Lot D : Détection anomalies et validation pré-soumission
from dsf_anomaly_detector import detect_abnormal_balances
from dsf_presubmission_validator import run_presubmission_validation

logger = logging.getLogger(__name__)


# Cellules DSF critiques où les affectations issues du moteur de règles
# doivent être prioritaires sur tout remplissage sémantique.
PROTECTED_CELLS = {
    # Bilan — Totaux
    ("BILAN PAYSAGE", "D99"),  # Total Actif N
    ("BILAN PAYSAGE", "F99"),  # Total Actif N-1
    ("BILAN PAYSAGE", "K99"),  # Total Passif N
    ("BILAN PAYSAGE", "M99"),  # Total Passif N-1
    # Compte de Résultat — Résultat Net (coordonnées à adapter si besoin)
    ("COMPTE DE RESULTAT", "D99"),  # Résultat net N
    ("COMPTE DE RESULTAT", "F99"),  # Résultat net N-1
}

# Sections / notes considérées comme sensibles et donc protégées
PROTECTED_SECTIONS = {
    "BILAN_ACTIF",
    "BILAN_PASSIF",
    "CR_RESULTAT",
    "NOTE_4",
    "NOTE_7",
    "NOTE_10",
    "NOTE_11",
    "NOTE_15",
    "NOTE_16",
    "NOTE_17",
}


@dataclass
class PipelineArtifacts:
    dsf_output: Path
    prefilled_template: Path
    report_json: Optional[Path]
    report_html: Optional[Path]
    controls_summary: Dict[str, int]
    pipeline_warnings: list = field(default_factory=list)  # P3-m1 : warnings non silencieux


@dataclass
class DSFPipelineConfig:
    template_dsf: Path = Path("templates/DSF Normal standard.xlsx")
    balance_input: Path = Path("input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx")
    previous_balance_input: Optional[Path] = None
    inventory_path: Path = Path("data/dsf_inventory.json")
    rules_path: Path = Path("config/dsf_rule_generated.yaml")
    prefill_mapping_path: Path = Path("config/dsf_prefill_mapping.json")
    prefilled_template_path: Path = Path("output/dsf_prefilled_template.xlsx")
    dsf_output: Path = Path("output/DSF_OUTPUT_2024.xlsx")
    report_json_path: Path = Path("output/reports/dsf_report.json")
    report_html_path: Path = Path("output/reports/dsf_report.html")
    exercice_year: int = 2024
    chunk_size: int = 2000
    writer_chunk_size: int = 1000
    control_tolerance: Decimal = Decimal("1")
    filling_method: str = "hybrid"  # "semantic", "hybrid" or "rule-based"
    fuzzy_threshold: float = 0.7  # For semantic filling (plus strict = moins de matching coûteux)
    apply_calculations: bool = True  # Apply TOTAL, VARIATION, RATIO formulas
    use_calculation_formulas: bool = True  # Use Excel formulas vs numeric values
    use_smart_general_filler: bool = True  # Prefer SmartGeneralFiller for ENTETE/R1/R2/R3/NOTE13/PAGE DE GARDE
    apply_note_rules_in_semantic: bool = True  # Apply intelligent rules for all notes in semantic mode
    general_info: DSF_InfosGenerales = field(default_factory=get_gulfcam_config)
    balance_column_overrides: Optional[Dict[str, object]] = field(
        default_factory=lambda: {
            "compte": 1,
            "label": 4,
            "debit_columns": [24, 21, 17, 11],
            "credit_columns": [27, 26, 20, 14, 13],
            "opening_debit_columns": [11],
            "opening_credit_columns": [14, 13],
            "movement_debit_columns": [17],
            "movement_credit_columns": [20],
            "closing_debit_columns": [24, 21],
            "closing_credit_columns": [27, 26],
        }
    )
    # Overrides pour la balance N-1 séparée (si previous_balance_input est fourni).
    # Par défaut None = utilise balance_column_overrides.
    previous_balance_column_overrides: Optional[Dict[str, object]] = None
    # Si True et qu'aucune balance N-1 séparée n'est fournie, extrait les soldes N-1
    # depuis les colonnes d'ouverture (opening_debit/credit_columns) de la balance N.
    use_opening_columns_as_n1: bool = True
    # P2 — Contrôles fiscaux (A1)
    enable_fiscal_controls: bool = True
    # P2 — Remplissage Notes Annexes 1-12 (A3) — désactivé par défaut pour plus de vitesse
    enable_notes_filler: bool = False
    # P2 — Tableau des Flux de Trésorerie (A4) — désactivé par défaut pour plus de vitesse
    enable_cash_flow: bool = False
    # Extraction et application des formules Excel depuis un DSF référence
    formula_reference_dsf: Optional[Path] = Path("input/DSF GULFCAM 2023 V3.xlsx")  # Si None, utilise template_dsf
    enable_formula_extraction: bool = False
    
    # NOTE 20 OPTIMIZATIONS (solution #2, #3, #5, #6)
    # Paramètres de performance spécifiques à NOTE 20
    note20_batch_size: int = 20  # Plus petit batch = moins de mémoire
    note20_max_workers: int = 2  # Limiter parallélisme pour éviter contention
    note20_fuzzy_threshold: float = 0.85  # Plus strict = moins de matching coûteux
    note20_min_match_score: float = 0.75  # Augmenter la barre de minimum
    note20_filter_by_class: bool = True  # Filtrer comptes classe 2 uniquement
    note20_prefilter_accounts: bool = True  # Pré-filtre les comptes avant matching
    note20_force_business_rules: bool = True  # Désactiver fuzzy pour NOTE 20, utiliser règles uniquement
    note20_garbage_collect: bool = True  # Force gc.collect() avant/après NOTE 20
    note20_detailed_timing: bool = True  # Logs détaillés de timing par étape

    def __post_init__(self) -> None:
        self.template_dsf = Path(self.template_dsf)
        self.balance_input = Path(self.balance_input)
        if self.previous_balance_input:
            self.previous_balance_input = Path(self.previous_balance_input)
        self.inventory_path = Path(self.inventory_path)
        self.rules_path = Path(self.rules_path)
        self.prefill_mapping_path = Path(self.prefill_mapping_path)
        self.prefilled_template_path = Path(self.prefilled_template_path)
        self.dsf_output = Path(self.dsf_output)
        self.report_json_path = Path(self.report_json_path)
        self.report_html_path = Path(self.report_html_path)
        if self.formula_reference_dsf:
            self.formula_reference_dsf = Path(self.formula_reference_dsf)

    @property
    def company_display_name(self) -> str:
        return getattr(self.general_info, "denomination_sociale", "ENTREPRISE")


class DSFPipeline:
    def __init__(
        self,
        config: DSFPipelineConfig,
        progress_callback: Optional[Callable[[int, str], None]] = None,
    ):
        self.config = config
        self.inventory: Optional[DSFInventory] = None
        self.rule_set: Optional[DSFRuleSet] = None
        self.progress_callback = progress_callback
        # Conserver une référence interne aux infos générales pour la validation pré-soumission
        self._general_info = self.config.general_info

    def _emit_progress(self, percent: int, message: str) -> None:
        percent = max(0, min(100, int(percent)))
        if self.progress_callback is None:
            return
        try:
            self.progress_callback(percent, message)
        except Exception as exc:
            logger.warning("Progress callback error: %s", exc)

    def run(self) -> PipelineArtifacts:
        start_time = time.time()
        self._emit_progress(3, "Vérification des entrées")
        self._validate_inputs()
        self._emit_progress(8, "Chargement de l'inventaire DSF")
        self.inventory = DSFInventory.from_json(self.config.inventory_path)
        self._emit_progress(15, "Initialisation du pipeline")

        logger.info(
            "Pipeline DSF %s - exercice %s (mode: %s)",
            self.config.company_display_name,
            self.config.exercice_year,
            self.config.filling_method,
        )

        self._emit_progress(22, "Pré-remplissage des sections générales")
        prefilled_template = self._prefill_general_sections()
        self._emit_progress(36, "Pré-remplissage terminé")

        if self.config.filling_method == "semantic":
            self._emit_progress(45, "Remplissage sémantique en cours")
            dsf_output = self._fill_semantic_balance(prefilled_template)
            self._emit_progress(80, "Remplissage sémantique terminé")
            reports = {"html": None, "json": None}
            
            # Apply calculations (TOTAL, VARIATION, RATIOS, etc)
            if self.config.apply_calculations:
                self._emit_progress(86, "Application des calculs DSF")
                self._apply_calculations(dsf_output)
                self._emit_progress(92, "Calculs appliqués")
        elif self.config.filling_method == "hybrid":
            # Mode hybride : sémantique en premier, puis overlay intelligent des règles
            self._emit_progress(44, "Chargement des règles pour mode hybride")
            self.rule_set = load_rule_set(self.config.rules_path)

            self._emit_progress(50, "Remplissage sémantique (hybride)")
            # 1. Remplissage sémantique sans overlay de règles
            dsf_output = self._fill_semantic_balance(
                prefilled_template,
                apply_note_overlay=False,
                apply_full_rule_overlay=False,
            )
            self._emit_progress(60, "Application moteur de règles (hybride)")

            # 2. Exécution du moteur de règles sur la même balance normalisée
            result_rules = self._run_rule_engine()

            # 3. Overlay prioritaire des règles sur certaines cellules (Bilan, CR, notes clés)
            self._emit_progress(70, "Overlay règles vs remplissage sémantique")
            self._apply_rule_overlay_hybrid_on_file(dsf_output, result_rules)

            reports = {"html": None, "json": None}
            if self.config.apply_calculations:
                self._emit_progress(88, "Application des calculs DSF")
                self._apply_calculations(dsf_output)
                self._emit_progress(93, "Calculs appliqués")
        else:
            self._emit_progress(46, "Chargement des règles")
            self.rule_set = load_rule_set(self.config.rules_path)
            self._emit_progress(56, "Affectation des comptes")
            result = self._run_rule_engine()
            # Construire les lignes normalisées pour Notes/TFT/validations en mode règles
            self._emit_progress(62, "Préparation des données N / N-1 pour Notes et TFT")
            self._build_normalized_rows_for_notes()
            self._emit_progress(72, "Contrôles de cohérence")
            controls = self._evaluate_controls(result)
            self._emit_progress(84, "Écriture du classeur DSF")
            dsf_output = self._write_assignments(prefilled_template, result)
            self._emit_progress(92, "Génération des rapports")
            reports = self._emit_reports(result, controls)

        # Enrichissements post-génération communs à tous les modes
        self._emit_progress(95, "Enrichissements post-génération (notes, flux, contrôles)")
        self._post_generation_enrichment(dsf_output)

        self._emit_progress(98, "Finalisation des sorties")
        duration = time.time() - start_time
        logger.info("Pipeline terminé en %.2fs", duration)
        logger.info("DSF généré: %s", dsf_output)
        if reports.get("html"):
            logger.info("Rapport HTML: %s", reports["html"])
        if reports.get("json"):
            logger.info("Rapport JSON: %s", reports["json"])
        self._emit_progress(100, "Génération DSF terminée")

        return PipelineArtifacts(
            dsf_output=dsf_output,
            prefilled_template=prefilled_template,
            report_json=reports.get("json"),
            report_html=reports.get("html"),
            controls_summary={},
            pipeline_warnings=getattr(self, "_pipeline_warnings", []),
        )

    # ------------------------------------------------------------------
    # Internal steps
    # ------------------------------------------------------------------
    def _validate_inputs(self) -> None:
        missing = [
            ("template DSF", self.config.template_dsf),
            ("balance", self.config.balance_input),
            ("inventaire", self.config.inventory_path),
            ("règles", self.config.rules_path),
            ("mapping préremplissage", self.config.prefill_mapping_path),
        ]
        if self.config.previous_balance_input:
            missing.append(("balance N-1", self.config.previous_balance_input))
        errors = [label for label, path in missing if not path.exists()]
        if errors:
            raise FileNotFoundError(f"Entrées manquantes: {', '.join(errors)}")

    def _build_normalized_rows_for_notes(self) -> None:
        """
        Construit les listes normalized_rows / previous_normalized_rows
        à partir des balances N et N-1 pour les modules Notes / TFT / validations.

        Utilisé notamment dans le mode purement 'rule-based' où _fill_semantic_balance
        n'est pas appelé et donc ne peuple pas ces attributs.
        """
        # Si déjà définis (cas du mode sémantique / hybride), ne rien refaire
        if hasattr(self, "_normalized_rows") and getattr(self, "_normalized_rows", None):
            return

        # Balance N
        normalizer_n = BalanceNormalizer(
            self.config.balance_input,
            chunk_size=self.config.chunk_size,
            column_overrides=self.config.balance_column_overrides,
        )
        try:
            self._normalized_rows = list(normalizer_n.iterate())
        finally:
            normalizer_n.close()

        # Balance N-1 (séparée ou dérivée des colonnes d'ouverture N)
        prev_overrides = self._resolve_previous_column_overrides()
        prev_file = self.config.previous_balance_input
        self._previous_normalized_rows = []
        if prev_file is None and prev_overrides is not None:
            prev_file = self.config.balance_input

        if prev_file is not None and prev_overrides is not None:
            normalizer_n1 = BalanceNormalizer(
                prev_file,
                chunk_size=self.config.chunk_size,
                column_overrides=prev_overrides,
            )
            try:
                self._previous_normalized_rows = list(normalizer_n1.iterate())
            finally:
                normalizer_n1.close()

    def _build_protected_cell_set(self) -> set[tuple[str, str]]:
        """
        Construit l'ensemble des cellules (feuille, cellule) considérées comme
        critiques et donc prioritairement alimentées par les règles en mode hybride.
        """
        protected: set[tuple[str, str]] = set(PROTECTED_CELLS)

        if self.inventory and self.rule_set:
            try:
                for rule in self.rule_set.rules:
                    if rule.section not in PROTECTED_SECTIONS:
                        continue
                    target = rule.target
                    for field in self.inventory.iter_fields(
                        target.sheet,
                        column_letter=target.column,
                        column_type=target.exercice_column_type,
                        row_min=target.row_start,
                        row_max=target.row_end,
                    ):
                        protected.add((field.sheet, field.cell))
            except Exception as exc:
                logger.debug("Impossible de construire l'ensemble des cellules protégées: %s", exc)

        return protected

    def _prefill_general_sections(self) -> Path:
        """
        Pré-remplit les sections générales (ENTÊTE, R1, R2, R3, NOTE13).
        Utilise soit le SmartGeneralFiller (reconnaissance auto) soit le DSFGeneralPrefiller (mapping manuel).
        """
        # Use a timestamp to avoid permission errors if the file is open
        import time
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        prefilled_path = self.config.prefilled_template_path.parent / f"dsf_prefilled_template_{timestamp}.xlsx"
        
        if self.config.use_smart_general_filler:
            # Nouveau système: reconnaissance automatique des zones de saisie
            logger.info("Mode: SmartGeneralFiller (reconnaissance automatique)")
            smart_filler = SmartGeneralFiller(self.config.template_dsf)
            smart_filler.load()
            filled = smart_filler.fill(self.config.general_info)
            smart_filler.save(prefilled_path)
            smart_filler.close()
            logger.info(f"Pré-remplissage intelligent: {filled} zones remplies")
        else:
            # Ancien système: mapping manuel JSON
            logger.info("Mode: DSFGeneralPrefiller (mapping manuel)")
            prefiller = DSFGeneralPrefiller(self.config.template_dsf, self.config.prefill_mapping_path)
            prefiller.load()
            filled = prefiller.fill(self.config.general_info)
            prefiller.save(prefilled_path)
            logger.info(f"Pré-remplissage manuel: {filled} éléments")
        
        logger.info("Pré-remplissage ENTETE/R1/R2/R3/NOTE13: %s éléments", filled)
        return prefilled_path

    def _apply_rule_overlay_hybrid_on_file(
        self,
        dsf_output: Path,
        result_rules: RuleEngineResult,
    ) -> None:
        """
        Mode hybride : applique les affectations issues du moteur de règles
        sur le DSF déjà rempli de façon sémantique.

        Priorité :
        - cellules protégées (Bilan / CR / notes clés) : règles > sémantique
        - autres cellules : on remplit uniquement si la cellule est vide / placeholder.
        """
        from openpyxl import load_workbook

        try:
            wb = load_workbook(dsf_output)
        except Exception as exc:
            logger.warning("Overlay hybride : impossible de rouvrir le DSF '%s' : %s", dsf_output, exc)
            return

        try:
            protected = self._build_protected_cell_set()

            for assign in result_rules.assignments:
                key = (assign.sheet, assign.cell)
                try:
                    ws = wb[assign.sheet]
                    cell = ws[assign.cell]
                except Exception as exc:
                    logger.debug("Overlay hybride : cellule %s!%s introuvable (%s)", assign.sheet, assign.cell, exc)
                    continue

                # Cas 1 : cellule critique → la règle est prioritaire
                if key in protected:
                    # Ne jamais écraser une formule existante
                    if isinstance(cell.value, str) and cell.value.lstrip().startswith("="):
                        continue
                    value = float(assign.amount)
                    cell.value = value
                    continue

                # Cas 2 : cellule non critique → on ne remplit que si vide / placeholder
                if self._is_placeholder(cell.value):
                    # Ne jamais écraser une formule existante
                    if isinstance(cell.value, str) and cell.value.lstrip().startswith("="):
                        continue
                    value = float(assign.amount)
                    cell.value = value
                    continue

                # Sinon : on laisse le remplissage sémantique / manuel tel quel

            wb.save(dsf_output)
        except Exception as exc:
            logger.warning("Overlay hybride : échec lors de l'application des règles : %s", exc)
        finally:
            try:
                wb.close()
            except Exception:
                pass

    @staticmethod
    def _is_placeholder(val) -> bool:
        """
        Vrai si la valeur de cellule peut être considérée comme vide / neutre
        (None, chaîne vide, ellipsis, tiret, zéro numérique).
        """
        if val is None:
            return True
        if isinstance(val, str):
            stripped = val.strip()
            return stripped in ("", "…", "...", "-", "_", "0")
        if isinstance(val, (int, float)) and val == 0:
            return True
        return False

    def _resolve_previous_column_overrides(self) -> Optional[Dict[str, object]]:
        """
        Retourne les column_overrides à utiliser pour la balance N-1.

        Principe : les comptes d'ouverture de l'année N sont les comptes de clôture de l'année N-1.
        On utilise donc les colonnes d'ouverture de la balance N pour remplir toutes les cellules
        du DSF qui demandent des infos N-1 (colonnes "N-1", "exercice précédent", etc.).

        - Si une balance N-1 séparée est fournie (previous_balance_input) :
            utilise previous_balance_column_overrides si défini, sinon balance_column_overrides.
        - Si aucune balance N-1 séparée et use_opening_columns_as_n1=True :
            construit des overrides qui pointent sur les colonnes d'ouverture de la balance N
            (opening_debit_columns / opening_credit_columns = soldes de clôture N-1).
        - Sinon : retourne None (pas de données N-1).
        """
        base = self.config.balance_column_overrides or {}

        if self.config.previous_balance_input:
            # Balance N-1 séparée fournie : utilise ses propres overrides si définis
            return self.config.previous_balance_column_overrides or base

        if self.config.use_opening_columns_as_n1:
            # Pas de balance N-1 séparée : extrait N-1 depuis les colonnes d'ouverture de la balance N
            opening_debit = base.get("opening_debit_columns", [])
            opening_credit = base.get("opening_credit_columns", [])
            if opening_debit or opening_credit:
                return {
                    **base,
                    "debit_columns": opening_debit,
                    "credit_columns": opening_credit,
                    "movement_debit_columns": [],
                    "movement_credit_columns": [],
                    "closing_debit_columns": opening_debit,
                    "closing_credit_columns": opening_credit,
                }
            logger.warning(
                "use_opening_columns_as_n1=True mais aucune colonne d'ouverture définie "
                "dans balance_column_overrides — colonnes N-1 non renseignées."
            )

        return None

    def _fill_semantic_balance(
        self,
        prefilled_template: Path,
        *,
        apply_note_overlay: bool = True,
        apply_full_rule_overlay: bool = False,
    ) -> Path:
        """
        NEW: Semantic balance filling approach with NOTE 20 optimizations
        Analyzes each cell label, fuzzy-matches balance accounts, fills values
        
        Solution #3, #7: Adds garbage collection and detailed timing
        """
        if not self.inventory:
            raise RuntimeError("Inventaire non chargé")

        logger.info("Mode: Remplissage sémantique par analyse de libellés")
        
        # Solution #3: Force garbage collection before heavy lifting
        if self.config.note20_garbage_collect:
            logger.info("[NOTE 20 OPT] Forcing garbage collection before semantic fill...")
            gc.collect()

        # N-1 : balance séparée OU colonnes d'ouverture de la balance N (clôture N-1 = ouverture N)
        prev_overrides = self._resolve_previous_column_overrides()
        prev_file = self.config.previous_balance_input
        if prev_file is None and prev_overrides is not None:
            prev_file = self.config.balance_input
            logger.info(
                "N-1 non fourni : utilisation des colonnes d'ouverture de la balance N "
                "(clôture N-1 = ouverture N) pour remplir toutes les cellules N-1 du DSF."
            )

        # Solution #7: Detailed timing for filler initialization
        load_start = time.time()
        
        filler = SemanticBalanceFiller(
            prefilled_template,
            self.config.balance_input,
            self.inventory,
            fuzzy_threshold=self.config.fuzzy_threshold,
            column_overrides=self.config.balance_column_overrides,
            previous_balance_file=prev_file,
            previous_column_overrides=prev_overrides,
            allow_n1_fallback_without_prev=self.config.use_opening_columns_as_n1,
        )
        
        logger.info("[TIMING] Filler init: %.2f s", time.time() - load_start)
        
        load_start = time.time()
        filler.load()
        logger.info("[TIMING] Filler load: %.2f s", time.time() - load_start)
        
        # Solution #2, #5, #6: Apply NOTE 20 optimizations
        if self.config.note20_prefilter_accounts:
            self._optimize_filler_for_note20(filler)
        
        # Solution #7: Detailed timing for semantic fill
        fill_start = time.time()
        filled = filler.fill()
        fill_time = time.time() - fill_start
        logger.info("[TIMING] Filler fill: %.2f s | Assignations: %s (~%.2f ms/assgn)", 
                   fill_time, filled, (fill_time * 1000 / filled) if filled > 0 else 0)
        logger.info("Remplissage sémantique: %s assignations", filled)

        # Restore normal parameters after fill
        if self.config.note20_prefilter_accounts:
            self._post_note20_restore_filler(filler)

        # Stocker les rows normalisés pour A3 (Notes) et A4 (Flux de Trésorerie)
        self._normalized_rows = getattr(filler, "normalized_rows", None) or \
                                 getattr(filler, "_normalized_rows", None) or []
        self._previous_normalized_rows = getattr(filler, "previous_normalized_rows", None) or \
                                          getattr(filler, "_previous_normalized_rows", None) or []
        logger.debug("Rows N stockés: %d / Rows N-1 stockés: %d",
                     len(self._normalized_rows), len(self._previous_normalized_rows))
        
        # Solution #3: Force garbage collection after heavy processing
        if self.config.note20_garbage_collect:
            logger.info("[NOTE 20 OPT] Forcing garbage collection after semantic fill...")
            gc.collect()

        # Optional rule overlays:
        # - semantic mode: notes only
        # - hybrid mode: full rules on empty cells
        if apply_full_rule_overlay:
            try:
                if not self.rule_set:
                    self.rule_set = load_rule_set(self.config.rules_path)
                result = self._run_rule_engine(rule_set=self.rule_set)
                applied = self._apply_rule_assignments_to_workbook(
                    filler.wb,
                    result,
                    filler=filler,
                    notes_only=False,
                )
                logger.info("Hybrid: %s rule assignments added on top of semantic fill", applied)
            except Exception as exc:
                logger.warning("Hybrid: failed to apply global rule overlay: %s", exc)
        elif apply_note_overlay and self.config.apply_note_rules_in_semantic:
            try:
                if not self.rule_set:
                    self.rule_set = load_rule_set(self.config.rules_path)
                note_rule_set = self._filter_note_rules(self.rule_set)
                if note_rule_set.rules:
                    note_result = self._run_rule_engine(rule_set=note_rule_set)
                    applied = self._apply_rule_assignments_to_workbook(
                        filler.wb,
                        note_result,
                        filler=filler,
                        notes_only=True,
                    )
                    logger.info("Notes: %s assignments applied in semantic mode", applied)
                else:
                    logger.warning("Notes: no rules found for notes in ruleset")
            except Exception as exc:
                logger.warning("Notes: failed to apply note rules: %s", exc)

        # Validate and generate reports after note-rule enrichment so reports
        # reflect the same logical state as the saved workbook.
        validator = SemanticFillerValidator(filler)
        report = validator.validate()
        logger.info("Validation: %s cells / %.1f%% success", report.total_cells_processed, report.success_rate)
        validator.print_console_report()
        json_report_path = validator.generate_json_report(self.config.report_json_path)
        html_report_path = validator.generate_html_report(self.config.report_html_path)
        logger.info("Rapport JSON: %s", json_report_path)
        logger.info("Rapport HTML: %s", html_report_path)

        # Emit formula application log
        try:
            formulas_log_path = self.config.report_json_path.with_name("dsf_formulas_log.json")
            filler.write_formula_log(formulas_log_path)
            logger.info("Formulas log: %s", formulas_log_path)
        except Exception as exc:
            logger.warning("Formulas log not written: %s", exc)

        # Save filled workbook
        output_path = filler.save(self.config.dsf_output)
        return output_path

    def _run_rule_engine(self, rule_set: Optional[DSFRuleSet] = None) -> RuleEngineResult:
        """
        Exécute le moteur de règles DSF sur la balance normalisée.

        Cette méthode se limite à la production des affectations / statistiques.
        Tous les enrichissements post-génération (notes, TFT, contrôles fiscaux, etc.)
        sont gérés séparément dans _post_generation_enrichment.
        """
        if not self.inventory or not (rule_set or self.rule_set):
            raise RuntimeError("Inventaire ou règle manquante")

        logger.info("Using overrides: %s", self.config.balance_column_overrides)

        normalizer = BalanceNormalizer(
            self.config.balance_input,
            chunk_size=self.config.chunk_size,
            column_overrides=self.config.balance_column_overrides,
        )
        engine = RuleEngine(rule_set or self.rule_set, self.inventory)
        try:
            result = engine.apply(normalizer.iterate())
        finally:
            normalizer.close()

        logger.info(
            "Affectations: %s | Comptes ignorés: %s",
            len(result.assignments),
            len(result.unmatched),
        )
        return result

    def _filter_note_rules(self, rule_set: DSFRuleSet) -> DSFRuleSet:
        """Return a ruleset containing only note-related rules."""
        note_rules = []
        for rule in rule_set.rules:
            section = (rule.section or "").upper()
            sheet = (rule.target.sheet or "").upper()
            if section.startswith("NOTE") or "NOTE" in sheet:
                note_rules.append(rule)
        return DSFRuleSet(
            version=rule_set.version,
            template_name=rule_set.template_name,
            generated_from_inventory=rule_set.generated_from_inventory,
            rules=note_rules,
            metadata=rule_set.metadata,
        )

    def _optimize_filler_for_note20(self, filler: SemanticBalanceFiller) -> None:
        """SOLUTION #2, #5, #6 : Optimize filler parameters for NOTE 20 heavy lifting."""
        if not self.config.note20_prefilter_accounts:
            return
        
        logger.info("[NOTE 20 OPT] Pre-filtering accounts by class...")
        
        # Solution #5: Pre-filter accounts to class 2 (Immobilisations) only for NOTE 20
        if self.config.note20_filter_by_class:
            original_count = len(filler.balance_accounts)
            
            # Filter: keep only class 2 accounts + class 1 opening balances (structure)
            filtered = {}
            for compte, row in filler.balance_accounts.items():
                try:
                    premiere_digit = compte[0] if compte else ""
                    # Pour NOTE 20 (immobilisations), garder classe 2 + classe 1 (base)
                    if premiere_digit in ("1", "2"):
                        filtered[compte] = row
                except (IndexError, TypeError):
                    pass
            
            filler.balance_accounts = filtered
            filtered_count = len(filler.balance_accounts)
            reduction_pct = ((original_count - filtered_count) / original_count * 100) if original_count > 0 else 0
            logger.info(f"[NOTE 20 OPT] Filtered {original_count} → {filtered_count} accounts (-{reduction_pct:.1f}%)")
        
        # Solution #6: Increase fuzzy threshold to reduce matching cost
        if self.config.note20_force_business_rules:
            old_threshold = filler.fuzzy_threshold
            filler.fuzzy_threshold = self.config.note20_fuzzy_threshold
            filler.min_match_score = self.config.note20_min_match_score
            logger.info(f"[NOTE 20 OPT] Fuzzy threshold: {old_threshold} → {filler.fuzzy_threshold}")
            logger.info(f"[NOTE 20 OPT] Min match score: → {filler.min_match_score}")

    def _post_note20_restore_filler(self, filler: SemanticBalanceFiller) -> None:
        """Restore normal filler parameters after NOTE 20 processing."""
        if not self.config.note20_prefilter_accounts:
            return
        
        # Restore normal parameters for remaining sheets
        filler.fuzzy_threshold = self.config.fuzzy_threshold
        filler.min_match_score = max(0.62, self.config.fuzzy_threshold)
        logger.info("[NOTE 20 OPT] Restored normal fuzzy parameters")

    def _apply_calculations(self, dsf_output: Path) -> None:
        """
        Applique les calculs (TOTAL, VARIATION, RATIOS) au DSF.
        Utilise le DSFCalculationApplier.
        """
        try:
            # P3- Lot B : Application des calculs (formules Excel ou valeurs)
            count = apply_calculations_to_dsf(
                dsf_output,
                use_formulas=self.config.use_calculation_formulas
            )
            logger.info("✓ %d calculs appliqués à %s", count, dsf_output.name)
        except Exception as exc:
            logger.warning("Échec de l'application des calculs : %s", exc)

    def _post_generation_enrichment(self, dsf_output: Path) -> None:
        """
        Étapes post-génération communes :
        - Remplissage Notes Annexes
        - Bilan composite (liens vers notes)
        - Tableau des Flux de Trésorerie
        - Contrôles fiscaux
        - Détection de soldes anormaux
        - Validation pré-soumission DGI
        """
        warnings_collected: list[str] = []

        # Tentative de réouverture du DSF pour enrichissements (notes / TFT / bilan composite)
        try:
            from openpyxl import load_workbook

            wb = load_workbook(dsf_output)
        except Exception as exc:
            msg = f"Post-génération : impossible de rouvrir le DSF pour enrichissements : {exc}"
            logger.warning(msg)
            self._pipeline_warnings = [msg]
            return

        normalized_rows = getattr(self, "_normalized_rows", [])
        previous_rows = getattr(self, "_previous_normalized_rows", [])

        # P2-A3 : Remplissage des Notes Annexes
        if self.config.enable_notes_filler:
            try:
                n_notes = fill_notes_from_balance(wb, normalized_rows, previous_rows)
                logger.info("✓ Notes annexes : %d lignes remplies", n_notes)
            except Exception as exc:
                msg = f"Notes annexes : {exc}"
                logger.warning(msg)
                warnings_collected.append(msg)

        # P3-B1 : Remplissage des cellules composites du BILAN
        try:
            n_bilan = fill_bilan_composite_cells(wb)
            if n_bilan:
                logger.info("✓ Bilan composite : %d cellules écrites", n_bilan)
        except Exception as exc:
            msg = f"Bilan composite : {exc}"
            logger.warning(msg)
            warnings_collected.append(msg)

        # P2-A4 : Tableau des Flux de Trésorerie
        if self.config.enable_cash_flow:
            try:
                n_tft = fill_cash_flow(wb, normalized_rows, previous_rows)
                logger.info("✓ Tableau des Flux : %d cellules écrites", n_tft)
            except Exception as exc:
                msg = f"Tableau des Flux de Trésorerie : {exc}"
                logger.warning(msg)
                warnings_collected.append(msg)

        # Sauvegarde du workbook enrichi
        try:
            wb.save(dsf_output)
        except Exception as exc:
            msg = f"Sauvegarde du DSF enrichi : {exc}"
            logger.warning(msg)
            warnings_collected.append(msg)
        finally:
            try:
                wb.close()
            except Exception:
                pass

        # P2-A1 : Contrôles fiscaux post-génération
        if self.config.enable_fiscal_controls:
            try:
                fiscal_report: FiscalReport = run_fiscal_controls(dsf_output)
                summary = fiscal_report.summary
                logger.info(
                    "✓ Contrôles fiscaux : %d PASS / %d WARN / %d FAIL",
                    summary["PASS"],
                    summary["WARN"],
                    summary["FAIL"],
                )
                if fiscal_report.has_failures():
                    msg = f"Contrôles fiscaux : {summary['FAIL']} contrôle(s) en ÉCHEC"
                    logger.error("ATTENTION : %s", msg)
                    warnings_collected.append(msg)
            except Exception as exc:
                msg = f"Contrôles fiscaux : {exc}"
                logger.warning(msg)
                warnings_collected.append(msg)

        # P3-MF4 : Détection soldes anormaux
        if normalized_rows:
            try:
                anomaly_reports = detect_abnormal_balances(normalized_rows)
                n_error = sum(1 for r in anomaly_reports if r.severity == "ERROR")
                n_warn = sum(1 for r in anomaly_reports if r.severity == "WARNING")
                if n_error > 0:
                    msg = f"P3-MF4 : {n_error} solde(s) anormal(aux) ERREUR détecté(s) dans la balance"
                    logger.error(msg)
                    warnings_collected.append(msg)
                elif n_warn > 0:
                    logger.warning("P3-MF4 : %d solde(s) en WARNING", n_warn)
            except Exception as exc:
                logger.debug("P3-MF4 ignoré : %s", exc)

        # P3-MF1 : Validation pré-soumission DGI
        try:
            from openpyxl import load_workbook as _lwb

            _wb_val = _lwb(dsf_output, data_only=True)
            _balance = normalized_rows
            _gi = getattr(self, "_general_info", None)
            presubmit_report = run_presubmission_validation(_wb_val, _balance, _gi)
            _wb_val.close()
            if not presubmit_report.is_valid:
                msg = f"P3-MF1 validation DGI : {presubmit_report.n_fail} FAIL(s) détecté(s)"
                logger.error(msg)
                warnings_collected.append(msg)
            elif presubmit_report.n_warn > 0:
                logger.warning("P3-MF1 : %d WARN pré-soumission DGI", presubmit_report.n_warn)
        except Exception as exc:
            logger.debug("P3-MF1 ignoré : %s", exc)

        # P3-m1 : Seuil d'alerte si trop de warnings
        self._pipeline_warnings = warnings_collected
        if len(warnings_collected) > 5:
            logger.error(
                "RAPPORT INCOMPLET — %d avertissements détectés. "
                "Vérifiez le DSF avant soumission DGI.",
                len(warnings_collected),
            )
        elif warnings_collected:
            logger.warning(
                "%d avertissement(s) calcul : %s",
                len(warnings_collected),
                " | ".join(warnings_collected),
            )


    def _emit_reports(self, result: RuleEngineResult, controls: ControlReport) -> Dict[str, Optional[Path]]:
        json_parent = self.config.report_json_path.parent
        html_parent = self.config.report_html_path.parent
        json_parent.mkdir(parents=True, exist_ok=True)
        html_parent.mkdir(parents=True, exist_ok=True)
        title = f"Rapport DSF {self.config.company_display_name} {self.config.exercice_year}"
        return generate_reports(
            result,
            controls,
            json_path=self.config.report_json_path,
            html_path=self.config.report_html_path,
            title=title,
        )


__all__ = ["DSFPipeline", "DSFPipelineConfig", "PipelineArtifacts"]


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    config = DSFPipelineConfig(enable_notes_filler=True)
    pipeline = DSFPipeline(config)
    artifacts = pipeline.run()
    print("\n✓ DSF généré:", artifacts.dsf_output)
    if artifacts.report_html:
        print("Rapport HTML:", artifacts.report_html)
    if artifacts.report_json:
        print("Rapport JSON:", artifacts.report_json)

if __name__ == "__main__":
    main()
