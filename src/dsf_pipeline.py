# -*- coding: utf-8 -*-
"""DSF production pipeline chaining normalization, rules, controls, and reporting."""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Callable, Dict, Optional

from balance_normalizer import BalanceNormalizer
from dsf_controls import ControlReport, ControlSuite
from dsf_general_info import DSF_InfosGenerales, get_gulfcam_config
from dsf_general_prefill import DSFGeneralPrefiller
from smart_general_filler import SmartGeneralFiller
from dsf_inventory import DSFInventory
from dsf_reporting import generate_reports
from dsf_rule_confihttps://github.com/lightGroup25/dsf/pull/1/conflict?name=src%252Fdsf_pipeline.py&ancestor_oid=351480c4c66912f2302c6d9c7dc1b973cf451ff5&base_oid=b5edf5a408da02a2695d643914244031c9fe2f46&head_oid=2ddd07ba9f154f3d7010afcb241ffa82e696222bg import DSFRuleSet, load_rule_set
from dsf_rule_engine import RuleEngine, RuleEngineResult
from dsf_secure_writer import ProtectedDSFWriter
from semantic_balance_filler import CellAssignment, SemanticBalanceFiller
from semantic_validators import SemanticFillerValidator
from openpyxl.utils import coordinate_to_tuple, range_boundaries
from dsf_calculation_applier import DSFCalculationApplier
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
    # P2 — Remplissage Notes Annexes 1-12 (A3)
    enable_notes_filler: bool = True
    # P2 — Tableau des Flux de Trésorerie (A4)
    enable_cash_flow: bool = True
    # Extraction et application des formules Excel depuis un DSF référence
    formula_reference_dsf: Optional[Path] = Path("input/DSF GULFCAM 2023 V3.xlsx")  # Si None, utilise template_dsf
    enable_formula_extraction: bool = True

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

    def _resolve_previous_column_overrides(self) -> Optional[Dict[str, object]]:
        """
        Retourne les column_overrides à utiliser pour la balance N-1.

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
        NEW: Semantic balance filling approach
        Analyzes each cell label, fuzzy-matches balance accounts, fills values
        """
        if not self.inventory:
            raise RuntimeError("Inventaire non chargé")

        logger.info("Mode: Remplissage sémantique par analyse de libellés")

        # N-1 : balance séparée OU colonnes d'ouverture de la balance N (use_opening_columns_as_n1)
        prev_overrides = self._resolve_previous_column_overrides()
        prev_file = self.config.previous_balance_input
        if prev_file is None and prev_overrides is not None:
            prev_file = self.config.balance_input

        filler = SemanticBalanceFiller(
            prefilled_template,
            self.config.balance_input,
            self.inventory,
            fuzzy_threshold=self.config.fuzzy_threshold,
            column_overrides=self.config.balance_column_overrides,
            previous_balance_file=prev_file,
            previous_column_overrides=prev_overrides,
        )
        filler.load()
        filled = filler.fill()
        logger.info("Remplissage sémantique: %s assignations", filled)

        # Stocker les rows normalisés pour A3 (Notes) et A4 (Flux de Trésorerie)
        self._normalized_rows = getattr(filler, "normalized_rows", None) or \
                                 getattr(filler, "_normalized_rows", None) or []
        self._previous_normalized_rows = getattr(filler, "previous_normalized_rows", None) or \
                                          getattr(filler, "_previous_normalized_rows", None) or []
        logger.debug("Rows N stockés: %d / Rows N-1 stockés: %d",
                     len(self._normalized_rows), len(self._previous_normalized_rows))

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

    def _apply_rule_assignments_to_workbook(
        self,
        wb,
        result: RuleEngineResult,
        filler: Optional[SemanticBalanceFiller] = None,
        *,
        notes_only: bool = True,
    ) -> int:
        if not self.inventory:
            raise RuntimeError("Inventaire non chargé")
        applied = 0
        skipped = 0
        appended = 0
        removed_unmatched = 0
        existing_keys = set()
        if filler is not None:
            existing_keys = {(a.sheet, a.cell) for a in filler.assignments}
        for assignment in result.assignments:
            sheet_upper = assignment.sheet.upper()
            if notes_only and "NOTE" not in sheet_upper:
                skipped += 1
                continue
            field = self.inventory.get_field(assignment.sheet, assignment.cell)
            if not field:
                skipped += 1
                continue
            if self._is_n1_column_type(getattr(field, "column_type", None)):
                # Rule engine currently reads only balance N. Never let it feed N-1 cells.
                skipped += 1
                continue
            if field.has_formula:
                skipped += 1
                continue
            if field.locked:
                if "NOTE" not in sheet_upper and "SYNTHESE" not in sheet_upper and "SYNTH" not in sheet_upper:
                    skipped += 1
                    continue
            row, col = coordinate_to_tuple(assignment.cell)
            if getattr(field, "merged", False) and field.merge_range:
                min_col, min_row, _, _ = range_boundaries(field.merge_range)
                row, col = min_row, min_col
            ws = wb[assignment.sheet]
            cell = ws.cell(row=row, column=col)
            if cell.value not in (None, ""):
                skipped += 1
                continue
            cell.value = assignment.amount
            applied += 1
            if filler is not None:
                key = (assignment.sheet, assignment.cell)
                if key not in existing_keys:
                    filler.assignments.append(
                        CellAssignment(
                            sheet=assignment.sheet,
                            cell=assignment.cell,
                            row_label=field.label or "",
                            col_label=field.column_type or "",
                            is_total=False,
                            matched_accounts=[],
                            total_amount=assignment.amount,
                            source_accounts=[
                                acc.get("compte", "")
                                for acc in (assignment.accounts or [])
                                if isinstance(acc, dict)
                            ],
                            confidence=0.88,
                            notes="source=rule_engine_note",
                        )
                    )
                    existing_keys.add(key)
                    appended += 1
                before = len(filler.unmatched_cells)
                filler.unmatched_cells = [
                    need
                    for need in filler.unmatched_cells
                    if not (need.sheet == assignment.sheet and need.cell == assignment.cell)
                ]
                removed_unmatched += max(0, before - len(filler.unmatched_cells))
        logger.info(
            "Rule overlay (notes_only=%s): applied=%s skipped=%s report_sync_added=%s report_sync_removed_unmatched=%s",
            notes_only,
            applied,
            skipped,
            appended,
            removed_unmatched,
        )
        return applied

    @staticmethod
    def _is_n1_column_type(column_type: Optional[str]) -> bool:
        txt = (column_type or "").lower().replace(" ", "")
        return "n-1" in txt or "n1" in txt or "exercice_n1" in txt

    def _write_assignments(self, template: Path, result: RuleEngineResult) -> Path:
        if not self.inventory:
            raise RuntimeError("Inventaire non chargé")
        writer = ProtectedDSFWriter(template, self.config.dsf_output, self.inventory, chunk_size=self.config.writer_chunk_size)
        try:
            safe_assignments = []
            skipped_n1 = 0
            for assignment in result.assignments:
                field = self.inventory.get_field(assignment.sheet, assignment.cell)
                if field and self._is_n1_column_type(getattr(field, "column_type", None)):
                    skipped_n1 += 1
                    continue
                safe_assignments.append(assignment)
            if skipped_n1:
                logger.info("Rule-based: %s assignments N-1 ignorées (balance N-1 non supportée par rule-engine)", skipped_n1)
            writer.write_assignments(safe_assignments)
            return writer.save()
        finally:
            writer.close()

    def _evaluate_controls(self, result: RuleEngineResult) -> ControlReport:
        if not self.rule_set:
            raise RuntimeError("Règles non chargées")
        suite = ControlSuite(self.rule_set, tolerance=self.config.control_tolerance)
        report = suite.evaluate(result)
        logger.info("Contrôles: %s", report.summary)
        return report

    def _apply_calculations(self, dsf_output: Path) -> None:
        """
        Apply automatic calculations to DSF:
        - Generate =SUM() formulas for TOTAL rows (P3-M5 : formules Excel préservées)
        - Calculate VARIATION (Closing - Opening)
        - Calculate PERCENTAGES and RATIOS
        P3-m1 : erreurs non silencieuses — warnings collectés, alerte si > 5
        """
        warnings_collected: list = []

        # --- Calculs DSF (TOTAL, VARIATION, RATIOS) ---
        try:
            # P3- Lot B : Application des calculs (formules Excel ou valeurs)
            count = apply_calculations_to_dsf(
                dsf_output,
                use_formulas=self.config.use_calculation_formulas,
                has_previous_balance_n1=bool(
                self.config.previous_balance_input
                or (self.config.use_opening_columns_as_n1 and self._resolve_previous_column_overrides())
            ),
                enforce_n1_from_previous_only=True,
                overwrite_formulas=False,  # P3-M5 : préserver les formules Excel du template
            )
            count = applier.apply()
            logger.info(
                "✓ %d calculs appliqués (%d formules Excel du template préservées)",
                count, applier.preserved_formulas,
            )
        except Exception as e:
            msg = f"Calculs DSF : {e}"
            logger.warning(msg)
            warnings_collected.append(msg)

        # --- Extraction et application des formules depuis le DSF référence ---
        if self.config.enable_formula_extraction:
            try:
                from dsf_formula_extractor_applier import extract_and_apply_formulas
                ref_dsf = self.config.formula_reference_dsf or self.config.template_dsf
                if ref_dsf.exists():
                    n_formulas = extract_and_apply_formulas(
                        ref_dsf,
                        dsf_output,
                        overwrite_existing_formulas=False,
                        overwrite_filled_cells=False,
                    )
                    if n_formulas > 0:
                        logger.info("✓ %d formules Excel appliquées depuis %s", n_formulas, ref_dsf.name)
                else:
                    logger.debug("DSF formule référence absent: %s", ref_dsf)
            except Exception as exc:
                logger.warning("Extraction formules: %s", exc)

        # P2-A3 : Remplissage des Notes Annexes 1-12
        if self.config.enable_notes_filler and hasattr(self, "_normalized_rows"):
            try:
                from openpyxl import load_workbook
                wb = load_workbook(dsf_output)
                prev_rows = getattr(self, "_previous_normalized_rows", None)
                n_notes = fill_notes_from_balance(wb, self._normalized_rows, prev_rows)
                wb.save(dsf_output)
                wb.close()
                logger.info("✓ Notes Annexes : %d lignes remplies", n_notes)
            except Exception as exc:
                msg = f"Notes Annexes : {exc}"
                logger.warning(msg)
                warnings_collected.append(msg)

        # P3-B1 : Remplissage cellules composites BILAN (références aux Notes)
        try:
            from openpyxl import load_workbook
            wb = load_workbook(dsf_output)
            n_composite = fill_bilan_composite_cells(wb)
            wb.save(dsf_output)
            wb.close()
            if n_composite > 0:
                logger.info("✓ Cellules composites BILAN : %d cellules remplies", n_composite)
        except Exception as exc:
            msg = f"Cellules composites BILAN : {exc}"
            logger.warning(msg)
            warnings_collected.append(msg)

        # P2-A4 : Tableau des Flux de Trésorerie
        if self.config.enable_cash_flow and hasattr(self, "_normalized_rows"):
            try:
                from openpyxl import load_workbook
                wb = load_workbook(dsf_output)
                prev_rows = getattr(self, "_previous_normalized_rows", None)
                n_tft = fill_cash_flow(wb, self._normalized_rows, prev_rows)
                wb.save(dsf_output)
                wb.close()
                logger.info("✓ Tableau des Flux : %d cellules écrites", n_tft)
            except Exception as exc:
                msg = f"Tableau des Flux de Trésorerie : {exc}"
                logger.warning(msg)
                warnings_collected.append(msg)

        # P2-A1 : Contrôles fiscaux post-génération
        if self.config.enable_fiscal_controls:
            try:
                fiscal_report: FiscalReport = run_fiscal_controls(dsf_output)
                summary = fiscal_report.summary
                logger.info(
                    "✓ Contrôles fiscaux : %d PASS / %d WARN / %d FAIL",
                    summary["PASS"], summary["WARN"], summary["FAIL"],
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
        if hasattr(self, "_normalized_rows"):
            try:
                anomaly_reports = detect_abnormal_balances(self._normalized_rows)
                n_error = sum(1 for r in anomaly_reports if r.severity == "ERROR")
                n_warn  = sum(1 for r in anomaly_reports if r.severity == "WARNING")
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
            _wb_val = _lwb(dsf_output)
            _balance = getattr(self, "_normalized_rows", [])
            _gi      = getattr(self, "_general_info", None)
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
