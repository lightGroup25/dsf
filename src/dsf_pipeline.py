# -*- coding: utf-8 -*-
"""DSF production pipeline chaining normalization, rules, controls, and reporting."""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Dict, Optional

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
from semantic_balance_filler import SemanticBalanceFiller
from semantic_validators import SemanticFillerValidator
from openpyxl.utils import coordinate_to_tuple, range_boundaries
from dsf_calculation_applier import DSFCalculationApplier

logger = logging.getLogger(__name__)


@dataclass
class PipelineArtifacts:
    dsf_output: Path
    prefilled_template: Path
    report_json: Optional[Path]
    report_html: Optional[Path]
    controls_summary: Dict[str, int]


@dataclass
class DSFPipelineConfig:
    template_dsf: Path = Path("templates/DSF Normal standard.xlsx")
    balance_input: Path = Path("input/BALANCE AU 31 12 2024 VERSION DU 26 05 25.xlsx")
    inventory_path: Path = Path("data/dsf_inventory.json")
    rules_path: Path = Path("config/dsf_rule_generated.yaml")
    prefill_mapping_path: Path = Path("config/dsf_prefill_mapping.json")
    prefilled_template_path: Path = Path("output/dsf_prefilled_template.xlsx")
    dsf_output: Path = Path("output/DSF_OUTPUT_2024.xlsx")
    report_json_path: Path = Path("output/reports/dsf_report.json")
    report_html_path: Path = Path("output/reports/dsf_report.html")
    exercice_year: int = 2024
    chunk_size: int = 500
    writer_chunk_size: int = 200
    control_tolerance: Decimal = Decimal("1")
    filling_method: str = "semantic"  # "semantic" or "rule-based"
    fuzzy_threshold: float = 0.6  # For semantic filling
    apply_calculations: bool = True  # Apply TOTAL, VARIATION, RATIO formulas
    use_calculation_formulas: bool = True  # Use Excel formulas vs numeric values
    use_smart_general_filler: bool = True  # Use intelligent zone detection for ENTETE/R1/R2/R3/NOTE13
    apply_note_rules_in_semantic: bool = True  # Apply intelligent rules for all notes in semantic mode
    general_info: DSF_InfosGenerales = field(default_factory=get_gulfcam_config)
    balance_column_overrides: Optional[Dict[str, object]] = field(
        default_factory=lambda: {
            "compte": 1,
            "label": 4,
            "debit_columns": [24, 21, 17, 11],
            "credit_columns": [27, 26, 20, 14, 13],
        }
    )

    def __post_init__(self) -> None:
        self.template_dsf = Path(self.template_dsf)
        self.balance_input = Path(self.balance_input)
        self.inventory_path = Path(self.inventory_path)
        self.rules_path = Path(self.rules_path)
        self.prefill_mapping_path = Path(self.prefill_mapping_path)
        self.prefilled_template_path = Path(self.prefilled_template_path)
        self.dsf_output = Path(self.dsf_output)
        self.report_json_path = Path(self.report_json_path)
        self.report_html_path = Path(self.report_html_path)

    @property
    def company_display_name(self) -> str:
        return getattr(self.general_info, "denomination_sociale", "ENTREPRISE")


class DSFPipeline:
    def __init__(self, config: DSFPipelineConfig):
        self.config = config
        self.inventory: Optional[DSFInventory] = None
        self.rule_set: Optional[DSFRuleSet] = None

    def run(self) -> PipelineArtifacts:
        start_time = time.time()
        self._validate_inputs()
        self.inventory = DSFInventory.from_json(self.config.inventory_path)

        logger.info(
            "Pipeline DSF %s - exercice %s (mode: %s)",
            self.config.company_display_name,
            self.config.exercice_year,
            self.config.filling_method,
        )

        prefilled_template = self._prefill_general_sections()

        if self.config.filling_method == "semantic":
            dsf_output = self._fill_semantic_balance(prefilled_template)
            reports = {"html": None, "json": None}
            
            # Apply calculations (TOTAL, VARIATION, RATIOS, etc)
            if self.config.apply_calculations:
                self._apply_calculations(dsf_output)
        else:
            self.rule_set = load_rule_set(self.config.rules_path)
            result = self._run_rule_engine()
            controls = self._evaluate_controls(result)
            dsf_output = self._write_assignments(prefilled_template, result)
            reports = self._emit_reports(result, controls)

        duration = time.time() - start_time
        logger.info("Pipeline terminé en %.2fs", duration)
        logger.info("DSF généré: %s", dsf_output)
        if reports.get("html"):
            logger.info("Rapport HTML: %s", reports["html"])
        if reports.get("json"):
            logger.info("Rapport JSON: %s", reports["json"])

        return PipelineArtifacts(
            dsf_output=dsf_output,
            prefilled_template=prefilled_template,
            report_json=reports.get("json"),
            report_html=reports.get("html"),
            controls_summary={},
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
        errors = [label for label, path in missing if not path.exists()]
        if errors:
            raise FileNotFoundError(f"Entrées manquantes: {', '.join(errors)}")

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

    def _fill_semantic_balance(self, prefilled_template: Path) -> Path:
        """
        NEW: Semantic balance filling approach
        Analyzes each cell label, fuzzy-matches balance accounts, fills values
        """
        if not self.inventory:
            raise RuntimeError("Inventaire non chargé")

        logger.info("Mode: Remplissage sémantique par analyse de libellés")
        
        filler = SemanticBalanceFiller(
            prefilled_template,
            self.config.balance_input,
            self.inventory,
            fuzzy_threshold=self.config.fuzzy_threshold,
        )
        filler.load()
        filled = filler.fill()
        logger.info("Remplissage sémantique: %s assignations", filled)

        # Validate and generate reports
        validator = SemanticFillerValidator(filler)
        report = validator.validate()
        logger.info("Validation: %s cells / %.1f%% success", 
                   report.total_cells_processed, report.success_rate)

        # Print console report
        validator.print_console_report()

        # Generate detailed reports
        json_report_path = validator.generate_json_report(self.config.report_json_path)
        html_report_path = validator.generate_html_report(self.config.report_html_path)

        logger.info("Rapport JSON: %s", json_report_path)
        logger.info("Rapport HTML: %s", html_report_path)

        # Apply intelligent rules for notes (strict mode)
        if self.config.apply_note_rules_in_semantic:
            try:
                if not self.rule_set:
                    self.rule_set = load_rule_set(self.config.rules_path)
                note_rule_set = self._filter_note_rules(self.rule_set)
                if note_rule_set.rules:
                    note_result = self._run_rule_engine(rule_set=note_rule_set)
                    applied = self._apply_note_assignments_to_workbook(filler.wb, note_result)
                    logger.info("Notes: %s assignments applied in semantic mode", applied)
                else:
                    logger.warning("Notes: no rules found for notes in ruleset")
            except Exception as exc:
                logger.warning("Notes: failed to apply note rules: %s", exc)

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
        note_rules = []
        for rule in rule_set.rules:
            section = (rule.section or "").upper()
            target_sheet = (rule.target.sheet or "").upper()
            if section.startswith("NOTE") or "NOTE" in target_sheet:
                note_rules.append(rule)
        return DSFRuleSet(
            version=rule_set.version,
            template_name=rule_set.template_name,
            generated_from_inventory=rule_set.generated_from_inventory,
            rules=note_rules,
            metadata=dict(rule_set.metadata),
        )

    def _apply_note_assignments_to_workbook(self, wb, result: RuleEngineResult) -> int:
        if not self.inventory:
            raise RuntimeError("Inventaire non chargé")
        applied = 0
        skipped = 0
        for assignment in result.assignments:
            field = self.inventory.get_field(assignment.sheet, assignment.cell)
            if not field:
                skipped += 1
                continue
            if field.has_formula:
                skipped += 1
                continue
            if field.locked:
                sheet_upper = assignment.sheet.upper()
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
        logger.info("Notes: applied=%s skipped=%s", applied, skipped)
        return applied

    def _write_assignments(self, template: Path, result: RuleEngineResult) -> Path:
        if not self.inventory:
            raise RuntimeError("Inventaire non chargé")
        writer = ProtectedDSFWriter(template, self.config.dsf_output, self.inventory, chunk_size=self.config.writer_chunk_size)
        try:
            writer.write_assignments(result.assignments)
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
        - Generate =SUM() formulas for TOTAL rows
        - Calculate VARIATION (Closing - Opening)
        - Calculate PERCENTAGES and RATIOS
        - Preserve and adapt complex formulas from template
        """
        try:
            logger.info("Applying automatic calculations (TOTAL, VARIATION, RATIOS)...")
            applier = DSFCalculationApplier(
                dsf_output,
                use_formulas=self.config.use_calculation_formulas
            )
            count = applier.apply()
            logger.info(f"✓ {count} calculations applied to DSF")
        except Exception as e:
            logger.warning(f"Error applying calculations: {e}")
            # Don't fail pipeline if calculations error

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
    pipeline = DSFPipeline(DSFPipelineConfig())
    artifacts = pipeline.run()
    print("\n✓ DSF généré:", artifacts.dsf_output)
    if artifacts.report_html:
        print("Rapport HTML:", artifacts.report_html)
    if artifacts.report_json:
        print("Rapport JSON:", artifacts.report_json)

if __name__ == "__main__":
    main()
