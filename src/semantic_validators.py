# -*- coding: utf-8 -*-
"""
Semantic Filler Validators - Validation and reporting
"""
from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Dict, List

from semantic_balance_filler import CellAssignment, CellNeed, SemanticBalanceFiller

logger = logging.getLogger(__name__)


@dataclass
class ValidationReport:
    """Comprehensive validation report"""
    total_cells_processed: int
    successful_assignments: int
    unmatched_cells: int
    success_rate: float
    total_amount_assigned: float
    high_confidence_assignments: int
    medium_confidence_assignments: int
    low_confidence_assignments: int
    assignments_by_sheet: Dict[str, int]
    unmatched_by_sheet: Dict[str, int]


class SemanticFillerValidator:
    def __init__(self, filler: SemanticBalanceFiller):
        self.filler = filler

    def validate(self) -> ValidationReport:
        """Perform comprehensive validation"""
        assignments = self.filler.assignments
        
        # Safety: ensure unmatched_cells is always a list
        unmatched = self.filler.unmatched_cells
        if not isinstance(unmatched, list):
            unmatched = [unmatched] if unmatched else []

        total_processed = len(assignments) + len(unmatched)
        success_rate = (len(assignments) / total_processed * 100) if total_processed > 0 else 0

        # Group by confidence
        high = len([a for a in assignments if a.confidence > 0.8])
        medium = len([a for a in assignments if 0.5 <= a.confidence <= 0.8])
        low = len([a for a in assignments if a.confidence < 0.5])

        # Group by sheet
        assignments_by_sheet = {}
        unmatched_by_sheet = {}

        for assignment in assignments:
            sheet = assignment.sheet
            assignments_by_sheet[sheet] = assignments_by_sheet.get(sheet, 0) + 1

        for unmatched_cell in unmatched:
            sheet = unmatched_cell.sheet
            unmatched_by_sheet[sheet] = unmatched_by_sheet.get(sheet, 0) + 1

        total_amount = sum(float(a.total_amount) for a in assignments)

        report = ValidationReport(
            total_cells_processed=total_processed,
            successful_assignments=len(assignments),
            unmatched_cells=len(unmatched),
            success_rate=success_rate,
            total_amount_assigned=total_amount,
            high_confidence_assignments=high,
            medium_confidence_assignments=medium,
            low_confidence_assignments=low,
            assignments_by_sheet=assignments_by_sheet,
            unmatched_by_sheet=unmatched_by_sheet,
        )

        return report

    def generate_json_report(self, output_path: Path | str) -> Path:
        """Generate detailed JSON report"""
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)

        report = self.validate()
        data = {
            "validation": asdict(report),
            "assignments": [
                {
                    "sheet": a.sheet,
                    "cell": a.cell,
                    "row_label": a.row_label,
                    "col_label": a.col_label,
                    "is_total": a.is_total,
                    "total_amount": float(a.total_amount),
                    "confidence": a.confidence,
                    "source_accounts": a.source_accounts,
                    "matched_accounts": [
                        {
                            "compte": m.compte,
                            "label": m.label,
                            "amount": float(m.amount),
                            "side": m.side,
                            "similarity": m.similarity_score,
                        }
                        for m in a.matched_accounts
                    ],
                    "notes": a.notes,
                }
                for a in self.filler.assignments
            ],
            "unmatched": [
                {
                    "sheet": u.sheet,
                    "cell": u.cell,
                    "row_label": u.row_label,
                    "col_label": u.col_label,
                }
                for u in self.filler.unmatched_cells
            ],
        }

        with output.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        logger.info(f"JSON report saved: {output}")
        return output

    def generate_html_report(self, output_path: Path | str) -> Path:
        """Generate pretty HTML validation report"""
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)

        report = self.validate()

        html = f"""
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport de Validation DSF</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #1e3a5f;
            border-bottom: 3px solid #0ea5e9;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #0ea5e9;
            margin-top: 30px;
        }}
        .stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }}
        .stat-card {{
            background: linear-gradient(135deg, #1e3a5f, #0ea5e9);
            color: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }}
        .stat-value {{
            font-size: 28px;
            font-weight: bold;
            margin: 10px 0;
        }}
        .stat-label {{
            font-size: 14px;
            opacity: 0.9;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }}
        th {{
            background: #1e3a5f;
            color: white;
            padding: 12px;
            text-align: left;
        }}
        td {{
            padding: 12px;
            border-bottom: 1px solid #ddd;
        }}
        tr:hover {{
            background: #f5f5f5;
        }}
        .success {{
            background: #d4edda;
            color: #155724;
        }}
        .warning {{
            background: #fff3cd;
            color: #856404;
        }}
        .danger {{
            background: #f8d7da;
            color: #721c24;
        }}
        .footer {{
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            font-size: 12px;
            color: #666;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Rapport de Validation - Remplissage Sémantique DSF</h1>
        
        <div class="stats">
            <div class="stat-card">
                <div class="stat-label">Cellules Traitées</div>
                <div class="stat-value">{report.total_cells_processed}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Assignations Réussies</div>
                <div class="stat-value">{report.successful_assignments}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Taux de Succès</div>
                <div class="stat-value">{report.success_rate:.1f}%</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Montant Total</div>
                <div class="stat-value">{report.total_amount_assigned:,.0f}</div>
            </div>
        </div>

        <h2>🎯 Confiance des Assignations</h2>
        <table>
            <tr>
                <th>Niveau</th>
                <th>Nombre</th>
                <th>Pourcentage</th>
            </tr>
            <tr class="success">
                <td>Haute (>80%)</td>
                <td>{report.high_confidence_assignments}</td>
                <td>{report.high_confidence_assignments / max(report.successful_assignments, 1) * 100:.1f}%</td>
            </tr>
            <tr class="warning">
                <td>Moyenne (50-80%)</td>
                <td>{report.medium_confidence_assignments}</td>
                <td>{report.medium_confidence_assignments / max(report.successful_assignments, 1) * 100:.1f}%</td>
            </tr>
            <tr class="danger">
                <td>Basse (<50%)</td>
                <td>{report.low_confidence_assignments}</td>
                <td>{report.low_confidence_assignments / max(report.successful_assignments, 1) * 100:.1f}%</td>
            </tr>
        </table>

        <h2>📑 Assignations par Feuille</h2>
        <table>
            <tr>
                <th>Feuille</th>
                <th>Assignées</th>
                <th>Non-Assignées</th>
            </tr>
"""

        for sheet in sorted(set(list(report.assignments_by_sheet.keys()) + list(report.unmatched_by_sheet.keys()))):
            assigned = report.assignments_by_sheet.get(sheet, 0)
            unassigned = report.unmatched_by_sheet.get(sheet, 0)
            total_sheet = assigned + unassigned
            coverage = assigned / total_sheet * 100 if total_sheet > 0 else 0
            
            row_class = "success" if coverage == 100 else "warning" if coverage >= 80 else "danger"
            html += f"""
            <tr class="{row_class}">
                <td>{sheet}</td>
                <td>{assigned} ({coverage:.1f}%)</td>
                <td>{unassigned}</td>
            </tr>
"""

        html += """
        </table>

        <div class="footer">
            <p>Rapport généré automatiquement par le système de remplissage sémantique DSF</p>
            <p>Pour plus de détails, consulter le rapport JSON détaillé</p>
        </div>
    </div>
</body>
</html>
"""

        with output.open("w", encoding="utf-8") as f:
            f.write(html)

        logger.info(f"HTML report saved: {output}")
        return output

    def print_console_report(self) -> None:
        """Print validation report to console"""
        report = self.validate()

        print("\n" + "=" * 70)
        print("📊 RAPPORT DE VALIDATION - REMPLISSAGE SÉMANTIQUE DSF")
        print("=" * 70)
        print(f"\n✅ Cellules traitées: {report.total_cells_processed}")
        print(f"✓ Assignations réussies: {report.successful_assignments}")
        print(f"✗ Cellules non-assignées: {report.unmatched_cells}")
        print(f"📈 Taux de succès: {report.success_rate:.1f}%")
        print(f"💰 Montant total assigné: {report.total_amount_assigned:,.0f}")

        print("\n🎯 Confiance des assignations:")
        print(f"   Haute (>80%): {report.high_confidence_assignments}")
        print(f"   Moyenne (50-80%): {report.medium_confidence_assignments}")
        print(f"   Basse (<50%): {report.low_confidence_assignments}")

        print("\n📑 Assignations par feuille:")
        for sheet, count in sorted(report.assignments_by_sheet.items()):
            unmatched = report.unmatched_by_sheet.get(sheet, 0)
            total = count + unmatched
            coverage = count / total * 100 if total > 0 else 0
            print(f"   {sheet:40s}: {count:3d}/{total:3d} ({coverage:5.1f}%)")

        if report.unmatched_cells > 0:
            print("\n⚠️  Cellules non-assignées:")
            for unmatched in self.filler.unmatched_cells[:10]:  # Show first 10
                print(f"   {unmatched.sheet}!{unmatched.cell}: '{unmatched.row_label}'")

        print("\n" + "=" * 70 + "\n")


__all__ = ["SemanticFillerValidator", "ValidationReport"]
