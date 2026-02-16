# -*- coding: utf-8 -*-
"""Generation of JSON/HTML DSF reports (controls + assignments)."""
from __future__ import annotations

import html
from pathlib import Path
from typing import Any, Dict, Optional

from dsf_controls import ControlReport
from dsf_rule_engine import RuleEngineResult


def build_report_payload(result: RuleEngineResult, controls: ControlReport) -> Dict[str, Any]:
    return {
        "stats": result.stats,
        "assignments": [assignment.to_dict() for assignment in result.assignments],
        "unmatched": controls.unmatched_accounts,
        "controls": [issue.__dict__ for issue in controls.issues],
        "controls_summary": controls.summary,
    }


def write_json_report(payload: Dict[str, Any], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(_json_dumps(payload), encoding="utf-8")
    return path


def write_html_report(payload: Dict[str, Any], path: Path, title: str = "Rapport DSF") -> Path:
    controls_rows = "".join(
        f"<tr><td>{html.escape(issue['name'])}</td><td>{html.escape(issue['status'])}</td><td>{html.escape(issue['message'])}</td></tr>"
        for issue in payload.get("controls", [])
    )
    unmatched_rows = "".join(
        f"<tr><td>{html.escape(item.get('compte', ''))}</td><td>{html.escape(item.get('label', ''))}</td><td>{html.escape(item.get('reason', ''))}</td></tr>"
        for item in payload.get("unmatched", [])
    )
    html_body = f"""
<!DOCTYPE html>
<html lang=\"fr\">
<head>
<meta charset=\"utf-8\">
<title>{html.escape(title)}</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 24px; }}
h1 {{ font-size: 20px; }}
table {{ border-collapse: collapse; width: 100%; margin-bottom: 24px; }}
th, td {{ border: 1px solid #ccc; padding: 8px; font-size: 13px; text-align: left; }}
th {{ background-color: #f2f2f2; }}
.status-PASS {{ color: #0a7d00; }}
.status-WARN {{ color: #c77700; }}
.status-FAIL {{ color: #c10000; }}
</style>
</head>
<body>
<h1>{html.escape(title)}</h1>
<section>
<h2>Contrôles</h2>
<table>
<thead><tr><th>Contrôle</th><th>Status</th><th>Message</th></tr></thead>
<tbody>{controls_rows or '<tr><td colspan="3">Aucun contrôle exécuté</td></tr>'}</tbody>
</table>
</section>
<section>
<h2>Comptes non affectés</h2>
<table>
<thead><tr><th>Compte</th><th>Libellé</th><th>Raison</th></tr></thead>
<tbody>{unmatched_rows or '<tr><td colspan="3">Aucun</td></tr>'}</tbody>
</table>
</section>
</body>
</html>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html_body, encoding="utf-8")
    return path


def generate_reports(
    result: RuleEngineResult,
    controls: ControlReport,
    json_path: Optional[Path] = None,
    html_path: Optional[Path] = None,
    title: str = "Rapport DSF",
) -> Dict[str, Optional[Path]]:
    payload = build_report_payload(result, controls)
    outputs: Dict[str, Optional[Path]] = {"json": None, "html": None}
    if json_path:
        outputs["json"] = write_json_report(payload, json_path)
    if html_path:
        outputs["html"] = write_html_report(payload, html_path, title=title)
    return outputs


def _json_dumps(payload: Dict[str, Any]) -> str:
    import json

    return json.dumps(payload, ensure_ascii=False, indent=2)


__all__ = [
    "generate_reports",
    "build_report_payload",
    "write_json_report",
    "write_html_report",
]
