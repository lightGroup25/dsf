"""Comprehensive DSF formula scanner.

Detects both real Excel formulas (`=SUM(...)`) and textual formulas written in
headers/descriptions such as "D=A+B-C". Outputs a concise report per sheet.
"""

from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl

DSF_TEMPLATE = Path("templates/DSF Normal standard.xlsx")
OUTPUT_REPORT = Path("output/reports/dsf_formula_report.json")

# Allow references such as "D=A+B", "6=1+2+3" or "C = A - (B+C)".
TEXTUAL_FORMULA_PATTERN = re.compile(r"([A-Z0-9]{1,3}\s*=\s*[A-Z0-9\+\-\*/() ]{3,})")

# Scan deeper by default (some notes start after row 150 / column 30)
MAX_ROWS_TO_SCAN = 300
MAX_COLS_TO_SCAN = 45


def number_to_excel_column(index: int) -> str:
    """Convert a 1-based column number to its Excel column letter (1 -> A)."""
    if index <= 0:
        return str(index)
    letters = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return ''.join(reversed(letters))


def translate_to_excel(formula_text: str) -> dict:
    """Return a dict with lhs/rhs tokens and an Excel-style formula string."""
    normalized = formula_text.upper().replace(' ', '')
    if '=' not in normalized:
        return {"lhs": formula_text, "excel_formula": None}

    lhs, rhs = normalized.split('=', 1)

    def token_to_excel(token: str) -> str:
        if token.isdigit():
            return number_to_excel_column(int(token))
        if token.isalpha():
            return token
        return token

    tokens = re.findall(r"[A-Z]+|\d+|[\+\-\*/()]", rhs)
    converted = ''.join(token_to_excel(tok) for tok in tokens)
    excel_formula = f"={converted}" if converted else None
    return {"lhs": lhs, "excel_formula": excel_formula}


def load_workbook(path: Path):
    if not path.exists():
        raise FileNotFoundError(f"Template not found: {path}")
    # data_only=False so Excel formulas are preserved
    return openpyxl.load_workbook(path, data_only=False, read_only=False)


def scan_sheet(ws):
    """Return dictionaries of real and textual formulas for a worksheet."""
    real_formulas = []
    textual_formulas = []

    # Scan actual cell formulas
    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            if value is None:
                continue
            if cell.data_type == "f" or (isinstance(value, str) and value.startswith("=")):
                real_formulas.append({
                    "cell": cell.coordinate,
                    "formula": str(value),
                })

    # Scan textual descriptions (headers/notes) for embedded formulas
    max_rows_to_scan = min(ws.max_row, MAX_ROWS_TO_SCAN)
    max_cols_to_scan = min(ws.max_column, MAX_COLS_TO_SCAN)
    for row_idx in range(1, max_rows_to_scan + 1):
        for col_idx in range(1, max_cols_to_scan + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            value = cell.value
            if not isinstance(value, str):
                continue
            for match in TEXTUAL_FORMULA_PATTERN.finditer(value.upper()):
                extracted = match.group(1).strip()
                translation = translate_to_excel(extracted)
                textual_formulas.append({
                    "cell": cell.coordinate,
                    "text": extracted,
                    "lhs": translation.get("lhs"),
                    "excel_formula": translation.get("excel_formula"),
                })

    return real_formulas, textual_formulas


def main():
    wb = load_workbook(DSF_TEMPLATE)

    report = {
        "template": str(DSF_TEMPLATE),
        "sheets": [],
        "totals": {"real": 0, "textual": 0},
    }

    for ws in wb.worksheets:
        real_formulas, textual_formulas = scan_sheet(ws)
        report["totals"]["real"] += len(real_formulas)
        report["totals"]["textual"] += len(textual_formulas)

        if not real_formulas and not textual_formulas:
            continue

        sheet_entry = {
            "sheet": ws.title,
            "real_count": len(real_formulas),
            "textual_count": len(textual_formulas),
            "samples": {
                "real": real_formulas[:10],
                "textual": textual_formulas[:10],
            },
        }
        report["sheets"].append(sheet_entry)

    OUTPUT_REPORT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_REPORT.write_text(json.dumps(report, indent=2, ensure_ascii=False))

    print("=" * 90)
    print("DSF FORMULA SCAN SUMMARY")
    print("=" * 90)
    print(f"Sheets with formulas: {len(report['sheets'])}/{len(wb.sheetnames)}")
    print(f"Real Excel formulas found   : {report['totals']['real']:,}")
    print(f"Textual formula hints found: {report['totals']['textual']:,}")
    print("Sample (per sheet):")
    for sheet in report["sheets"][:15]:
        print(f"  - {sheet['sheet']}: real={sheet['real_count']} textual={sheet['textual_count']}")
    print(f"Detailed JSON report written to: {OUTPUT_REPORT}")


if __name__ == "__main__":
    main()
