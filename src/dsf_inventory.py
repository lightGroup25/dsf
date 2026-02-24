# -*- coding: utf-8 -*-
"""Helpers to consume the DSF inventory exported from dsf_template_inventory.py."""
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence


@dataclass(frozen=True)
class InventoryField:
    sheet: str
    cell: str
    row: int
    column: int
    column_letter: str
    column_type: str
    label: Optional[str]
    section_hint: Optional[str]
    locked: bool
    has_formula: bool
    number_format: Optional[str]
    data_validation: Sequence[Dict[str, str]]
    merged: bool
    merge_range: Optional[str]


class DSFInventory:
    """Index DSF template fields for fast rule lookups and validation."""

    def __init__(self, payload: Dict[str, object]):
        self.raw_payload = payload
        sheets = payload.get("sheets", {})
        self._fields_by_sheet: Dict[str, Dict[str, InventoryField]] = {}
        self._columns_by_sheet: Dict[str, Dict[str, List[InventoryField]]] = {}
        for sheet_name, sheet_payload in sheets.items():
            fields = {}
            columns: Dict[str, List[InventoryField]] = {}
            for field_data in sheet_payload.get("fields", []):
                field = InventoryField(
                    sheet=sheet_name,
                    cell=field_data["cell"],
                    row=int(field_data["row"]),
                    column=int(field_data["column"]),
                    column_letter=field_data["column_letter"],
                    column_type=field_data["column_type"],
                    label=field_data.get("label"),
                    section_hint=field_data.get("section_hint"),
                    locked=bool(field_data.get("locked", False)),
                    has_formula=bool(field_data.get("has_formula", False)),
                    number_format=field_data.get("number_format"),
                    data_validation=tuple(field_data.get("data_validation", [])),
                    merged=bool(field_data.get("merged", False)),
                    merge_range=field_data.get("merge_range"),
                )
                fields[field.cell] = field
                columns.setdefault(field.column_letter, []).append(field)
            for column_fields in columns.values():
                column_fields.sort(key=lambda f: f.row)
            self._fields_by_sheet[sheet_name] = fields
            self._columns_by_sheet[sheet_name] = columns

    @classmethod
    def from_json(cls, path: Path) -> "DSFInventory":
        payload = json.loads(path.read_text(encoding="utf-8"))
        return cls(payload)

    def to_dict(self) -> Dict[str, object]:  # pragma: no cover - debug helper
        result = {}
        for sheet, fields in self._fields_by_sheet.items():
            result[sheet] = {cell: field for cell, field in fields.items()}
        return result

    def get_field(self, sheet: str, cell: str) -> Optional[InventoryField]:
        return self._fields_by_sheet.get(sheet, {}).get(cell)

    def get_field_with_fallback(self, sheet: str, cell: str) -> Optional[InventoryField]:
        """Recherche un champ avec tolérance aux variations de nom (espaces, casse)."""
        field = self.get_field(sheet, cell)
        if field:
            return field
        # Essayer sans espaces de fin (ex: "NOTE 12 " -> "NOTE 12")
        alt = sheet.rstrip()
        if alt != sheet:
            field = self.get_field(alt, cell)
            if field:
                return field
        # Essayer avec espace final (inverse)
        alt2 = sheet if sheet.endswith(" ") else (sheet + " ")
        if alt2 in self._fields_by_sheet:
            return self._fields_by_sheet[alt2].get(cell)
        return None

    def iter_fields(
        self,
        sheet: str,
        column_letter: Optional[str] = None,
        column_type: Optional[str] = None,
        row_min: Optional[int] = None,
        row_max: Optional[int] = None,
    ) -> Iterable[InventoryField]:
        if sheet not in self._columns_by_sheet:
            return []
        columns = self._columns_by_sheet[sheet]
        if column_letter:
            candidates = columns.get(column_letter, [])
        else:
            candidates = [field for column in columns.values() for field in column]
        for field in candidates:
            if column_type and field.column_type != column_type:
                continue
            if row_min and field.row < row_min:
                continue
            if row_max and field.row > row_max:
                continue
            yield field

    def validate_writable_cell(self, sheet: str, cell: str, require_column_type: str = "exercice_n") -> InventoryField:
        field = self.get_field(sheet, cell)
        if not field:
            raise ValueError(f"Cellule {sheet}!{cell} absente de l'inventaire DSF")
        if field.locked:
            raise ValueError(f"Cellule {sheet}!{cell} verrouillée dans le template")
        if field.has_formula:
            raise ValueError(f"Cellule {sheet}!{cell} contient une formule et ne peut pas être écrite")
        if field.column_type != require_column_type:
            raise ValueError(
                f"Cellule {sheet}!{cell} appartient à {field.column_type}, attendu {require_column_type}"
            )
        return field

    def list_sheet_names(self) -> List[str]:
        return list(self._fields_by_sheet.keys())


__all__ = ["InventoryField", "DSFInventory"]
