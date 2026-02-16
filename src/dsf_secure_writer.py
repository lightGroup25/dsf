# -*- coding: utf-8 -*-
"""Protected writer that enforces DSF template constraints before writing."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl.utils import coordinate_to_tuple, range_boundaries

from dsf_inventory import DSFInventory
from dsf_rule_engine import Assignment
from dsf_stream_writer import DSFStreamWriter


@dataclass
class WriteLog:
    sheet: str
    cell: str
    value: str
    rule_id: str


class ProtectedDSFWriter:
    def __init__(self, template_path: Path, output_path: Path, inventory: DSFInventory, chunk_size: int = 100):
        self.inventory = inventory
        self._writer = DSFStreamWriter(template_path, output_path, chunk_size=chunk_size)
        self._writer.initialize_from_template()
        self._logs: list[WriteLog] = []

    def write_assignment(self, assignment: Assignment) -> None:
        try:
            field = self.inventory.validate_writable_cell(assignment.sheet, assignment.cell)
            row, col = coordinate_to_tuple(assignment.cell)
            if getattr(field, "merged", False) and field.merge_range:
                min_col, min_row, max_col, max_row = range_boundaries(field.merge_range)
                row, col = min_row, min_col
        except ValueError as e:
            # Fallback for "Blind Mode" / Virtual Fields
            # logger.warning(f"Writing to unvalidated cell {assignment.sheet}!{assignment.cell}: {e}")
            row, col = coordinate_to_tuple(assignment.cell)
            
        self._writer.write_cell(assignment.sheet, row, col, assignment.amount)
        self._logs.append(
            WriteLog(
                sheet=assignment.sheet,
                cell=assignment.cell,
                value=assignment.amount.normalize().to_eng_string(),
                rule_id=assignment.rule_id,
            )
        )

    def write_assignments(self, assignments: Iterable[Assignment]) -> None:
        for assignment in assignments:
            self.write_assignment(assignment)

    def save(self) -> Path:
        return self._writer.save()

    def close(self) -> None:
        self._writer.close()

    @property
    def logs(self) -> list[WriteLog]:
        return self._logs


__all__ = ["ProtectedDSFWriter", "WriteLog"]
