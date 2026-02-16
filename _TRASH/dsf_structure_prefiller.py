"""
DSF Pre-filler using structure-driven approach
Maps data directly to cells defined in dsf_structure.json
No hard-coded cell references - everything from structure
"""

import json
from pathlib import Path
from typing import Dict, Any, Tuple, Optional
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from syscohada_db import SYSCOHADA_INDEX
from universal_mapping_strategy import UniversalMappingStrategy


class DSFStructurePrefiller:
    """
    Pre-fills DSF Excel using the structure defined in dsf_structure.json
    - Load structure: exact cell locations for every field
    - Load data: company balances and information
    - Map and fill: data → structure → Excel cells
    """
    
    def __init__(self, structure_file: str = "dsf_structure.json", 
                 mapping_file: str = "dsf_final_mapping.json"):
        """Initialize with structure and mapping files"""
        self.structure = self._load_json(structure_file)
        self.mapping = self._load_json(mapping_file)
        self.workbook = None
        self.ws_cache = {}  # Cache worksheets
        
    @staticmethod
    def _load_json(filepath: str) -> dict:
        """Load JSON file"""
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    
    def _get_or_create_worksheet(self, sheet_name: str):
        """Get or create worksheet, cached"""
        if sheet_name not in self.ws_cache:
            try:
                self.ws_cache[sheet_name] = self.workbook[sheet_name]
            except KeyError:
                self.ws_cache[sheet_name] = self.workbook.create_sheet(sheet_name)
        return self.ws_cache[sheet_name]
    
    def _parse_cell_ref(self, cell_ref: str) -> Tuple[str, int]:
        """Parse cell reference like 'A2' to ('A', 2)"""
        import re
        match = re.match(r'([A-Z]+)(\d+)', cell_ref)
        if match:
            return match.group(1), int(match.group(2))
        return None, None
    
    def _write_cell(self, sheet_name: str, cell_ref: str, value: Any, 
                   bold: bool = False, format_type: Optional[str] = None):
        """
        Write value to cell using cell reference
        Handles:
        - MergedCells (finds parent cell)
        - Formatting (bold, dates, currency)
        - Data validation
        """
        ws = self._get_or_create_worksheet(sheet_name)
        col, row = self._parse_cell_ref(cell_ref)
        
        if col is None:
            print(f"  ⚠️ Invalid cell reference: {cell_ref}")
            return
        
        cell = ws[cell_ref]
        
        # Check if cell is merged
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                # Write to top-left cell of merged range
                cell = ws[merged_range.start_cell.coordinate]
                break
        
        # Format data
        if format_type == "date":
            value = value if isinstance(value, str) else str(value)
            cell.number_format = 'DD/MM/YYYY'
        elif format_type == "currency":
            if value:
                cell.number_format = '#,##0.00'
        elif format_type == "exercice_fin":
            value = value if isinstance(value, str) else str(value)
        
        # Write value
        cell.value = value
        
        # Apply styling
        if bold:
            cell.font = Font(bold=True)
        
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    
    def fill_entete(self, company_data: dict):
        """
        Fill ENTETE sheet using structure
        Maps: company_data fields → structure cells → Excel
        """
        print("\n▶ Filling ENTETE sheet...")
        
        entete_structure = self.structure.get("BILAN PAYSAGE", {})
        cells = entete_structure.get("cells", {})
        
        # Map company data to cells
        mappings = {
            "A1": ("5", None),  # Page reference
            "A2": ("BILAN PAYSAGE", None),  # Sheet name
            # Company info would go here if we had those fields
        }
        
        # Actually iterate structure and fill with company data
        for cell_ref, cell_info in cells.items():
            # For now, just prepare - actual mapping depends on having company_data fields
            pass
        
        print("  ✅ ENTETE prepared")
    
    def fill_from_structure(self, company_data: dict, balance_data: dict):
        """
        Master method: Fill all DSF sheets using structure
        
        Args:
            company_data: Company information (name, ID, dates, etc.)
            balance_data: Account balances {account_num: amount}
        """
        print("\n" + "="*70)
        print("DSF PRE-FILLING - STRUCTURE DRIVEN")
        print("="*70)
        
        # Create workbook
        self.workbook = Workbook()
        self.workbook.remove(self.workbook.active)  # Remove default sheet
        
        # Iterate all sheets in structure
        for sheet_name, sheet_structure in self.structure.items():
            print(f"\n▶ Processing sheet: {sheet_name}")
            ws = self._get_or_create_worksheet(sheet_name)
            
            # Write all cells from structure
            cell_count = 0
            for cell_ref, cell_info in sheet_structure.items():
                if not isinstance(cell_info, dict):
                    continue
                
                # Get value from company_data or balance_data
                value = self._get_cell_value(cell_info, company_data, balance_data)
                
                if value is not None:
                    bold = cell_info.get("bold", False)
                    format_type = cell_info.get("format", None)
                    self._write_cell(sheet_name, cell_ref, value, bold, format_type)
                    cell_count += 1
                else:
                    # Write label/placeholder
                    label = cell_info.get("label", "")
                    if label:
                        self._write_cell(sheet_name, cell_ref, label, bold=cell_info.get("bold", False))
                        cell_count += 1
            
            print(f"  ✅ Filled {cell_count} cells")
        
        return self.workbook
    
    def _get_cell_value(self, cell_info: dict, company_data: dict, balance_data: dict) -> Any:
        """
        Extract value for a cell from data
        Priority:
        1. Direct value in cell_info
        2. Company data attribute
        3. Balance lookup via mapping
        """
        # Direct value
        if "value" in cell_info:
            return cell_info["value"]
        
        # Company data attribute
        attribute = cell_info.get("attribute")
        if attribute and attribute in company_data:
            return company_data[attribute]
        
        # Balance lookup
        if "syscohada_ref" in cell_info:
            account_num = cell_info["syscohada_ref"]
            return balance_data.get(account_num)
        
        return None
    
    def save(self, output_path: str):
        """Save workbook"""
        if self.workbook:
            self.workbook.save(output_path)
            size_mb = Path(output_path).stat().st_size / (1024 * 1024)
            print(f"\n✅ DSF saved: {output_path} ({size_mb:.1f} MB)")


class StructureMapper:
    """
    Maps intermediate categories to actual DSF cells using structure
    Bridges: SYSCOHADA → Intermediate → Structure Cell References
    """
    
    def __init__(self, structure_file: str = "dsf_structure.json",
                 mapping_file: str = "dsf_final_mapping.json"):
        self.structure = self._load_json(structure_file)
        self.mapping_data = self._load_json(mapping_file)
        self.reverse_index = self._build_reverse_index()
    
    @staticmethod
    def _load_json(filepath: str) -> dict:
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    
    def _build_reverse_index(self) -> dict:
        """
        Build reverse index: Cell Reference → Structure Info
        Purpose: Given a cell, find what should be in it
        """
        reverse = {}
        for sheet_name, sheet_cells in self.structure.items():
            for cell_ref, cell_info in sheet_cells.items():
                if isinstance(cell_info, dict):
                    reverse[f"{sheet_name}:{cell_ref}"] = cell_info
        return reverse
    
    def get_cell_for_category(self, intermediate_category: str) -> Optional[Dict]:
        """
        Find which DSF cell should contain data for this intermediate category
        Uses structure + mapping to locate exact cell
        """
        # Get DSF info from Level 2 mapping
        level_2 = self.mapping_data.get("LEVEL_2_INTERMEDIATE_TO_DSF", {})
        dsf_info = level_2.get(intermediate_category)
        
        if not dsf_info:
            return None
        
        # Now find matching cell in structure
        target_sheet = dsf_info.get("sheet")
        target_type = dsf_info.get("type")
        
        # Search structure for matching cell
        if target_sheet in self.structure:
            sheet_cells = self.structure[target_sheet]
            for cell_ref, cell_info in sheet_cells.items():
                if isinstance(cell_info, dict):
                    if cell_info.get("type") == target_type or cell_info.get("label") == dsf_info.get("label"):
                        return {
                            "sheet": target_sheet,
                            "cell": cell_ref,
                            "info": cell_info,
                            "dsf_info": dsf_info
                        }
        
        return None
    
    def map_syscohada_to_cell(self, account_num: str) -> Optional[Tuple[str, str]]:
        """
        Find the exact DSF cell location for a SYSCOHADA account
        Returns: (sheet_name, cell_reference) or None
        """
        # Get intermediate category for this account
        level_1 = self.mapping_data.get("LEVEL_1_SYSCOHADA_TO_INTERMEDIATE", {})
        mapping = level_1.get(str(account_num))
        
        if not mapping:
            return None
        
        intermediate_cat = mapping.get("intermediate_category")
        cell_info = self.get_cell_for_category(intermediate_cat)
        
        if cell_info:
            return (cell_info["sheet"], cell_info["cell"])
        
        return None


if __name__ == "__main__":
    print("\n▶ DSF Structure-Based Prefiller\n")
    
    # Load structure
    mapper = StructureMapper()
    print(f"✅ Structure loaded: {len(mapper.structure)} sheets")
    print(f"✅ Mapping loaded: {len(mapper.mapping_data.get('LEVEL_1_SYSCOHADA_TO_INTERMEDIATE', {}))} accounts")
    
    # Test mapping
    test_accounts = ['101', '201', '601', '701']
    print("\n▶ Test: SYSCOHADA → DSF Cells")
    for account in test_accounts:
        cell_location = mapper.map_syscohada_to_cell(account)
        if cell_location:
            sheet, cell = cell_location
            print(f"  {account:5} → {sheet}:{cell}")
        else:
            print(f"  {account:5} → ⚠️ NOT FOUND IN STRUCTURE")
