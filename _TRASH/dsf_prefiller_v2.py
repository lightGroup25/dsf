"""
DSF PREFILLER - REVISED ARCHITECTURE v2
======================================

Maps data through 3 layers:
1. SYSCOHADA accounts → Intermediate categories (via dsf_final_mapping.json)
2. Intermediate categories → DSF cell locations (via category_to_dsf_mapping.py)
3. Fills Excel cells with aggregated balances

Uses dsf_structure.json as reference for cell locations
"""

import json
import re
from pathlib import Path
from typing import Dict, Tuple, Optional, List
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from syscohada_db import SYSCOHADA_INDEX
from universal_mapping_strategy import UniversalMappingStrategy
from category_to_dsf_mapping import CATEGORY_TO_DSF_LINES


class DSFPrefiller:
    """
    Pre-fills DSF using structure-driven mapping
    Data flow: Balance → SYSCOHADA → Category → DSF Cell → Excel
    """
    
    def __init__(self, structure_file: str = "dsf_structure.json",
                 mapping_file: str = "dsf_final_mapping.json",
                 template_file: Optional[str] = None):
        """Initialize prefiller with structure and mappings"""
        self.structure = self._load_json(structure_file)
        self.mapping = self._load_json(mapping_file)
        self.category_to_dsf = CATEGORY_TO_DSF_LINES
        
        # Load or create workbook
        if template_file and Path(template_file).exists():
            self.workbook = load_workbook(template_file)
        else:
            self.workbook = Workbook()
            self.workbook.remove(self.workbook.active)
    
    @staticmethod
    def _load_json(filepath: str) -> dict:
        """Load JSON file"""
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    
    def _get_worksheet(self, sheet_name: str):
        """Get or create worksheet"""
        if sheet_name in self.workbook.sheetnames:
            return self.workbook[sheet_name]
        else:
            return self.workbook.create_sheet(sheet_name)
    
    def _cell_ref_to_coords(self, cell_ref: str) -> Tuple[int, int]:
        """Convert cell ref like 'A2' to column (1) and row (2)"""
        match = re.match(r'([A-Z]+)(\d+)', cell_ref)
        if match:
            col_letters = match.group(1)
            row = int(match.group(2))
            # Convert column letters to number
            col = 0
            for char in col_letters:
                col = col * 26 + (ord(char) - ord('A') + 1)
            return row, col
        return None, None
    
    def prefill_from_balances(self, company_info: Dict, balances: Dict[str, float]):
        """
        Main method: Fill DSF from balance data
        
        Args:
            company_info: Company metadata {name, id, dates, etc}
            balances: {account_num: amount}
        """
        print("\n" + "="*80)
        print("DSF PREFILLER - STRUCTURE-DRIVEN v2")
        print("="*80)
        
        # Step 1: Aggregate balances by intermediate category
        print("\n▶ Step 1: Aggregating balances by category...")
        category_balances = self._aggregate_balances(balances)
        print(f"  ✅ Aggregated {len(category_balances)} categories")
        print(f"     Sample: {list(category_balances.items())[:3]}")
        
        # Step 2: Map categories to DSF cells
        print("\n▶ Step 2: Mapping categories to DSF cells...")
        cell_data = self._map_categories_to_cells(category_balances)
        print(f"  ✅ Mapped {len(cell_data)} cells")
        
        # Step 3: Fill structure
        print("\n▶ Step 3: Filling structure...")
        self._fill_structure(company_info)
        print(f"  ✅ Structure filled")
        
        # Step 4: Fill cells with balance data
        print("\n▶ Step 4: Writing balance data to cells...")
        cells_written = self._write_cell_data(cell_data)
        print(f"  ✅ Wrote {cells_written} cell values")
        
        print("\n" + "="*80)
        print("✅ PREFILL COMPLETE")
        print("="*80)
    
    def _aggregate_balances(self, balances: Dict[str, float]) -> Dict[str, float]:
        """
        Aggregate balances by intermediate category
        SYSCOHADA → Intermediate → Sum by category
        """
        level_1 = self.mapping["LEVEL_1_SYSCOHADA_TO_INTERMEDIATE"]
        category_totals = {}
        
        for account_num, amount in balances.items():
            account_str = str(account_num)
            if account_str in level_1:
                mapping = level_1[account_str]
                category = mapping["intermediate_category"]
                
                if category not in category_totals:
                    category_totals[category] = 0
                category_totals[category] += amount
        
        return category_totals
    
    def _map_categories_to_cells(self, category_balances: Dict[str, float]) -> Dict[str, Dict]:
        """
        Map each category to its DSF cell location and value
        Returns: {(sheet, cell): {value, type, ...}}
        """
        cell_data = {}
        
        for category, balance in category_balances.items():
            if category in self.category_to_dsf:
                dsf_line = self.category_to_dsf[category]
                sheet_name = dsf_line["sheet"]
                location = dsf_line["location"]
                
                # For now, use the first value column
                if "values_cols" in location and location["values_cols"]:
                    cell_col = location["values_cols"][0]  # Current year
                    cell_row = location["label_row"]
                    cell_ref = f"{cell_col}{cell_row}"
                    
                    key = (sheet_name, cell_ref)
                    cell_data[key] = {
                        "value": balance,
                        "category": category,
                        "type": dsf_line.get("type")
                    }
        
        return cell_data
    
    def _fill_structure(self, company_info: Dict):
        """
        Write the static structure from dsf_structure.json
        Fills labels, headers, company info, etc.
        """
        for sheet_name, cells in self.structure.items():
            ws = self._get_worksheet(sheet_name)
            
            for cell_ref, value in cells.items():
                if isinstance(value, str):
                    # Write to cell
                    try:
                        cell = ws[cell_ref]
                        cell.value = value
                        
                        # Style headers
                        if cell_ref.endswith(("8", "9", "10")):  # Header rows
                            cell.font = Font(bold=True, size=10)
                        elif value.isupper() and len(value) > 5:  # Section headers
                            cell.font = Font(bold=True)
                        
                        cell.alignment = Alignment(wrap_text=True, vertical='center')
                    except Exception as e:
                        pass  # Ignore cell write errors
    
    def _write_cell_data(self, cell_data: Dict) -> int:
        """Write balance values to cells"""
        written = 0
        
        for (sheet_name, cell_ref), data in cell_data.items():
            try:
                ws = self._get_worksheet(sheet_name)
                cell = ws[cell_ref]
                
                # Handle merged cells
                for merged_range in ws.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        cell = ws[merged_range.start_cell.coordinate]
                        break
                
                # Write value
                value = data["value"]
                cell.value = float(value) if value else 0
                
                # Format as currency
                cell.number_format = '#,##0.00' if isinstance(value, (int, float)) else '@'
                cell.alignment = Alignment(horizontal='right', wrap_text=True)
                
                written += 1
            except Exception as e:
                print(f"  ⚠️ Error writing {sheet_name}:{cell_ref}: {e}")
        
        return written
    
    def save(self, output_path: str):
        """Save workbook"""
        self.workbook.save(output_path)
        size_mb = Path(output_path).stat().st_size / (1024 * 1024)
        print(f"\n✅ DSF saved: {output_path}")
        print(f"   Size: {size_mb:.1f} MB")
        print(f"   Sheets: {len(self.workbook.sheetnames)}")


# ============================================================================
# EXAMPLE USAGE
# ============================================================================

if __name__ == "__main__":
    # Sample company data
    company_info = {
        "denomination_sociale": "GULFCAM SAS",
        "num_identification_fiscale": "123456789",
        "exercice_fin": "2024-12-31",
        "forme_juridique": "SARL"
    }
    
    # Sample balance data (simplified)
    sample_balances = {
        "101": 100000,      # Capital
        "1011": 50000,      # Capital account detail
        "201": 250000,      # Intangible assets
        "2011": 75000,      # Development costs
        "301": 150000,      # Inventory
        "401": 80000,       # Suppliers
        "601": 500000,      # Material purchases
        "701": 2000000,     # Sales
    }
    
    # Run prefiller
    prefiller = DSFPrefiller()
    prefiller.prefill_from_balances(company_info, sample_balances)
    prefiller.save("DSF_TEST_v2.xlsx")
    
    print("\n✅ Test complete!")
