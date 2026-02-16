"""
DSF PREFILLER - PRODUCTION VERSION
Handles merged cells correctly and fills from balance data
"""

import json
import logging
from pathlib import Path
from datetime import date
from openpyxl import load_workbook
from dsf_general_info import DSF_InfosGenerales, format_currency, format_date

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


class DSFPrefillerProduction:
    """
    Production prefiller that:
    - Correctly handles merged cells by routing to master cell
    - Writes ONLY values (preserves 100% of template design)
    - Fills company info in ENTETE sheet
    - Fills general info in INFORMATIONS GENERALES sheet
    - Fills balance data in subsequent sheets (R1, R2, R3, etc.)
    """
    
    def __init__(self, template_path: str = "DSF Normal standard.xlsx", 
                 mapping_path: str = "dsf_prefill_mapping.json"):
        self.template_path = Path(template_path)
        self.mapping_path = Path(mapping_path)
        self.wb = None
        self.mapping = {}
        self.sheets_written = {}
        
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {self.template_path}")
        if not self.mapping_path.exists():
            raise FileNotFoundError(f"Mapping not found: {self.mapping_path}")
    
    def load_template(self):
        """Load the Excel template"""
        logger.info(f"Loading template: {self.template_path.name}")
        self.wb = load_workbook(str(self.template_path))
        logger.info(f"  ✓ Sheets: {len(self.wb.sheetnames)} sheets loaded")
        return self.wb
    
    def load_mapping(self):
        """Load the JSON mapping"""
        logger.info(f"Loading mapping: {self.mapping_path.name}")
        with open(self.mapping_path, 'r', encoding='utf-8') as f:
            self.mapping = json.load(f)
        logger.info(f"  ✓ Mapping loaded for {len(self.mapping.get('ENTETE', {}))} ENTETE cells")
        return self.mapping
    
    def _find_master_cell(self, ws, cell_ref: str):
        """
        Find the master cell for a given cell reference.
        If the cell is in a merged range, return the master (top-left) cell.
        Otherwise, return the cell itself.
        """
        try:
            cell = ws[cell_ref]
            
            # Check if cell is in any merged range
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Return the master cell (start of merged range)
                    return ws[merged_range.start_cell.coordinate]
            
            # Not in a merged range, return as is
            return cell
        except Exception as e:
            logger.warning(f"    ✗ Error finding master cell for {cell_ref}: {str(e)[:60]}")
            return None
    
    def write_cell(self, sheet_name: str, cell_ref: str, value):
        """
        Write a value to a cell in a specific sheet.
        Handles merged cells automatically.
        Returns True if successful, False otherwise.
        """
        if sheet_name not in self.wb.sheetnames:
            return False
        
        ws = self.wb[sheet_name]
        
        try:
            master_cell = self._find_master_cell(ws, cell_ref)
            if master_cell is None:
                return False
            
            master_cell.value = value
            return True
        except Exception as e:
            logger.warning(f"    ✗ {cell_ref} in {sheet_name}: {str(e)[:60]}")
            return False
    
    def fill_entete(self, company_info: DSF_InfosGenerales):
        """
        Fill the ENTETE (header) sheet with company information.
        Since the data area is typically a merged cell, we write multi-line text.
        """
        logger.info("\n📝 Filling ENTETE sheet...")
        
        if "ENTETE" not in self.wb.sheetnames:
            logger.warning("  ✗ ENTETE sheet not found")
            return 0
        
        ws = self.wb["ENTETE"]
        
        # Format entete data as multi-line text in the merged cell
        entete_lines = [
            f"Dénomination sociale: {company_info.denomination_sociale or '---'}",
            f"Sigle usuel: {company_info.sigle_usuel or '---'}",
            f"Adresse: {company_info.adresse_complete or '---'}",
            f"N° identification fiscale: {company_info.num_identification_fiscale or '---'}",
            f"Système comptable: {company_info.systeme_comptable or 'Système Normal'}",
            f"Exercice au: {format_date(company_info.exercice_fin) if company_info.exercice_fin else '---'}",
            f"Centre de dépôt: {company_info.centre_depot or '---'}",
            f"Ministère: {company_info.ministere or '---'}",
            f"Direction générale: {company_info.direction_generale or '---'}"
        ]
        
        entete_data = "\n".join(entete_lines)
        
        # Try to write to A2 (typically the master cell of the merged area)
        success = self.write_cell("ENTETE", "A2", entete_data)
        
        if success:
            logger.info(f"  ✓ ENTETE data written to A2")
            return 1
        else:
            logger.warning(f"  ✗ Failed to write ENTETE data")
            return 0
    
    def fill_info_generales(self, company_info: DSF_InfosGenerales):
        """
        Fill the INFORMATIONS GENERALES sheet with company details.
        """
        logger.info("\n📋 Filling INFORMATIONS GENERALES sheet...")
        
        if "INFORMATIONS GENERALES" not in self.wb.sheetnames:
            logger.warning("  ✗ INFORMATIONS GENERALES sheet not found")
            return 0
        
        ws = self.wb["INFORMATIONS GENERALES"]
        
        # Map data to cells (these may be merged)
        data_map = {
            "C2": company_info.denomination_sociale or "---",
            "C3": company_info.sigle_usuel or "---",
            "C4": company_info.num_identification_fiscale or "---",
            "C5": company_info.adresse_complete or "---",
            "C6": company_info.systeme_comptable or "Système Normal",
            "C7": format_date(company_info.exercice_fin) if company_info.exercice_fin else "---",
        }
        
        count = 0
        for cell_ref, value in data_map.items():
            if self.write_cell("INFORMATIONS GENERALES", cell_ref, value):
                count += 1
        
        logger.info(f"  ✓ {count} cells filled")
        return count
    
    def fill_sheet_from_mapping(self, sheet_name: str, data_dict: dict):
        """
        Fill a sheet using a mapping of cell references to values.
        
        Usage:
            balance_data = {'B2': 100000, 'B3': 200000, ...}
            prefiller.fill_sheet_from_mapping('R1', balance_data)
        """
        if sheet_name not in self.wb.sheetnames:
            logger.warning(f"  ✗ Sheet {sheet_name} not found")
            return 0
        
        count = 0
        for cell_ref, value in data_dict.items():
            if self.write_cell(sheet_name, cell_ref, value):
                count += 1
        
        return count
    
    def save(self, output_path: str = "DSF_OUTPUT.xlsx"):
        """
        Save the workbook.
        """
        logger.info(f"\n💾 Saving to: {output_path}")
        try:
            self.wb.save(output_path)
            file_size = Path(output_path).stat().st_size / 1024
            logger.info(f"  ✓ File saved ({file_size:.1f} KB)")
            return True
        except Exception as e:
            logger.error(f"  ✗ Failed to save: {str(e)}")
            return False


def create_gulfcam_company_info() -> DSF_InfosGenerales:
    """Create company information for GULFCAM"""
    return DSF_InfosGenerales(
        denomination_sociale="GULFCAM S.A.S.",
        sigle_usuel="GULFCAM",
        adresse_complete="CENTRE DES AFFAIRES MARITIMES, B.P 3876 DOUALA",
        num_identification_fiscale="M050900027774W",
        exercice_debut=date(2024, 1, 1),
        exercice_fin=date(2024, 12, 31),
        date_arrete_comptes=date(2024, 12, 31),
        forme_juridique="08",
        registre_fiscal="1",
        systeme_comptable="Système Normal",
        centre_depot="DIRECTION DES GRANDES ENTREPRISES",
        ministere="MINISTERE DES FINANCES",
        direction_generale="DIRECTION GENERALE DES IMPOTS"
    )


def main_demo():
    """
    Demo run filling headers and company info only.
    For full balance data, use the balance_transformer output.
    """
    
    logger.info("\n" + "="*70)
    logger.info("DSF PREFILLER - PRODUCTION VERSION")
    logger.info("="*70)
    
    # Initialize prefiller
    prefiller = DSFPrefillerProduction()
    prefiller.load_template()
    prefiller.load_mapping()
    
    # Create company info
    logger.info("\n📋 Preparing company information...")
    company_info = create_gulfcam_company_info()
    logger.info("  ✓ Company info created")
    
    # Fill sheets
    entete_count = prefiller.fill_entete(company_info)
    info_count = prefiller.fill_info_generales(company_info)
    
    # Save
    prefiller.save("DSF_GULFCAM_FINAL.xlsx")
    
    logger.info("\n" + "="*70)
    logger.info(f"✅ PREFILLING COMPLETE")
    logger.info(f"   ENTETE: {entete_count} entries")
    logger.info(f"   INFORMATIONS GENERALES: {info_count} entries")
    logger.info("="*70 + "\n")


if __name__ == "__main__":
    main_demo()
