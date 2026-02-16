"""
DSF Prefiller - FIXED VERSION FOR MERGED CELLS
- Detects that data area is one giant merged cell (A2:I43)
- Writes to the master cell (A2)
- Formats data as multi-line text
"""

import json
import logging
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from dsf_general_info import DSF_InfosGenerales, format_currency, format_date

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


class DSFPrefillerMergedCell:
    """
    Handles merged cell structure in ENTETE sheet.
    The entire data area (A2:I43) is ONE merged cell.
    We must write to A2 (the master cell).
    """
    
    def __init__(self, template_path: str = "DSF Normal standard.xlsx", 
                 mapping_path: str = "dsf_prefill_mapping.json"):
        self.template_path = Path(template_path)
        self.mapping_path = Path(mapping_path)
        self.wb = None
        self.mapping = None
        
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
        logger.info(f"  ✓ Mapping loaded")
        return self.mapping
    
    def _find_master_cell(self, ws, cell_ref: str):
        """
        Find the master cell for a given cell reference.
        If the cell is in a merged range, return the master (top-left) cell.
        Otherwise, return the cell itself.
        """
        cell = ws[cell_ref]
        
        # Check if cell is in any merged range
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                # Return the master cell (start of merged range)
                return ws[merged_range.start_cell.coordinate]
        
        # Not in a merged range, return as is
        return cell
    
    def _write_to_master(self, ws, cell_ref: str, value):
        """
        Write to the master cell of a merged range.
        This bypasses the read-only MergedCell restriction.
        """
        try:
            master_cell = self._find_master_cell(ws, cell_ref)
            master_cell.value = value
            logger.info(f"    ✓ {cell_ref} → {master_cell.coordinate} = {str(value)[:50]}")
            return True
        except Exception as e:
            logger.warning(f"    ✗ {cell_ref}: {str(e)[:80]}")
            return False
    
    def prefill_entete(self, company_info: DSF_InfosGenerales):
        """
        Fill the ENTETE sheet.
        
        NOTE: The entire data area (A2:I43) is merged into ONE cell.
        We can only write to A2 (the master cell).
        We'll format data as multi-line text within that single cell.
        """
        logger.info("\n📝 Filling ENTETE sheet (merged cell structure)...")
        
        if "ENTETE" not in self.wb.sheetnames:
            logger.warning("  ✗ ENTETE sheet not found")
            return 0
        
        ws = self.wb["ENTETE"]
        
        # Check merged cells
        logger.info(f"  Merged ranges: {len(list(ws.merged_cells.ranges))}")
        for merged_range in ws.merged_cells.ranges:
            logger.info(f"    • {merged_range} (master: {merged_range.start_cell.coordinate})")
        
        # Since A2:I43 is a single merged cell, we write everything to A2
        # Format as multi-line text
        entete_data = (
            f"Dénomination sociale: {company_info.denomination_sociale or '---'}\n"
            f"Sigle usuel: {company_info.sigle_usuel or '---'}\n"
            f"Adresse: {company_info.adresse_complete or '---'}\n"
            f"N° identification fiscale: {company_info.num_identification_fiscale or '---'}\n"
            f"Système comptable: {company_info.systeme_comptable or 'Système Normal'}\n"
            f"Exercice au: {format_date(company_info.exercice_fin) if company_info.exercice_fin else '---'}\n"
            f"Centre de dépôt: {company_info.centre_depot or '---'}\n"
            f"Ministère: {company_info.ministere or '---'}\n"
            f"Direction générale: {company_info.direction_generale or '---'}"
        )
        
        # Write to master cell
        success = self._write_to_master(ws, "A2", entete_data)
        
        if success:
            logger.info(f"  ✓ ENTETE data written to A2 (master cell)")
            return 1
        else:
            logger.warning(f"  ✗ Failed to write ENTETE data")
            return 0
    
    def prefill_info_generales(self, company_info: DSF_InfosGenerales):
        """Fill the INFORMATIONS GENERALES sheet if possible"""
        logger.info("\n📝 Filling INFORMATIONS GENERALES sheet...")
        
        if "INFORMATIONS GENERALES" not in self.wb.sheetnames:
            logger.warning("  ✗ INFORMATIONS GENERALES sheet not found")
            return 0
        
        ws = self.wb["INFORMATIONS GENERALES"]
        
        # Try to write to individual cells
        # Map the data
        data_to_write = {
            "C2": company_info.denomination_sociale or "---",
            "C3": company_info.sigle_usuel or "---",
            "C4": company_info.num_identification_fiscale or "---",
            "C5": company_info.adresse_complete or "---",
            "C6": company_info.systeme_comptable or "Système Normal",
            "C7": format_date(company_info.exercice_fin) if company_info.exercice_fin else "---",
        }
        
        count = 0
        for cell_ref, value in data_to_write.items():
            if self._write_to_master(ws, cell_ref, value):
                count += 1
        
        logger.info(f"  ✓ {count} cells filled in INFORMATIONS GENERALES")
        return count
    
    def save(self, output_path: str = "DSF_GULFCAM_FINAL.xlsx"):
        """Save the workbook"""
        logger.info(f"\n💾 Saving to: {output_path}")
        self.wb.save(output_path)
        
        file_size = Path(output_path).stat().st_size / 1024
        logger.info(f"  ✓ File saved ({file_size:.1f} KB)")


def main():
    """Main execution"""
    
    logger.info("\n" + "="*70)
    logger.info("DSF PREFILLER - MERGED CELL FIX")
    logger.info("="*70)
    
    # Load template and mapping
    prefiller = DSFPrefillerMergedCell()
    prefiller.load_template()
    prefiller.load_mapping()
    
    # Create test company data
    from datetime import date
    logger.info("\n📋 Creating test company data...")
    company_info = DSF_InfosGenerales(
        denomination_sociale="GULFCAM S.A.S.",
        sigle_usuel="GULFCAM",
        adresse_complete="Douala - Littoral - Cameroun",
        num_identification_fiscale="R.C. 1998/SDE/CM",
        exercice_debut=date(2024, 1, 1),
        exercice_fin=date(2024, 12, 31),
        date_arrete_comptes=date(2024, 12, 31),
        forme_juridique="08",
        registre_fiscal="1",
        systeme_comptable="Système Normal",
        centre_depot="Douala",
        ministere="MINFI",
        direction_generale="DGI"
    )
    logger.info("  ✓ Test data created")
    
    # Fill sheets
    entete_count = prefiller.prefill_entete(company_info)
    info_count = prefiller.prefill_info_generales(company_info)
    
    # Save
    prefiller.save("DSF_GULFCAM_MERGED_FIX.xlsx")
    
    logger.info("\n" + "="*70)
    logger.info(f"✅ PREFILLING COMPLETE")
    logger.info(f"   ENTETE: {entete_count} cells")
    logger.info(f"   INFORMATIONS GENERALES: {info_count} cells")
    logger.info("="*70 + "\n")


if __name__ == "__main__":
    main()
