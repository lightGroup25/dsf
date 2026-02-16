"""
DSF Prefiller - SIMPLE & SAFE VERSION
- Écrit SEULEMENT les valeurs
- NE TOUCHE jamais aux styles/format/merged cells
- Préserve 100% du design du template
"""

import json
import logging
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from dsf_general_info import DSF_InfosGenerales, format_currency, format_date

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


class DSFPrefillerSimple:
    """
    Version SIMPLE du pré-remplisseur:
    - Charge template
    -Écrit SEULEMENT les valeurs (pas de style)
    - Sauvegarde
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
        """Charge le template Excel"""
        logger.info(f"Loading template: {self.template_path.name}")
        self.wb = load_workbook(str(self.template_path))
        logger.info(f"  ✓ Sheets: {len(self.wb.sheetnames)} sheets loaded")
        return self.wb
    
    def load_mapping(self):
        """Charge le mapping JSON"""
        logger.info(f"Loading mapping: {self.mapping_path.name}")
        with open(self.mapping_path, 'r', encoding='utf-8') as f:
            self.mapping = json.load(f)
        logger.info(f"  ✓ Mapping loaded")
        return self.mapping
    
    def _write_cell_safe(self, ws, cell_ref: str, value):
        """
        Écrit une valeur dans une cellule SANS TOUCHER au style
        Gère les merged cells correctement
        """
        try:
            cell = ws[cell_ref]
            
            # Si la cellule est dans une merged cell, écrire dans la cellule parent (top-left)
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Utiliser la cellule de base de la merged cell
                    cell = ws[merged_range.start_cell.coordinate]
                    break
            
            # Écrire uniquement la valeur - NE PAS TOUCHER AU STYLE
            cell.value = value
            return True
            
        except Exception as e:
            logger.warning(f"    ⚠ {cell_ref}: {str(e)[:80]}")
            return False
    
    def prefill_entete(self, company_info: DSF_InfosGenerales):
        """Remplit la page d'en-tête ENTETE"""
        logger.info("\nFilling ENTETE sheet...")
        
        if "ENTETE" not in self.wb.sheetnames:
            logger.warning("  ⚠ ENTETE sheet not found")
            return 0
        
        ws = self.wb["ENTETE"]
        
        # Mapping directe des cellules et valeurs
        data_to_write = {
            "B2": company_info.denomination_sociale or "---",
            "B3": company_info.sigle_usuel or "---",
            "B4": (company_info.adresse_complete or "---"),
            "B5": company_info.num_identification_fiscale or "---",
            "B6": company_info.systeme_comptable or "Système Normal",
            "B7": format_date(company_info.exercice_fin) if company_info.exercice_fin else "---",
            "B9": company_info.centre_depot or "---",
            "B10": company_info.ministere or "---",
            "B11": company_info.direction_generale or "---",
        }
        
        count = 0
        for cell_ref, value in data_to_write.items():
            if self._write_cell_safe(ws, cell_ref, value):
                count += 1
                logger.info(f"  ✓ {cell_ref} = {str(value)[:40]}")
        
        logger.info(f"  Total: {count} cells filled")
        return count
    
    def prefill_entete_simple_version(self, company_info: DSF_InfosGenerales):
        """Version alternative for ENTETE si les références sont différentes"""
        logger.info("\nFilling ENTETE sheet (alternative mapping)...")
        
        if "ENTETE" not in self.wb.sheetnames:
            logger.warning("  ⚠ ENTETE sheet not found")
            return 0
        
        ws = self.wb["ENTETE"]
        
        # Essayer plusieurs emplacements possibles
        possible_locations = {
            "denomination_sociale": ["A2", "B2", "C2", "D2"],
            "sigle_usuel": ["A3", "B3", "C3"],
            "num_identification_fiscale": ["A5", "B5", "C5"],
        }
        
        # Trouver les bonnes cellules (chercher le label)
        count = 0
        for key, cell_list in possible_locations.items():
            value = getattr(company_info, key, None)
            if value:
                for cell_ref in cell_list:
                    try:
                        if self._write_cell_safe(ws, cell_ref, value):
                            logger.info(f"  ✓ {cell_ref} = {str(value)[:40]}")
                            count += 1
                            break
                    except:
                        pass
        
        return count
    
    def prefill_info_generales(self, company_info: DSF_InfosGenerales):
        """Remplit la page 'INFORMATIONS GENERALES' si elle existe"""
        logger.info("\nFilling INFORMATIONS GENERALES sheet...")
        
        if "INFORMATIONS GENERALES" not in self.wb.sheetnames:
            logger.warning("  ⚠ INFORMATIONS GENERALES sheet not found")
            return 0
        
        ws = self.wb["INFORMATIONS GENERALES"]
        
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
            try:
                if self._write_cell_safe(ws, cell_ref, value):
                    count += 1
                    logger.info(f"  ✓ {cell_ref} = {str(value)[:40]}")
            except:
                pass
        
        if count > 0:
            logger.info(f"  Total: {count} cells filled")
        return count
    
    def save(self, output_path: str):
        """Sauvegarde le fichier rempli"""
        output_path = Path(output_path)
        logger.info(f"\nSaving to: {output_path.name}")
        self.wb.save(str(output_path))
        logger.info(f"  ✓ File saved ({output_path.stat().st_size / 1024:.1f} KB)")
        return output_path


def main():
    """Test avec les données GULFCAM"""
    
    print("\n" + "="*70)
    print("DSF SIMPLE PREFILLER - DESIGN PRESERVATION")
    print("="*70)
    
    prefiller = DSFPrefillerSimple(
        template_path="DSF Normal standard.xlsx",
        mapping_path="dsf_prefill_mapping.json"
    )
    
    # Charger le template
    prefiller.load_template()
    
    # Charger le mapping
    prefiller.load_mapping()
    
    # Créer des données de test GULFCAM
    company_info = DSF_InfosGenerales(
        denomination_sociale="GULFCAM S.A.S.",
        sigle_usuel="GULFCAM",
        num_identification_fiscale="R.C. 1998/SDE/CM",
        adresse_complete="Douala - Littoral - Cameroun",
        exercice_debut=datetime(2024, 1, 1).date(),
        exercice_fin=datetime(2024, 12, 31).date(),
        date_arrete_comptes=datetime(2024, 12, 31).date(),
        systeme_comptable="Système Normal",
        forme_juridique="S.A.S.",
        registre_fiscal="SCIMPEX",
        centre_depot="Douala",
        ministere="MINFI",
        direction_generale="DGI"
    )
    
    # Remplir
    prefiller.prefill_entete(company_info)
    prefiller.prefill_entete_simple_version(company_info)
    prefiller.prefill_info_generales(company_info)
    
    # Sauvegarder
    output = prefiller.save("DSF_GULFCAM_FINAL.xlsx")
    
    print("\n" + "="*70)
    print("✅ COMPLETE!")
    print("="*70)
    print(f"\nOutput: {output}")
    print("\n📊 Check file to verify:")
    print("  ✓ Data filled correctly")
    print("  ✓ Design completely preserved")
    print("  ✓ Styles intact")
    print("  ✓ Layout unchanged\n")


if __name__ == "__main__":
    main()
