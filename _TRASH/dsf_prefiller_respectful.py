"""
DSF Prefiller - Respecte le design du template Excel
- Charge le template DSF Normal standard.xlsx
- Écrit UNIQUEMENT dans les cellules de données
- Préserve TOUS les styles, formats, cellules fusionnées, formules
"""

import json
import logging
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border
from dsf_general_info import DSF_InfosGenerales, format_currency, format_date

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


class DSFPrefillerRespectful:
    """
    Pré-remplisseur qui respecte le design du template
    - Charge le template depuis DSF Normal standard.xlsx
    - Écrit les données sans modifier les styles
    - Préserve les merged cells
    """
    
    def __init__(self, template_path: str = "DSF Normal standard.xlsx", 
                 mapping_path: str = "dsf_prefill_mapping.json"):
        self.template_path = Path(template_path)
        self.mapping_path = Path(mapping_path)
        
        if not self.template_path.exists():
            raise FileNotFoundError(f"❌ Template not found: {self.template_path}")
        if not self.mapping_path.exists():
            raise FileNotFoundError(f"❌ Mapping not found: {self.mapping_path}")
        
        self.wb = None
        self.mapping = None
        self.prefill_data = None
        
    def load_template(self):
        """Charge le template Excel en préservant formules et styles"""
        logger.info(f"📂 Loading template: {self.template_path}")
        self.wb = load_workbook(str(self.template_path))
        logger.info(f"   ✓ Sheets: {self.wb.sheetnames}")
        return self.wb
    
    def load_mapping(self):
        """Charge le mapping JSON"""
        logger.info(f"📋 Loading mapping: {self.mapping_path}")
        with open(self.mapping_path, 'r', encoding='utf-8') as f:
            self.mapping = json.load(f)
        logger.info(f"   ✓ Sheets in mapping: {list(self.mapping.keys())}")
        return self.mapping
    
    def _get_cell_style(self, cell):
        """Extrait le style d'une cellule pour le réappliquer"""
        return {
            "font": Font(
                name=cell.font.name,
                size=cell.font.size,
                bold=cell.font.bold,
                italic=cell.font.italic,
                color=cell.font.color
            ) if cell.font else None,
            "fill": PatternFill(
                start_color=cell.fill.start_color,
                end_color=cell.fill.end_color,
                fill_type=cell.fill.fill_type
            ) if cell.fill else None,
            "alignment": Alignment(
                horizontal=cell.alignment.horizontal,
                vertical=cell.alignment.vertical,
                wrap_text=cell.alignment.wrap_text
            ) if cell.alignment else None,
            "border": cell.border,
            "number_format": cell.number_format
        }
    
    def _apply_cell_style(self, cell, style_dict):
        """Applique un style à une cellule"""
        if style_dict.get("font"):
            cell.font = style_dict["font"]
        if style_dict.get("fill"):
            cell.fill = style_dict["fill"]
        if style_dict.get("alignment"):
            cell.alignment = style_dict["alignment"]
        if style_dict.get("border"):
            cell.border = style_dict["border"]
        if style_dict.get("number_format"):
            cell.number_format = style_dict["number_format"]
    
    def _write_cell_safe(self, ws, cell_ref: str, value, preserve_style=True):
        """
        Écrit dans une cellule en préservant le style
        Gère les cellules fusionnées correctement
        """
        try:
            cell = ws[cell_ref]
            
            # Vérifier si la cellule est dans une merged cell
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Écrire dans la cellule top-left de la merged cell
                    cell_ref = merged_range.start_cell.coordinate
                    cell = ws[cell_ref]
                    break
            
            # Copier le style avant modification
            if preserve_style:
                original_style = self._get_cell_style(cell)
            
            # Écrire la valeur
            cell.value = value
            
            # Restaurer le style
            if preserve_style:
                self._apply_cell_style(cell, original_style)
            
            logger.debug(f"   ✓ {cell_ref}: {value}")
            return True
            
        except Exception as e:
            logger.warning(f"   ⚠ Could not write to {cell_ref}: {e}")
            return False
    
    def prefill_entete(self, company_info: DSF_InfosGenerales):
        """Remplit la page d'en-tête (ENTETE)"""
        if "ENTETE" not in self.wb.sheetnames:
            logger.warning("⚠  Sheet ENTETE not found")
            return False
        
        ws = self.wb["ENTETE"]
        logger.info("📝 Filling ENTETE (header) sheet...")
        
        mapping_entete = self.mapping.get("ENTETE", {}).get("cells", {})
        
        data_map = {
            "denomination_sociale": company_info.denomination_sociale,
            "sigle_usuel": company_info.sigle_usuel or "",
            "adresse_complete": company_info.adresse_complete or "",
            "num_identification_fiscale": company_info.num_identification_fiscale or "",
            "systeme_comptable": company_info.systeme_comptable or "Système Normal",
            "systeme_normal": format_date(company_info.exercice_fin) if company_info.exercice_fin else "",
            "centre_depot": company_info.centre_depot or "",
            "ministere": company_info.ministere or "",
            "direction_generale": company_info.direction_generale or "",
        }
        
        count = 0
        for cell_ref, mapping_info in mapping_entete.items():
            attribute = mapping_info.get("attribute")
            if attribute in data_map:
                value = data_map[attribute]
                if self._write_cell_safe(ws, cell_ref, value):
                    count += 1
        
        logger.info(f"   ✓ {count} cells filled in ENTETE")
        return True
    
    def prefill_r1(self, company_info: DSF_InfosGenerales):
        """Remplit la page R1 (Informations exercice)"""
        if "R1" not in self.wb.sheetnames:
            logger.warning("⚠  Sheet R1 not found")
            return False
        
        ws = self.wb["R1"]
        logger.info("📝 Filling R1 (Informations exercice) sheet...")
        
        mapping_r1 = self.mapping.get("R1", {}).get("cells", {})
        
        data_map = {
            "denomination_sociale": company_info.denomination_sociale,
            "num_identification_fiscale": company_info.num_identification_fiscale or "",
            "exercice_debut": format_date(company_info.exercice_debut) if company_info.exercice_debut else "",
            "exercice_fin": format_date(company_info.exercice_fin) if company_info.exercice_fin else "",
            "date_arrete_comptes": format_date(company_info.date_arrete_comptes) if company_info.date_arrete_comptes else "",
            "exercice_precedent_fin": format_date(company_info.exercice_precedent_fin) if company_info.exercice_precedent_fin else "",
        }
        
        count = 0
        for cell_ref, mapping_info in mapping_r1.items():
            attribute = mapping_info.get("attribute")
            if attribute in data_map:
                value = data_map[attribute]
                if self._write_cell_safe(ws, cell_ref, value):
                    count += 1
        
        logger.info(f"   ✓ {count} cells filled in R1")
        return True
    
    def prefill_r2(self, company_info: DSF_InfosGenerales):
        """Remplit la page R2 (Caractérisation)"""
        if "R2" not in self.wb.sheetnames:
            logger.warning("⚠  Sheet R2 not found")
            return False
        
        ws = self.wb["R2"]
        logger.info("📝 Filling R2 (Caractérisation) sheet...")
        
        mapping_r2 = self.mapping.get("R2", {}).get("cells", {})
        
        data_map = {
            "forme_juridique": company_info.forme_juridique or "",
            "registre_fiscal": company_info.registre_fiscal or "",
        }
        
        count = 0
        for cell_ref, mapping_info in mapping_r2.items():
            attribute = mapping_info.get("attribute")
            if attribute in data_map:
                value = data_map[attribute]
                if self._write_cell_safe(ws, cell_ref, value):
                    count += 1
        
        # Remplir le tableau des activités si présent
        if "activites" in company_info.__dict__ and company_info.__dict__["activites"]:
            table_config = self.mapping.get("R2", {}).get("tables", {}).get("activites", {})
            if table_config:
                logger.info("   Filling activities table...")
                count += self._fill_table(ws, "activites", company_info.__dict__["activites"], table_config)
        
        logger.info(f"   ✓ {count} cells filled in R2")
        return True
    
    def _fill_table(self, ws, table_name: str, data_list: list, table_config: dict):
        """Remplit un tableau dans le sheet"""
        count = 0
        data_start_row = int(table_config.get("data_start_row", "A7").replace("A", ""))
        
        for idx, row_data in enumerate(data_list):
            current_row = data_start_row + idx
            for col, field_name in table_config.get("fields", {}).items():
                if field_name in row_data:
                    cell_ref = f"{col}{current_row}"
                    value = row_data[field_name]
                    if self._write_cell_safe(ws, cell_ref, value):
                        count += 1
        
        return count
    
    def save(self, output_path: str):
        """Sauvegarde le fichier rempli"""
        output_path = Path(output_path)
        logger.info(f"💾 Saving to: {output_path}")
        self.wb.save(str(output_path))
        logger.info(f"   ✓ File saved: {output_path}")
        return output_path


def main():
    """Test avec les données GULFCAM"""
    prefiller = DSFPrefillerRespectful(
        template_path="DSF Normal standard.xlsx",
        mapping_path="dsf_prefill_mapping.json"
    )
    
    # Charger le template
    prefiller.load_template()
    
    # Charger le mapping
    prefiller.load_mapping()
    
    # Créer des données de test
    company_info = DSF_InfosGenerales(
        denomination_sociale="GULFCAM S.A.S.",
        sigle_usuel="GULFCAM",
        num_identification_fiscale="R.C. 1998/SDE/CM",
        adresse_complete="Douala, Cameroun",
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
    
    # Remplir les pages
    prefiller.prefill_entete(company_info)
    prefiller.prefill_r1(company_info)
    prefiller.prefill_r2(company_info)
    
    # Sauvegarder
    output = prefiller.save("DSF_GULFCAM_CLEAN.xlsx")
    
    logger.info(f"\n✅ Pre-filling complete!")
    logger.info(f"   Output: {output}")


if __name__ == "__main__":
    main()
