# -*- coding: utf-8 -*-
"""
Smart General Filler - Reconnaissance automatique des zones de saisie
Analyse les feuilles ENTÊTE, R1, R2, R3, NOTE13 et détecte automatiquement
les cellules à remplir en analysant les labels et patterns.
"""
from __future__ import annotations

import logging
import re
from dataclasses import asdict
from datetime import date
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from dsf_general_info import DSF_InfosGenerales, format_date

logger = logging.getLogger(__name__)


class InputZone:
    """Représente une zone de saisie détectée"""
    def __init__(
        self,
        sheet: str,
        cell_ref: str,
        row: int,
        col: int,
        label: str,
        label_cell: str,
        zone_type: str = "text",  # text, date, number, currency
        context: str = ""
    ):
        self.sheet = sheet
        self.cell_ref = cell_ref
        self.row = row
        self.col = col
        self.label = label
        self.label_cell = label_cell
        self.zone_type = zone_type
        self.context = context
        self.matched_field: Optional[str] = None
        self.confidence: float = 0.0

    def __repr__(self):
        return f"InputZone({self.sheet}!{self.cell_ref}, label='{self.label}', type={self.zone_type})"


class SmartGeneralFiller:
    """
    Remplisseur intelligent avec reconnaissance automatique des zones de saisie.
    
    Processus:
    1. Scanne les feuilles cibles (ENTÊTE, R1, R2, R3, NOTE13)
    2. Détecte les patterns de labels (texte suivi de ":", cellules en gras, etc.)
    3. Identifie les cellules de saisie adjacentes (droite, bas, ou même cellule)
    4. Fait un matching fuzzy entre labels et attributs de DSF_InfosGenerales
    5. Remplit automatiquement avec le bon format
    """
    
    TARGET_SHEETS = ["ENTETE", "ENTÊTE", "Fiche R1", "R1", "Fiche R2", "R2", "Fiche R3", "R3", "NOTE 13", "NOTE13"]
    
    # Patterns de labels courants
    LABEL_PATTERNS = [
        r"dénomination\s*(sociale)?",
        r"raison\s*sociale",
        r"sigle",
        r"adresse",
        r"identification\s*(fiscale)?",
        r"n[°u]m[ée]ro",
        r"téléphone",
        r"t[ée]l[ée]phone",
        r"fax",
        r"email",
        r"e[-\s]?mail",
        r"exercice",
        r"capital",
        r"date",
        r"dur[ée]e",
        r"activit[ée]",
        r"forme\s*juridique",
        r"syst[èe]me\s*comptable",
        r"d[ée]but",
        r"fin",
        r"clos",
        r"centre.*d[ée]p[ôo]t",
        r"ville",
        r"pays",
        r"secteur",
        r"code",
    ]
    
    def __init__(self, template_path: Path | str):
        self.template_path = Path(template_path)
        self.wb = None
        self.input_zones: List[InputZone] = []
        self.info_dict: Dict[str, Any] = {}
        self.filled_count = 0
    
    def load(self) -> None:
        """Charge le template et détecte les zones de saisie"""
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {self.template_path}")
        
        logger.info(f"Loading template: {self.template_path}")
        self.wb = load_workbook(self.template_path, data_only=False)
        
        # Détecte les zones de saisie dans chaque feuille cible
        for sheet_name in self.wb.sheetnames:
            if self._is_target_sheet(sheet_name):
                logger.info(f"Scanning sheet: {sheet_name}")
                zones = self._detect_input_zones(self.wb[sheet_name])
                self.input_zones.extend(zones)
                logger.info(f"  → {len(zones)} input zones detected")
        
        logger.info(f"Total input zones detected: {len(self.input_zones)}")
    
    def fill(self, info: DSF_InfosGenerales) -> int:
        """Remplit automatiquement toutes les zones détectées"""
        if self.wb is None:
            raise RuntimeError("Filler not loaded. Call load() first.")
        
        self.info_dict = asdict(info)
        self.filled_count = 0
        
        # Match chaque zone avec un champ de DSF_InfosGenerales
        for zone in self.input_zones:
            field_name, confidence = self._match_field(zone.label, zone.context)
            zone.matched_field = field_name
            zone.confidence = confidence
            
            if field_name and confidence > 0.5:  # Seuil de confiance
                value = self._get_field_value(info, field_name, zone.zone_type)
                if value is not None and value != "":
                    success = self._write_zone(zone, value)
                    if success:
                        self.filled_count += 1
                        logger.debug(f"✓ Filled {zone.sheet}!{zone.cell_ref}: {field_name} = {value} (conf={confidence:.2f})")
                    else:
                        logger.debug(f"✗ Failed to fill {zone.sheet}!{zone.cell_ref}")
        
        # Remplir les tableaux de dirigeants et conseil d'administration
        filled_tables = self._fill_r3_tables(info)
        self.filled_count += filled_tables
        
        logger.info(f"Filled {self.filled_count}/{len(self.input_zones)} zones + {filled_tables} table rows")
        return self.filled_count
    
    def _fill_r3_tables(self, info: DSF_InfosGenerales) -> int:
        """Remplit les tableaux de dirigeants et conseil d'administration dans la fiche R3"""
        from openpyxl.cell.cell import MergedCell
        
        filled = 0
        
        # Chercher la fiche R3
        r3_sheet = None
        for sheet_name in self.wb.sheetnames:
            if "R3" in sheet_name.upper() or "FICHE R3" in sheet_name.upper():
                r3_sheet = self.wb[sheet_name]
                break
        
        if not r3_sheet:
            logger.warning("Fiche R3 not found, skipping tables")
            return 0
        
        logger.info(f"Filling R3 tables: {len(info.dirigeants)} dirigeants, {len(info.conseil_administration)} conseil members")
        
        # TABLEAU 1: DIRIGEANTS (ligne 9 = headers, données à partir ligne 10)
        if info.dirigeants:
            dirigeants_start_row = 10
            for idx, dirigeant in enumerate(info.dirigeants):
                if idx >= 14:  # Max 14 lignes disponibles (10-23)
                    logger.warning(f"Too many dirigeants ({len(info.dirigeants)}), only first 14 filled")
                    break
                
                row_num = dirigeants_start_row + idx
                # Écrire dans chaque colonne en gérant les cellules fusionnées
                self._write_cell_safe(r3_sheet, row_num, 1, dirigeant.nom)  # Colonne A: Nom
                self._write_cell_safe(r3_sheet, row_num, 2, dirigeant.prenoms)  # Colonne B: Prénoms
                self._write_cell_safe(r3_sheet, row_num, 3, dirigeant.qualite)  # Colonne C: Qualité
                self._write_cell_safe(r3_sheet, row_num, 4, dirigeant.num_identification_fiscale or "")  # Colonne D
                self._write_cell_safe(r3_sheet, row_num, 5, dirigeant.adresse or "")  # Colonne E
                
                filled += 5
                logger.debug(f"Filled dirigeant row {row_num}: {dirigeant.nom} {dirigeant.prenoms}")
        
        # TABLEAU 2: CONSEIL D'ADMINISTRATION (ligne 32 = headers, données à partir ligne 33)
        if info.conseil_administration:
            conseil_start_row = 33
            for idx, membre in enumerate(info.conseil_administration):
                if idx >= 15:  # Max environ 15 lignes
                    logger.warning(f"Too many conseil members ({len(info.conseil_administration)}), only first 15 filled")
                    break
                
                row_num = conseil_start_row + idx
                # Écrire dans chaque colonne en gérant les cellules fusionnées
                self._write_cell_safe(r3_sheet, row_num, 1, membre.nom)  # Colonne A: Nom
                self._write_cell_safe(r3_sheet, row_num, 2, membre.prenoms)  # Colonne B: Prénoms
                self._write_cell_safe(r3_sheet, row_num, 3, membre.qualite)  # Colonne C: Qualité
                self._write_cell_safe(r3_sheet, row_num, 5, membre.adresse or "")  # Colonne E: Adresse
                
                filled += 4
                logger.debug(f"Filled conseil row {row_num}: {membre.nom} {membre.prenoms}")
        
        logger.info(f"Filled {filled} cells in R3 tables")
        return filled
    
    def _write_cell_safe(self, ws, row: int, col: int, value: Any) -> bool:
        """Écrit dans une cellule en gérant les cellules fusionnées"""
        from openpyxl.cell.cell import MergedCell
        
        try:
            cell = ws.cell(row=row, column=col)
            
            # Si c'est une cellule fusionnée, trouver la cellule principale
            if isinstance(cell, MergedCell):
                # Trouver la zone fusionnée contenant cette cellule
                for merged_range in ws.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        # Écrire dans la cellule top-left de la zone fusionnée
                        top_left = merged_range.start_cell
                        ws[top_left.coordinate].value = value
                        logger.debug(f"Wrote to merged cell {cell.coordinate} via {top_left.coordinate}")
                        return True
                return False
            else:
                # Cellule normale
                cell.value = value
                return True
        except Exception as e:
            logger.warning(f"Failed to write cell {row},{col}: {e}")
            return False
    
    def save(self, output_path: Path | str) -> Path:
        """Sauvegarde le workbook rempli"""
        if self.wb is None:
            raise RuntimeError("Workbook not loaded")
        
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(output)
        logger.info(f"Saved to: {output}")
        return output
    
    def close(self) -> None:
        """Ferme le workbook"""
        if self.wb:
            self.wb.close()
    
    def get_report(self) -> Dict[str, Any]:
        """Génère un rapport de remplissage"""
        by_sheet = {}
        for zone in self.input_zones:
            if zone.sheet not in by_sheet:
                by_sheet[zone.sheet] = {"detected": 0, "filled": 0, "zones": []}
            by_sheet[zone.sheet]["detected"] += 1
            if zone.matched_field:
                by_sheet[zone.sheet]["filled"] += 1
            by_sheet[zone.sheet]["zones"].append({
                "cell": zone.cell_ref,
                "label": zone.label,
                "type": zone.zone_type,
                "matched_field": zone.matched_field,
                "confidence": zone.confidence
            })
        
        return {
            "total_zones": len(self.input_zones),
            "total_filled": self.filled_count,
            "fill_rate": self.filled_count / len(self.input_zones) if self.input_zones else 0,
            "by_sheet": by_sheet
        }
    
    # ------------------------------------------------------------------
    # Détection des zones de saisie
    # ------------------------------------------------------------------
    
    def _is_target_sheet(self, sheet_name: str) -> bool:
        """Vérifie si la feuille est une cible pour le remplissage"""
        normalized = sheet_name.upper().strip()
        for target in self.TARGET_SHEETS:
            if target.upper() in normalized:
                return True
        return False
    
    def _detect_input_zones(self, ws: Worksheet) -> List[InputZone]:
        """
        Détecte les zones de saisie dans une feuille.
        
        Stratégies:
        1. Label suivi de ":" avec cellule vide à droite
        2. Label en gras avec cellule vide en dessous
        3. Cellule fusionnée contenant un pattern de label avec zone vide
        4. Rows avec alternance label/valeur
        """
        zones = []
        sheet_name = ws.title
        
        # Scan les 50 premières lignes (zone typique des headers/infos générales)
        for row_idx in range(1, min(51, ws.max_row + 1)):
            for col_idx in range(1, min(10, ws.max_column + 1)):  # Colonnes A-J
                cell = ws.cell(row=row_idx, column=col_idx)
                
                # Stratégie 1: Label avec ":" et cellule vide à droite
                zone = self._detect_label_colon_pattern(ws, cell, row_idx, col_idx, sheet_name)
                if zone:
                    zones.append(zone)
                    continue
                
                # Stratégie 2: Label en gras avec cellule vide en dessous
                zone = self._detect_bold_label_pattern(ws, cell, row_idx, col_idx, sheet_name)
                if zone:
                    zones.append(zone)
                    continue
                
                # Stratégie 3: Cellule fusionnée avec pattern de label
                zone = self._detect_merged_label_pattern(ws, cell, row_idx, col_idx, sheet_name)
                if zone:
                    zones.append(zone)
        
        return zones
    
    def _detect_label_colon_pattern(
        self, ws: Worksheet, cell: Cell, row: int, col: int, sheet: str
    ) -> Optional[InputZone]:
        """
        Détecte: "Label:" [cellule vide]
        ou: "Label: _______"
        """
        if not isinstance(cell.value, str):
            return None
        
        text = cell.value.strip()
        if not text:
            return None
        
        # Cherche les patterns de labels
        if not self._contains_label_keyword(text):
            return None
        
        # Vérifie si le texte contient ":" ou se termine par ":"
        label = text
        input_cell = None
        zone_type = "text"
        
        if ":" in text:
            # Cas 1: "Label : " avec valeur dans cellule adjacente
            if text.endswith(":") or text.endswith(": "):
                # Cellule de saisie à droite
                input_cell = ws.cell(row=row, column=col+1)
                label = text.rstrip(":").strip()
            else:
                # Cas 2: "Label : _____" ou "Label : [vide]" dans même cellule
                parts = text.split(":", 1)
                label = parts[0].strip()
                value_part = parts[1].strip()
                
                # Si la partie valeur est vide ou contient des underscores, c'est une zone de saisie
                if not value_part or value_part.replace("_", "").strip() == "":
                    input_cell = cell
                else:
                    # Déjà rempli, ignorer
                    return None
        else:
            # Label sans ":", chercher cellule vide à droite
            input_cell = ws.cell(row=row, column=col+1)
            label = text
        
        # Vérifie que la cellule d'input est vide ou presque
        if input_cell and not self._is_empty_or_placeholder(input_cell):
            return None
        
        # Détermine le type de zone
        zone_type = self._infer_zone_type(label)
        
        input_ref = input_cell.coordinate if input_cell else None
        if not input_ref:
            return None
        
        return InputZone(
            sheet=sheet,
            cell_ref=input_ref,
            row=input_cell.row,
            col=input_cell.column,
            label=label,
            label_cell=cell.coordinate,
            zone_type=zone_type,
            context=text
        )
    
    def _detect_bold_label_pattern(
        self, ws: Worksheet, cell: Cell, row: int, col: int, sheet: str
    ) -> Optional[InputZone]:
        """
        Détecte: Label en gras
                 [cellule vide en dessous]
        """
        if not isinstance(cell.value, str):
            return None
        
        text = cell.value.strip()
        if not text or not self._contains_label_keyword(text):
            return None
        
        # Vérifie si en gras
        is_bold = cell.font and cell.font.bold
        if not is_bold:
            return None
        
        # Cellule en dessous
        input_cell = ws.cell(row=row+1, column=col)
        
        # Vérifie que la cellule d'input est vide
        if not self._is_empty_or_placeholder(input_cell):
            return None
        
        zone_type = self._infer_zone_type(text)
        
        return InputZone(
            sheet=sheet,
            cell_ref=input_cell.coordinate,
            row=input_cell.row,
            col=input_cell.column,
            label=text,
            label_cell=cell.coordinate,
            zone_type=zone_type,
            context=text
        )
    
    def _detect_merged_label_pattern(
        self, ws: Worksheet, cell: Cell, row: int, col: int, sheet: str
    ) -> Optional[InputZone]:
        """
        Détecte les cellules fusionnées contenant un pattern de label et zone de saisie.
        Exemple: "Dénomination sociale : _______________"
        """
        if not isinstance(cell.value, str):
            return None
        
        text = cell.value.strip()
        if not text or len(text) < 5:
            return None
        
        # Vérifie si c'est une cellule fusionnée
        is_merged = False
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                is_merged = True
                break
        
        if not is_merged:
            return None
        
        # Cherche pattern "Label : ____" ou "Label : [espace]"
        if ":" not in text:
            return None
        
        parts = text.split(":", 1)
        label = parts[0].strip()
        value_part = parts[1].strip()
        
        if not self._contains_label_keyword(label):
            return None
        
        # Vérifie si la zone de valeur est vide ou placeholder
        if value_part and not value_part.replace("_", "").replace(" ", "").strip() == "":
            # Déjà rempli
            return None
        
        zone_type = self._infer_zone_type(label)
        
        return InputZone(
            sheet=sheet,
            cell_ref=cell.coordinate,
            row=row,
            col=col,
            label=label,
            label_cell=cell.coordinate,
            zone_type=zone_type,
            context=text
        )
    
    def _is_empty_or_placeholder(self, cell: Cell) -> bool:
        """Vérifie si une cellule est vide ou contient juste des placeholders"""
        if cell.value is None:
            return True
        
        if isinstance(cell.value, str):
            text = cell.value.strip()
            # Vide, ou seulement des underscores, ou seulement des espaces
            if not text or text.replace("_", "").replace(" ", "").strip() == "":
                return True
        
        return False
    
    def _contains_label_keyword(self, text: str) -> bool:
        """Vérifie si le texte contient un mot-clé de label"""
        text_lower = text.lower()
        for pattern in self.LABEL_PATTERNS:
            if re.search(pattern, text_lower):
                return True
        return False
    
    def _infer_zone_type(self, label: str) -> str:
        """Infère le type de zone de saisie à partir du label"""
        label_lower = label.lower()
        
        # Email (check first)
        if any(kw in label_lower for kw in ["email", "e-mail"]):
            return "email"
        
        # Phone
        if any(kw in label_lower for kw in ["téléphone", "telephone", "fax", "tél"]):
            return "phone"
        
        # ID/Code/Number patterns (BEFORE date check)
        if any(kw in label_lower for kw in ["n°", "numero", "numéro", "code", "identification", "fiscal", "registre", "répertoire", "contribuable", "immatriculation"]):
            # Exception: if it specifically says "date" or "année", it's still a date
            if not any(kw in label_lower for kw in ["date d", "année d"]):
                return "text"
        
        # Date patterns (only if clear context)
        if any(kw in label_lower for kw in ["date", "exercice", "clos", "clôture", "arrêt", "approuvé"]):
            # Exception: "durée" is a number, not a date
            if "durée" not in label_lower and "duree" not in label_lower:
                return "date"
        
        # Duration/count = number
        if any(kw in label_lower for kw in ["durée", "duree", "mois", "année", "annee"]) and "exercice" not in label_lower:
            return "number"
        
        # Quantity = number  
        if any(kw in label_lower for kw in ["nombre", "quantité", "effectif"]):
            return "number"
        
        # Currency
        if any(kw in label_lower for kw in ["capital", "montant", "valeur", "somme"]):
            return "currency"
        
        return "text"
    
    # ------------------------------------------------------------------
    # Matching avec DSF_InfosGenerales
    # ------------------------------------------------------------------
    
    def _match_field(self, label: str, context: str) -> Tuple[Optional[str], float]:
        """
        Fait un matching fuzzy entre un label et les champs de DSF_InfosGenerales.
        
        Returns:
            (field_name, confidence_score)
        """
        # Mapping manuel prioritaire (pour les cas évidents)
        manual_mappings = {
            "dénomination sociale": "denomination_sociale",
            "raison sociale": "denomination_sociale",
            "sigle": "sigle_usuel",
            "sigle usuel": "sigle_usuel",
            "adresse": "adresse_complete",
            "adresse complète": "adresse_complete",
            "identification fiscale": "num_identification_fiscale",
            "n° identification": "num_identification_fiscale",
            "numéro identification": "num_identification_fiscale",
            "niu": "num_identification_fiscale",
            "nif": "num_identification_fiscale",
            "téléphone": "telephone",
            "telephone": "telephone",
            "fax": "fax",
            "email": "email",
            "e-mail": "email",
            "exercice clos": "exercice_fin",
            "clos le": "exercice_fin",
            "date clôture": "exercice_fin",
            "date début": "exercice_debut",
            "durée": "duree_mois",
            "duree": "duree_mois",
            "mois": "duree_mois",
            "capital": "capital_social",
            "capital social": "capital_social",
            "forme juridique": "forme_juridique",
            "activité": "activite_principale",
            "activite": "activite_principale",
            "système comptable": "systeme_comptable",
            "systeme": "systeme_comptable",
            "ville": "ville",
            "pays": "pays",
            "secteur": "secteur_activite",
            "code activité": "code_activite",
            "centre dépôt": "centre_depot",
            "centre de dépôt": "centre_depot",
            "depot": "centre_depot",
        }
        
        label_normalized = label.lower().strip()
        
        # Essai mapping manuel d'abord
        for key, field in manual_mappings.items():
            if key in label_normalized:
                return field, 1.0
        
        # Sinon, fuzzy matching avec tous les champs disponibles
        best_match = None
        best_score = 0.0
        
        for field_name in self.info_dict.keys():
            # Convertir field_name en texte lisible (enlever underscores, etc.)
            field_readable = field_name.replace("_", " ").lower()
            
            # Calcul de similarité
            similarity = SequenceMatcher(None, label_normalized, field_readable).ratio()
            
            if similarity > best_score:
                best_score = similarity
                best_match = field_name
        
        return best_match, best_score
    
    def _get_field_value(self, info: DSF_InfosGenerales, field_name: str, zone_type: str) -> Any:
        """Récupère et formate la valeur d'un champ"""
        value = getattr(info, field_name, None)
        
        if value is None:
            return None
        
        # Format selon le type de zone
        if zone_type == "date" and isinstance(value, date):
            return format_date(value)
        
        if zone_type == "currency" and isinstance(value, (int, float)):
            return float(value)
        
        if zone_type == "number" and isinstance(value, (int, float)):
            return int(value) if isinstance(value, int) else value
        
        return str(value) if value else ""
    
    def _write_zone(self, zone: InputZone, value: Any) -> bool:
        """Écrit une valeur dans une zone de saisie"""
        try:
            ws = self.wb[zone.sheet]
            cell = ws[zone.cell_ref]
            
            # Gère les cellules fusionnées
            master_cell = self._get_master_cell(ws, cell.coordinate)
            
            # Si le contexte contient ":", on reconstruit "Label : Valeur"
            if ":" in zone.context and zone.cell_ref == zone.label_cell:
                # Même cellule, reformater
                new_value = f"{zone.label} : {value}"
                master_cell.value = new_value
            else:
                # Cellule séparée
                master_cell.value = value
            
            return True
        except Exception as e:
            logger.error(f"Error writing zone {zone.cell_ref}: {e}")
            return False
    
    @staticmethod
    def _get_master_cell(ws: Worksheet, coord: str) -> Cell:
        """Récupère la cellule maître d'une cellule fusionnée"""
        cell = ws[coord]
        for merged_range in ws.merged_cells.ranges:
            if coord in merged_range:
                return ws[merged_range.start_cell.coordinate]
        return cell


__all__ = ["SmartGeneralFiller", "InputZone"]
