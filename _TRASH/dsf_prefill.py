#!/usr/bin/env python3
"""
Module de pré-remplissage des pages DSF
Crée et remplit les pages ENTETE, R1, R2, R3 avec les informations générales
"""

import logging
from pathlib import Path
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from dsf_general_info import DSF_InfosGenerales, format_currency, format_date

logger = logging.getLogger(__name__)

class DSF_Prefiller:
    """Classe pour créer les pages DSF pré-remplies avec informations générales"""
    
    def __init__(self, template_path: str = None):
        """
        Initialise le pré-remplisseur
        
        Args:
            template_path: Chemin du fichier DSF template (optionnel)
        """
        self.wb = None
        self.template_path = template_path
        self.loaded = False
        
    def create_new_workbook(self):
        """Crée un nouveau workbook"""
        self.wb = Workbook()
        # Supprimer la feuille par défaut
        if 'Sheet' in self.wb.sheetnames:
            del self.wb['Sheet']
        self.loaded = True
        logger.info("Nouveau workbook créé")
    
    def load_template(self):
        """Charge le template DSF"""
        if not self.template_path or not Path(self.template_path).exists():
            logger.info("Pas de template, création d'un workbook vierge")
            self.create_new_workbook()
            return
        
        try:
            self.wb = load_workbook(self.template_path)
            self.loaded = True
            logger.info(f"Template chargé: {self.template_path}")
        except Exception as e:
            logger.warning(f"Impossible de charger le template: {e}")
            self.create_new_workbook()
    
    def get_or_create_sheet(self, sheet_name: str):
        """Obtient ou crée une feuille"""
        if sheet_name not in self.wb.sheetnames:
            self.wb.create_sheet(sheet_name)
            logger.debug(f"Feuille créée: {sheet_name}")
        return self.wb[sheet_name]
    
    def _write_cell(self, ws, cell_ref: str, value, bold: bool = False, bg_color: str = None):
        """Écrit une cellule avec formatage optionnel"""
        cell = ws[cell_ref]
        cell.value = value
        
        # Appliquer le formatage
        if bold:
            cell.font = Font(bold=True, size=11)
        
        if bg_color:
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        
        cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        return cell
    
    def fill_entete(self, infos: DSF_InfosGenerales):
        """Remplit la page ENTETE"""
        if not self.loaded:
            self.load_template()
        
        ws = self.get_or_create_sheet("ENTETE")
        
        # Nettoyer la feuille au besoin
        ws.sheet_format.defaultRowHeight = 18
        
        # En-têtes ministériels
        row = 1
        self._write_cell(ws, f"A{row}", "RÉPUBLIQUE DU CAMEROUN", bold=True)
        row += 1
        self._write_cell(ws, f"A{row}", "MINISTERE DES FINANCES", bold=True)
        row += 1
        self._write_cell(ws, f"A{row}", "DIRECTION GÉNÉRALE DES IMPOTS", bold=True)
        row += 1
        self._write_cell(ws, f"A{row}", f"CENTRE DE DÉPÔT DE : {infos.centre_depot}", bold=True)
        
        row += 2
        self._write_cell(ws, f"A{row}", "ÉTATS FINANCIERS NORMALISÉS", bold=True)
        row += 1
        self._write_cell(ws, f"A{row}", f"SYSTÈME COMPTABLE {infos.systeme_comptable}", bold=True)
        row += 1
        self._write_cell(ws, f"A{row}", f"EXERCICE CLOS LE : {format_date(infos.exercice_fin)}", bold=True)
        
        # Identification entreprise
        row += 3
        self._write_cell(ws, f"A{row}", "DÉSIGNATION DE L'ENTITÉ", bold=True)
        row += 2
        self._write_cell(ws, f"A{row}", "DÉNOMINATION SOCIALE :", bold=True)
        self._write_cell(ws, f"C{row}", infos.denomination_sociale)
        
        row += 1
        self._write_cell(ws, f"A{row}", "SIGLE USUEL :", bold=True)
        self._write_cell(ws, f"C{row}", infos.sigle_usuel)
        
        row += 1
        self._write_cell(ws, f"A{row}", "ADRESSE COMPLÈTE :", bold=True)
        self._write_cell(ws, f"C{row}", infos.adresse_complete)
        
        row += 1
        self._write_cell(ws, f"A{row}", "N° D'IDENTIFICATION FISCALE :", bold=True)
        self._write_cell(ws, f"C{row}", infos.num_identification_fiscale)
        
        row += 1
        self._write_cell(ws, f"A{row}", "SYSTÈME :", bold=True)
        self._write_cell(ws, f"C{row}", infos.systeme_normal)
        
        # Documents déposés
        row += 3
        self._write_cell(ws, f"A{row}", "Documents déposés", bold=True)
        row += 1
        for doc, deposited in infos.documents_deposes.items():
            checkbox = "☑" if deposited else "☐"
            doc_label = doc.replace("_", " ").upper()
            self._write_cell(ws, f"A{row}", f"{checkbox} {doc_label}")
            row += 1
        
        logger.info("Page ENTETE remplie")
    
    
    def fill_fiche_r1(self, infos: DSF_InfosGenerales):
        """Remplit la fiche R1"""
        if not self.loaded:
            self.load_template()
        
        ws = self.get_or_create_sheet("R1")
        
        self._write_cell(ws, "A1", "1 – FICHE R1", bold=True)
        
        # Identification
        row = 3
        self._write_cell(ws, f"A{row}", "Description sociale de l'entreprise :")
        self._write_cell(ws, f"D{row}", infos.denomination_sociale, bold=True)
        
        row += 1
        self._write_cell(ws, f"A{row}", "N° d'identification fiscal :")
        self._write_cell(ws, f"D{row}", infos.num_identification_fiscale, bold=True)
        
        # Exercice comptable
        row += 2
        self._write_cell(ws, f"A{row}", "ZA EXERCICE COMPTABLE :", bold=True)
        self._write_cell(ws, f"C{row}", f"DU : {format_date(infos.exercice_debut)}")
        self._write_cell(ws, f"E{row}", f"AU : {format_date(infos.exercice_fin)}")
        
        # Date d'arrêté
        row += 1
        self._write_cell(ws, f"A{row}", "ZB DATE D'ARRÊTÉ EFFECTIF DES COMPTES :", bold=True)
        self._write_cell(ws, f"D{row}", format_date(infos.date_arrete_comptes))
        
        # Exercice précédent
        if infos.exercice_precedent_fin:
            row += 1
            self._write_cell(ws, f"A{row}", "ZC EXERCICE PRÉCÉDANT CLOS LE :", bold=True)
            self._write_cell(ws, f"D{row}", format_date(infos.exercice_precedent_fin))
        
        logger.info("Fiche R1 remplie")
    
    
    def fill_fiche_r2(self, infos: DSF_InfosGenerales):
        """Remplit la fiche R2 avec les caractéristiques de l'entreprise"""
        if not self.loaded:
            self.load_template()
        
        ws = self.get_or_create_sheet("R2")
        
        self._write_cell(ws, "A1", "2 – FICHE R2 : CARACTÉRISATION DE L'ENTREPRISE", bold=True)
        
        # Forme juridique
        row = 3
        self._write_cell(ws, f"A{row}", "FORME JURIDIQUE :")
        forme_juridique_val = infos.forme_juridique.value if hasattr(infos.forme_juridique, 'value') else infos.forme_juridique
        self._write_cell(ws, f"D{row}", forme_juridique_val, bold=True)
        
        # Registre fiscal
        row += 1
        self._write_cell(ws, f"A{row}", "REGISTRE FISCAL :")
        registre_val = infos.registre_fiscal.value if hasattr(infos.registre_fiscal, 'value') else infos.registre_fiscal
        self._write_cell(ws, f"D{row}", registre_val if registre_val else "")
        
        # Activités
        row += 2
        self._write_cell(ws, f"A{row}", "ACTIVITÉS DE L'ENTREPRISE", bold=True)
        
        # En-têtes tableau
        row += 1
        self._write_cell(ws, f"A{row}", "N°", bold=True)
        self._write_cell(ws, f"B{row}", "Désignation", bold=True)
        self._write_cell(ws, f"C{row}", "Code", bold=True)
        self._write_cell(ws, f"D{row}", "% du CA", bold=True)
        self._write_cell(ws, f"E{row}", "CA HT en CFA", bold=True)
        
        # Données des activités
        row += 1
        for i, activite in enumerate(infos.activites, 1):
            self._write_cell(ws, f"A{row}", str(i))
            self._write_cell(ws, f"B{row}", activite.designation)
            self._write_cell(ws, f"C{row}", activite.code_nomenclature)
            self._write_cell(ws, f"D{row}", f"{activite.pourcentage_ca:.2f}%")
            self._write_cell(ws, f"E{row}", format_currency(activite.chiffre_affaire_ht))
            row += 1
        
        # Total
        self._write_cell(ws, f"B{row}", "TOTAL", bold=True)
        total_ca = sum(a.chiffre_affaire_ht for a in infos.activites)
        total_pourcentage = sum(a.pourcentage_ca for a in infos.activites)
        self._write_cell(ws, f"D{row}", f"{total_pourcentage:.2f}%", bold=True)
        self._write_cell(ws, f"E{row}", format_currency(total_ca), bold=True)
        
        logger.info("Fiche R2 remplie")
    
    
    def fill_fiche_r3(self, infos: DSF_InfosGenerales):
        """Remplit la fiche R3 avec les dirigeants et conseil"""
        if not self.loaded:
            self.load_template()
        
        ws = self.get_or_create_sheet("R3")
        
        self._write_cell(ws, "A1", "3 – FICHE R3 : ADMINISTRATION ET DIRECTION", bold=True)
        
        # Dirigeants
        row = 3
        self._write_cell(ws, f"A{row}", "DIRIGEANTS/ADMINISTRATEURS", bold=True)
        
        row += 1
        self._write_cell(ws, f"A{row}", "Nom", bold=True)
        self._write_cell(ws, f"B{row}", "Prénoms", bold=True)
        self._write_cell(ws, f"C{row}", "Qualité", bold=True)
        self._write_cell(ws, f"D{row}", "Adresse", bold=True)
        
        row += 1
        for dirigeant in infos.dirigeants:
            self._write_cell(ws, f"A{row}", dirigeant.nom)
            self._write_cell(ws, f"B{row}", dirigeant.prenoms)
            self._write_cell(ws, f"C{row}", dirigeant.qualite)
            self._write_cell(ws, f"D{row}", getattr(dirigeant, 'adresse', ''))
            row += 1
        
        # Conseil d'administration
        row += 2
        self._write_cell(ws, f"A{row}", "CONSEIL D'ADMINISTRATION", bold=True)
        
        row += 1
        self._write_cell(ws, f"A{row}", "Nom", bold=True)
        self._write_cell(ws, f"B{row}", "Prénoms", bold=True)
        self._write_cell(ws, f"C{row}", "Qualité", bold=True)
        self._write_cell(ws, f"D{row}", "Adresse", bold=True)
        
        row += 1
        for membre in infos.conseil_administration:
            self._write_cell(ws, f"A{row}", membre.nom)
            self._write_cell(ws, f"B{row}", membre.prenoms)
            self._write_cell(ws, f"C{row}", membre.qualite)
            self._write_cell(ws, f"D{row}", getattr(membre, 'adresse', ''))
            row += 1
        
        logger.info("Fiche R3 remplie")
    
    
    def fill_note_capital(self, infos: DSF_InfosGenerales):
        """Remplit la note 13 du capital (actionnaires)"""
        if not self.loaded:
            self.load_template()
        
        ws = self.get_or_create_sheet("NOTE13")
        
        self._write_cell(ws, "A1", "NOTE 13 - CAPITAL : VALEUR NOMINALE DES ACTIONS/PARTS", bold=True)
        
        # Entête informations
        row = 3
        self._write_cell(ws, f"A{row}", "ENTITÉ :")
        self._write_cell(ws, f"C{row}", infos.denomination_sociale)
        
        row += 1
        self._write_cell(ws, f"A{row}", "N° IDENTIFICATION :")
        self._write_cell(ws, f"C{row}", infos.num_identification_fiscale)
        
        row += 1
        self._write_cell(ws, f"A{row}", "EXERCICE AU :")
        self._write_cell(ws, f"C{row}", format_date(infos.exercice_fin))
        
        # Tableau actionnaires
        row += 2
        self._write_cell(ws, f"A{row}", "Noms/Prénoms/Raison Sociale", bold=True)
        self._write_cell(ws, f"B{row}", "Nationalité", bold=True)
        self._write_cell(ws, f"C{row}", "Nature Actions", bold=True)
        self._write_cell(ws, f"D{row}", "Nombre", bold=True)
        self._write_cell(ws, f"E{row}", "Montant Total CFA", bold=True)
        
        row += 1
        total_montant = 0
        
        for actionnaire in infos.actionnaires:
            nom_complet = f"{actionnaire.nom} {actionnaire.prenoms}".strip()
            self._write_cell(ws, f"A{row}", nom_complet)
            self._write_cell(ws, f"B{row}", getattr(actionnaire, 'nationalite', ''))
            self._write_cell(ws, f"C{row}", getattr(actionnaire, 'nature_actions', ''))
            self._write_cell(ws, f"D{row}", str(getattr(actionnaire, 'nombre', '')))
            montant = getattr(actionnaire, 'montant_total', 0)
            self._write_cell(ws, f"E{row}", format_currency(montant))
            total_montant += montant
            row += 1
        
        # Total
        row += 1
        self._write_cell(ws, f"A{row}", "TOTAL CAPITAL APPORTÉ", bold=True)
        self._write_cell(ws, f"E{row}", format_currency(total_montant), bold=True)
        
        # Capital non appelé
        if infos.capital_non_appele and infos.capital_non_appele > 0:
            row += 1
            self._write_cell(ws, f"A{row}", "Capital non appelé")
            self._write_cell(ws, f"E{row}", format_currency(infos.capital_non_appele))
            
            row += 1
            self._write_cell(ws, f"A{row}", "TOTAL CAPITAL", bold=True)
            self._write_cell(ws, f"E{row}", format_currency(total_montant + infos.capital_non_appele), bold=True)
        
        logger.info("Note 13 remplie")
    
    def fill_all(self, infos: DSF_InfosGenerales):
        """Remplit toutes les pages du DSF"""
        if not self.loaded:
            self.load_template()
        
        self.fill_entete(infos)
        self.fill_fiche_r1(infos)
        self.fill_fiche_r2(infos)
        self.fill_fiche_r3(infos)
        self.fill_note_capital(infos)
        
        logger.info("✓ Toutes les pages remplies avec succès")
    
    def save(self, output_path: str = None):
        """Sauvegarde le fichier DSF"""
        if not self.loaded:
            raise RuntimeError("Aucun fichier DSF chargé")
        
        if output_path is None:
            output_path = self.dsf_file
        
        self.wb.save(output_path)
        logger.info(f"Fichier sauvegardé: {output_path}")
    
    def _write_cell(self, ws, cell_ref: str, value, bold: bool = False):
        """Écrit une cellule avec formatage optionnel"""
        try:
            cell = ws[cell_ref]
            
            # Vérifier si c'est une cellule fusionnée - si oui, utiliser le coin supérieur gauche
            if isinstance(cell, type(None)) or str(type(cell).__name__) == 'MergedCell':
                # Trouver la cellule de départ de la fusion
                for merged_range in ws.merged_cells.ranges:
                    if cell_ref in merged_range:
                        cell_ref = merged_range.start_cell.coordinate
                        cell = ws[cell_ref]
                        break
            
            cell.value = value
            
            if bold:
                cell.font = Font(bold=True)
            
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        except (AttributeError, RuntimeError) as e:
            # En cas d'erreur, ignorer et continuer
            logger.warning(f"Impossible d'écrire dans {cell_ref}: {e}")

def prefill_dsf_with_infos(dsf_file: str, infos: DSF_InfosGenerales, output_file: str = None):
    """
    Remplit facilement un fichier DSF avec les informations générales
    
    Args:
        dsf_file: Chemin du fichier DSF
        infos: Objet DSF_InfosGenerales contenant les infos
        output_file: Fichier de sortie (optionnel)
    """
    prefiller = DSF_Prefiller(dsf_file)
    prefiller.load_dsf()
    prefiller.fill_all(infos)
    prefiller.save(output_file or dsf_file)
    
    return True

if __name__ == "__main__":
    # Test avec GULFCAM
    from dsf_general_info import GULFCAM_SAS
    
    logging.basicConfig(level=logging.INFO)
    
    dsf_path = "DSF_GULFCAM_PREFILLED.xlsx"
    prefill_dsf_with_infos("DSF Normal standard.xlsx", GULFCAM_SAS, dsf_path)
    print(f"✓ DSF pré-rempli sauvegardé: {dsf_path}")
