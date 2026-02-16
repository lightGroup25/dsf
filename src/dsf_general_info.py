# -*- coding: utf-8 -*-
#!/usr/bin/env python3
"""
Module de gestion des informations générales DSF
Contient les données de pré-remplissage pour ENTETE, R1, R2, R3
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional
from datetime import date
from enum import Enum

class Forme_Juridique(Enum):
    """Codes de forme juridique"""
    ENTREPRISE_INDIVIDUELLE = "01"
    EIRL = "02"
    SARL = "03"
    EURL = "04"
    SA = "05"
    SA_DIRECTOIRE = "06"
    GIE = "07"
    SAS = "08"
    AUTRE = "99"

class Registre_Fiscal(Enum):
    """Registres fiscaux"""
    IUTS = "1"
    IMV = "2"
    SEMI_TAXE = "3"
    TME = "4"
    PME = "5"

@dataclass
class Dirigeant:
    """Représente un dirigeant de l'entreprise"""
    nom: str
    prenoms: str
    qualite: str  # PRÉSIDENT, DG, DGA, ADMINISTRATEUR, etc.
    num_identification_fiscale: str = ""
    adresse: str = ""

@dataclass
class Conseil_Administration:
    """Représente un membre du conseil"""
    nom: str
    prenoms: str
    qualite: str  # PRÉSIDENT, MEMBRE, etc.
    adresse: str = ""

@dataclass
class Activite_Entreprise:
    """Représente une activité de l'entreprise"""
    designation: str
    code_nomenclature: str
    chiffre_affaire_ht: float = 0.0
    pourcentage_ca: float = 0.0

@dataclass
class Actionnaire:
    """Représente un actionnaire"""
    nom: str
    prenoms: str = ""
    nationalite: str = ""
    nature_actions: str = "ORDINAIRES"
    nombre: int = 0
    montant_total: float = 0.0
    cessions_remboursements: str = ""

@dataclass
class DSF_InfosGenerales:
    """Informations générales du DSF - Fiche ENTETE"""
    
    # Identification entreprise
    denomination_sociale: str
    sigle_usuel: str
    adresse_complete: str
    num_identification_fiscale: str
    
    # Exercice comptable
    exercice_debut: date
    exercice_fin: date
    date_arrete_comptes: date
    exercice_precedent_fin: Optional[date] = None
    duree_mois: int = 12
    
    # Informations légales
    forme_juridique: str = "08"  # SAS par défaut
    registre_fiscal: str = "1"  # IUTS par défaut
    pays_siege_social: str = "2"  # Cameroun
    
    # Centre de dépôt
    centre_depot: str = "DIRECTION DES GRANDES ENTREPRISES"
    ministere: str = "MINISTERE DES FINANCES"
    direction_generale: str = "DIRECTION GENERALE DES IMPOTS"
    
    # Localisation
    pays: str = "CAMEROUN"
    ville: str = "DOUALA"
    region: str = ""
    
    # Système comptable
    systeme_comptable: str = "SYSTEME NORMAL"
    
    # Note 13 specific
    note_13_commentaire: str = ""

    # Listes de détails
    activites: List[Activite_Entreprise] = field(default_factory=list)
    dirigeants: List[Dirigeant] = field(default_factory=list)
    conseil_administration: List[Conseil_Administration] = field(default_factory=list)
    actionnaires: List[Actionnaire] = field(default_factory=list)

    @property
    def periode_exercice_str(self) -> str:
        d = self.exercice_debut.strftime("%d/%m/%Y")
        f = self.exercice_fin.strftime("%d/%m/%Y")
        return f"DU : {d} AU : {f}"

def format_date(d: Optional[date]) -> str:
    """Formate une date au format JJ/MM/AAAA"""
    if d is None:
        return ""
    return d.strftime("%d/%m/%Y")

def get_gulfcam_config() -> DSF_InfosGenerales:
    """Retourne la configuration spécifique pour GULFCAM SAS 2024"""
    return DSF_InfosGenerales(
        denomination_sociale="GULFCAM SAS",
        sigle_usuel="GULFCAM SAS",
        adresse_complete="CENTRE DES AFFAIRES MARITIMES, B.P 3876 DOUALA",
        num_identification_fiscale="M050900027774W",
        
        exercice_debut=date(2024, 1, 1),
        exercice_fin=date(2024, 12, 31),
        date_arrete_comptes=date(2024, 12, 31),
        duree_mois=12,
        
        forme_juridique="08",
        registre_fiscal="1",
        
        activites=[
            Activite_Entreprise(
                designation="TRANSPORT MARITIME",
                code_nomenclature="0 3 4 0 0",
                chiffre_affaire_ht=3_744_332_179,
                pourcentage_ca=0.04  # 4%
            ),
            Activite_Entreprise(
                designation="VENTE DES PRODUITS PETROLIERS ET AUTRES",
                code_nomenclature="0 3 1 0 0",
                chiffre_affaire_ht=16_797_523_137,
                pourcentage_ca=0.94  # 94%
            ),
            Activite_Entreprise(
                designation="AUTRES SERVICES VENDUS",
                code_nomenclature="0 3 8 0 0",
                chiffre_affaire_ht=3_312_121_101,
                pourcentage_ca=0.02  # 2%
            ),
        ],
        
        dirigeants=[
            Dirigeant(
                nom="NYODOG",
                prenoms="PERRIAL JEAN",
                qualite="PRESIDENT",
                adresse="3876 DOUALA, CAMEROUN"
            ),
            Dirigeant(
                nom="BOU",
                prenoms="ALBERT ROGER",
                qualite="DG",
                adresse="3876 DOUALA, CAMEROUN"
            ),
            Dirigeant(
                nom="MIKE",
                prenoms="LEOPOLD",
                qualite="DGA",
                adresse="3876 DOUALA, CAMEROUN"
            ),
        ],
        
        conseil_administration=[
            Conseil_Administration(
                nom="MBAYEN",
                prenoms="RENE",
                qualite="PRESIDENT CONSEIL SURVEILLANCE",
                adresse="DOUALA, CAMEROUN"
            ),
            Conseil_Administration(
                nom="NYODOG",
                prenoms="PERRIAL JEAN",
                qualite="PRESIDENT DU DIRECTOIRE",
                adresse="DOUALA, CAMEROUN"
            ),
            Conseil_Administration(
                nom="TCHOUTA MOUSSA",
                prenoms="ESTHER",
                qualite="MEMBRE",
                adresse="DOUALA, CAMEROUN"
            ),
            Conseil_Administration(
                nom="NJOVAGE",
                prenoms="GEORGES",
                qualite="MEMBRE",
                adresse="DOUALA, CAMEROUN"
            ),
            Conseil_Administration(
                nom="MBAYEN HEGBA",
                prenoms="FLORIAN",
                qualite="MEMBRE",
                adresse="DOUALA, CAMEROUN"
            ),
            Conseil_Administration(
                nom="NGO MBAYEN Epse MALANANGE DORA",
                prenoms="",
                qualite="MEMBRE",
                adresse="DOUALA, CAMEROUN"
            ),
            Conseil_Administration(
                nom="FAYCAL",
                prenoms="ABDOULAYE",
                qualite="MEMBRE",
                adresse="YAOUNDE, CAMEROUN"
            ),
            Conseil_Administration(
                nom="WOLINO Epse BOOTO A N",
                prenoms="COLETTE",
                qualite="MEMBRE",
                adresse="DOUALA, CAMEROUN"
            ),
        ],
        
        actionnaires=[
            Actionnaire(
                nom="SOFIMAR SICAV",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=276_954,
                montant_total=2_769_540_000
            ),
            Actionnaire(
                nom="MINISTERE DES FINANCES",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=66_657,
                montant_total=666_570_000
            ),
            Actionnaire(
                nom="LIQUIDATION CIC",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=50_443,
                montant_total=504_430_000
            ),
            Actionnaire(
                nom="CIP-GIE",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=32_291,
                montant_total=322_910_000
            ),
            Actionnaire(
                nom="TRIANGLE SA",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=19_050,
                montant_total=190_500_000
            ),
            Actionnaire(
                nom="SNI",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=8_716,
                montant_total=87_160_000
            ),
            Actionnaire(
                nom="TCHOUTA MOUSSA ESTHER",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=5_062,
                montant_total=50_620_000
            ),
            Actionnaire(
                nom="MBAYEN RENE",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=4_000,
                montant_total=40_000_000
            ),
            Actionnaire(
                nom="CAPLAME",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=2_405,
                montant_total=24_050_000
            ),
            Actionnaire(
                nom="HESNAULT",
                nationalite="CAMEROUNAISE",
                nature_actions="ORDINAIRES",
                nombre=1_683,
                montant_total=16_830_000
            ),
            Actionnaire(
                nom="Apporteurs, capital non appelé",
                nationalite="",
                nombre=467_261,
                montant_total=4_672_610_000
            ),
        ],
        
        note_13_commentaire="L'actionnariat repose sur l'opération de fusion-absorption de CAMSHIP-CLGG par GULFIN devenue GULFCAM du 28 octobre 2021, AGM du 18 novembre 2021 et AGE du 22 juin 2022."
    )

if __name__ == "__main__":
    import json
    from dataclasses import asdict
    
    # Test de la configuration
    config = get_gulfcam_config()
    print(f"Configuration pour {config.denomination_sociale} chargée.")
    print(f"Nombre d'activités: {len(config.activites)}")
    print(f"Nombre de dirigeants: {len(config.dirigeants)}")
    print(f"Nombre de membres CA: {len(config.conseil_administration)}")
    print(f"Nombre d'actionnaires: {len(config.actionnaires)}")
