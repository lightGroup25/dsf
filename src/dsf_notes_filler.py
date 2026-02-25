# -*- coding: utf-8 -*-
"""
Module de remplissage des Notes Annexes 1-12 + Personnel + Engagements (P2-A3 / P3-A3)
Remplit les Notes à partir des données de balance normalisées, ligne par ligne,
avec détection des colonnes par header, et sans jamais écraser de cellule déjà remplie.
"""
from __future__ import annotations

import logging
import unicodedata
import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Mapping Comptes SYSCOHADA → Notes Annexes (P3-A3 : ligne par ligne)
# Chaque entrée : (préfixe_compte, libellé_ligne, type_valeur)
# type_valeur : "brut"=solde brut, "amort"=amortissements, "net"=solde net, "variation"=N-N1
# ---------------------------------------------------------------------------
NOTE_ACCOUNT_MAPPING: Dict[str, List[Tuple[str, str, str]]] = {
    # Note 1 — Immobilisations incorporelles
    "NOTE 1": [
        ("211", "Frais de developpement",        "brut"),
        ("212", "Brevets",                       "brut"),
        ("213", "Logiciels",                      "brut"),
        ("214", "Fonds commercial",               "brut"),
        ("215", "Autres",                         "brut"),
        ("281", "Amortissements",                 "amort"),
        ("291", "Provisions",                     "amort"),
    ],
    # Note 2 — Immobilisations corporelles
    "NOTE 2": [
        ("22",  "Terrains",                       "brut"),
        ("23",  "Batiments",                      "brut"),
        ("24",  "Materiel et outillage",           "brut"),
        ("25",  "Materiel de transport",           "brut"),
        ("26",  "Mobilier",                       "brut"),
        ("27",  "Materiel informatique",           "brut"),
        ("282", "Amortissements",                 "amort"),
        ("292", "Provisions",                     "amort"),
    ],
    # Note 3 — Immobilisations financières (Split A/B/C/D...)
    "NOTE 3A": [
        ("26",  "Titres de participation",        "brut"),
        ("271", "Prets",                          "brut"),
        ("272", "Depots et cautionnements",       "brut"),
    ],
    "NOTE 3B": [
        ("286", "Amortissements titres",          "amort"),
        ("296", "Provisions titres",              "amort"),
        ("297", "Provisions prets",               "amort"),
    ],
    "NOTE 3C": [
        ("26",  "Titres de participation",        "net"),
        ("27",  "Autres immobilisations financieres", "net"),
    ],
    # Note 3D — Plus-values et moins-values de cession
    "NOTE 3D": [
        ("67",  "Charges de cession",             "variation"),
        ("77",  "Produits de cession",            "variation"),
    ],
    # Note 3E — Autres immobilisations financières (détail)
    "NOTE 3E": [
        ("26",  "Titres de participation",        "net"),
        ("27",  "Autres creances immobilisees",   "net"),
    ],
    # Note 3F — Tableau d'étalement des charges immobilisées
    "NOTE 3F": [
        ("21",  "Immobilisations incorporelles",  "net"),
        ("22",  "Immobilisations corporelles",    "net"),
        ("28",  "Amortissements",                 "variation"),
        ("29",  "Provisions pour depreciation",   "variation"),
    ],
    # Note 4 — Stocks
    "NOTE 4": [
        ("31",  "Marchandises",                   "net"),
        ("32",  "Matieres premieres",             "net"),
        ("33",  "Autres approvisionnements",      "net"),
        ("34",  "Produits en cours",              "net"),
        ("35",  "Produits finis",                 "net"),
        ("36",  "Produits intermediaires",        "net"),
        ("37",  "Stocks en cours de route",       "net"),
        ("38",  "Stocks en consignation",         "net"),
    ],
    # Note 12 — Clients / Créances commerciales (détail)
    "NOTE 12": [
        ("411", "Clients",                        "net"),
        ("412", "Clients effets a recevoir",      "net"),
        ("413", "Clients douteux",                "net"),
        ("461", "Debiteurs divers",               "net"),
        ("409", "Fournisseurs debiteurs",         "net"),
    ],
    # Note 5 — Créances
    "NOTE 5": [
        ("411", "Clients",                        "net"),
        ("412", "Clients effets a recevoir",      "net"),
        ("413", "Clients douteux",                "net"),
        ("461", "Debiteurs divers",               "net"),
        ("462", "Creances sur cessions",          "net"),
        ("409", "Fournisseurs debiteurs",         "net"),
    ],
    # Note 6 — Trésorerie
    "NOTE 6": [
        ("511", "Valeurs a l encaissement",       "net"),
        ("52",  "Banques",                        "net"),
        ("571", "Caisse",                         "net"),
        ("50",  "Titres de placement",            "net"),
    ],
    # Note 7 — Capitaux propres
    "NOTE 7": [
        ("101", "Capital social",                 "net"),
        ("102", "Capital non appele",             "net"),
        ("111", "Reserve legale",                 "net"),
        ("112", "Reserves statutaires",           "net"),
        ("118", "Autres reserves",                "net"),
        ("121", "Report a nouveau",               "net"),
        ("131", "Resultat net",                   "net"),
    ],
    # Note 8 — Dettes financières
    "NOTE 8": [
        ("161", "Emprunts obligataires",          "net"),
        ("162", "Emprunts etablissements",        "net"),
        ("163", "Dettes de credit bail",          "net"),
        ("164", "Avances recues",                 "net"),
    ],
    # Note 9 — Dettes d'exploitation
    "NOTE 9": [
        ("401", "Fournisseurs",                   "net"),
        ("402", "Fournisseurs effets a payer",    "net"),
        ("42",  "Personnel",                      "net"),
        ("43",  "Organismes sociaux",             "net"),
        ("441", "Etat impots",                    "net"),
        ("444", "TVA collectee",                  "net"),
    ],
    # Note 10 — Charges
    "NOTE 10": [
        ("60",  "Achats",                         "variation"),
        ("61",  "Variation de stocks",            "variation"),
        ("62",  "Transports",                     "variation"),
        ("63",  "Services exterieurs",            "variation"),
        ("64",  "Impots et taxes",                "variation"),
        ("65",  "Autres charges",                 "variation"),
        ("66",  "Charges de personnel",           "variation"),
        ("67",  "Frais financiers",               "variation"),
    ],
    # Note 11 — Produits
    "NOTE 11": [
        ("70",  "Ventes",                         "variation"),
        ("71",  "Production vendue",              "variation"),
        ("72",  "Production stockee",             "variation"),
        ("75",  "Autres produits",                "variation"),
        ("77",  "Produits financiers",            "variation"),
    ],
    # Note 13 — Capital social (Actionnariat)
    "NOTE 13": [
        ("101", "Nombre d actions",               "net"),
        ("102", "Valeur nominale",                "net"),
    ],
    # Note 14 — Réserves
    "NOTE 14": [
        ("111", "Reserve legale",                 "net"),
        ("112", "Reserves statutaires",           "net"),
        ("118", "Autres reserves",                "net"),
        ("121", "Report a nouveau crediteur",     "net"),
        ("129", "Report a nouveau debiteur",      "net"),
    ],
    # Note 15 — Subventions
    "NOTE 15A": [
        ("141", "Subventions d equipement",       "net"),
    ],
    "NOTE 15B": [
        ("151", "Amortissements derogatoires",    "net"),
    ],
    # Note 16 — Emprunts
    "NOTE 16A": [
        ("161", "Emprunts obligataires",          "net"),
    ],
    "NOTE 16B": [
        ("162", "Emprunts banques",               "net"),
    ],
    "NOTE 16C": [
        ("164", "Avances recues",                 "net"),
        ("165", "Depots et cautionnements",       "net"),
    ],
    # Note 16B BIS — Détail étalement emprunts bancaires
    "NOTE 16B BIS": [
        ("162", "Emprunts banques",               "net"),
    ],
    # Note 17 — Fournisseurs
    "NOTE 17": [
        ("401", "Fournisseurs",                   "net"),
        ("404", "Fournisseurs d immobilisations", "net"),
    ],
    # Note 18 — Dettes sociales/fiscales
    "NOTE 18": [
        ("42",  "Personnel",                      "net"),
        ("43",  "Organismes sociaux",             "net"),
        ("441", "Etat",                           "net"),
        ("444", "TVA collectee",                  "net"),
    ],
    # Note 19 — Autres dettes
    "NOTE 19": [
        ("45",  "Associes",                       "net"),
        ("46",  "Debiteurs et crediteurs divers", "net"),
    ],
    # Note 20 — Banques
    "NOTE 20": [
        ("52",  "Banques",                        "net"),
        ("56",  "Banques etablissements de credit", "net"),
    ],
    # Note 21-30 — Gestion (Variation)
    "NOTE 21": [ ("70", "Chiffre d affaires", "variation") ],
    "NOTE 22": [ ("60", "Achats", "variation") ],
    "NOTE 23": [ ("62", "Transports", "variation") ],
    "NOTE 24": [ ("63", "Services exterieurs", "variation") ],
    "NOTE 25": [ ("64", "Impots et taxes", "variation") ],
    "NOTE 26": [ ("65", "Autres charges", "variation") ],
    "NOTE 27A": [ ("66", "Charges de personnel", "variation") ],
    "NOTE 27B": [ ("42", "Personnel", "net") ],
    "NOTE 28": [ ("68", "Amortissements", "variation") ],
    "NOTE 29": [ ("67", "Frais financiers", "variation") ],
    "NOTE 30": [ ("8", "Opérations HAO", "variation") ],
    "NOTE 31": [ ("47", "Comptes de regularisation", "net") ],
    "NOTE 32": [ ("45", "Parties liees", "net") ],
    "NOTE 33": [ ("80", "Engagements", "variation") ],
    "NOTE 34": [ ("41", "Clients", "variation"), ("40", "Fournisseurs", "variation") ],
    # Note 35 — Annexe fiscale
    "NOTE 35": [
        ("70",  "Chiffre d affaires",             "variation"),
        ("60",  "Achats",                         "variation"),
        ("66",  "Charges de personnel",           "variation"),
        ("68",  "Dotations aux amortissements",   "variation"),
    ],
    # -----------------------------------------------------------------------
    # Feuilles consolidation (CO1, C1, C2) — mêmes mappings que les notes de base
    # -----------------------------------------------------------------------
    "CO1-NOTE 3C": [
        ("26",  "Titres de participation",        "net"),
        ("27",  "Autres immobilisations financieres", "net"),
    ],
    "C1-NOTE 17": [
        ("401", "Fournisseurs",                   "net"),
        ("404", "Fournisseurs d immobilisations", "net"),
    ],
    "C1-NOTE 25": [
        ("64",  "Impots et taxes",                "variation"),
    ],
    "C2-NOTE 25": [
        ("64",  "Impots et taxes",                "variation"),
    ],
    "C1-NOTE 28": [
        ("68",  "Amortissements",                 "variation"),
    ],
    "C2-NOTE 28": [
        ("68",  "Amortissements",                 "variation"),
    ],
    "C1-NOTE 27A": [
        ("66",  "Charges de personnel",           "variation"),
    ],
    # -----------------------------------------------------------------------
    # Notes statistiques et annexes fiscales
    # -----------------------------------------------------------------------
    "NOTES STAT SOCIALE ET ENVI": [
        ("42",  "Personnel",                      "net"),
        ("43",  "Organismes sociaux",             "net"),
        ("66",  "Charges de personnel",           "variation"),
        ("421", "Salaires bruts",                 "variation"),
        ("431", "Charges sociales",               "variation"),
    ],
    "NOTES STAT A CARACTERE COMMERCI": [
        ("70",  "Chiffre d affaires",             "variation"),
        ("60",  "Achats",                         "variation"),
        ("31",  "Marchandises",                   "net"),
        ("35",  "Produits finis",                 "net"),
    ],
    "AUTRES ANNEXES FISCA": [
        ("70",  "Chiffre d affaires",             "variation"),
        ("60",  "Achats",                         "variation"),
        ("66",  "Charges de personnel",           "variation"),
        ("68",  "Amortissements",                 "variation"),
        ("67",  "Charges financieres",            "variation"),
        ("44",  "Etat impots",                    "net"),
    ],
}


# Note Personnel (C4)
NOTE_PERSONNEL_MAPPING = [
    ("421", "Salaires bruts",                 "variation"),
    ("661", "Rémunération du personnel",      "variation"),
    ("431", "Charges sociales",               "variation"),
]

# Note Engagements (C7)
NOTE_ENGAGEMENTS_MAPPING = [
    ("801", "Engagements donnés",             "net"),
    ("811", "Engagements reçus",              "net"),
    ("803", "Engagements crédit-bail",        "net"),
]


# ---------------------------------------------------------------------------
# Résultat d'une ligne de Note
# ---------------------------------------------------------------------------
@dataclass
class NoteLineResult:
    """Résultat d'une ligne de Note Annexe."""
    note: str
    account_prefix: str
    label: str
    value_type: str  # "brut", "amort", "net", "variation", "flux"
    
    # Valeurs N
    initial_n: Decimal = Decimal("0")
    augm_n: Decimal = Decimal("0")
    dim_n: Decimal = Decimal("0")
    brut_n: Decimal = Decimal("0")
    amort_n: Decimal = Decimal("0")
    net_n: Decimal = Decimal("0")
    
    # Valeurs N-1
    initial_n1: Decimal = Decimal("0")
    net_n1: Decimal = Decimal("0")
    
    # Calculs dérivés
    variation: Decimal = Decimal("0")


# ---------------------------------------------------------------------------
# Utilitaires de normalisation
# ---------------------------------------------------------------------------

def _normalize_label(text: str) -> str:
    """Normalise un libellé : minuscules, sans accents, sans espaces multiples."""
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.lower()
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _is_placeholder(val) -> bool:
    """Vrai si la valeur est un placeholder (None, chaîne vide, 0, ou ellipsis)."""
    if val is None:
        return True
    if isinstance(val, str):
        stripped = val.strip()
        return stripped in ("", "…", "...", "-", "_", "0")
    if isinstance(val, (int, float)) and val == 0:
        return True
    return False


# ---------------------------------------------------------------------------
# Classe principale
# ---------------------------------------------------------------------------
class DSFNotesFiller:
    """
    P3-A3 : Remplit les Notes Annexes ligne par ligne avec détection des colonnes par header.
    - Ne jamais écraser une cellule déjà remplie avec une valeur non nulle.
    - Détecte les colonnes de destination par leur header (N / N-1 / Brut / Amort / Net).
    - Cherche les lignes par libellé normalisé.
    - Supporte les Notes Personnel et Engagements hors bilan.
    """

    def __init__(self, wb, normalized_rows: list, previous_rows: Optional[list] = None):
        self.wb = wb
        self.normalized_rows = normalized_rows or []
        self.previous_rows = previous_rows or []
        # Index {compte: Decimal} pour N et N-1
        self._index_n:  Dict[str, Decimal] = self._build_index(self.normalized_rows)
        self._index_n1: Dict[str, Decimal] = self._build_index(self.previous_rows)
        # Cache de headers de feuilles
        self._col_header_cache: Dict[str, Dict[str, int]] = {}

    # ------------------------------------------------------------------
    # API publique
    # ------------------------------------------------------------------

    def fill_all_notes(self) -> int:
        """Remplit toutes les Notes disponibles dans le workbook. Retourne le nombre de lignes écrites."""
        total = 0
        # 1. Notes standard de la balance
        for note_name, account_lines in NOTE_ACCOUNT_MAPPING.items():
            ws = self._find_sheet(note_name)
            if ws is None:
                continue
            col_map = self._detect_columns(ws)
            lines = self._compute_note_lines(note_name, account_lines)
            written = self._write_note_structured(ws, lines, col_map, note_name)
            total += written
            if written > 0:
                logger.info("Note %s : %d lignes écrites", note_name, written)

        # 2. Note Personnel
        ws_pers = self._find_sheet("NOTE PERSONNEL") or self._find_sheet("C4")
        if ws_pers:
            col_map = self._detect_columns(ws_pers)
            lines = self._compute_note_lines("PERSONNEL", NOTE_PERSONNEL_MAPPING)
            total += self._write_note_structured(ws_pers, lines, col_map, "PERSONNEL")

        # 3. Note Engagements
        ws_eng = self._find_sheet("NOTE ENGAGEMENTS") or self._find_sheet("C7")
        if ws_eng:
            col_map = self._detect_columns(ws_eng)
            lines = self._compute_note_lines("ENGAGEMENTS", NOTE_ENGAGEMENTS_MAPPING)
            total += self._write_note_structured(ws_eng, lines, col_map, "ENGAGEMENTS")

        return total

    def compute_aggregates(self) -> Dict[str, List[NoteLineResult]]:
        """Calcule les agrégats sans écrire dans le workbook (utile pour tests)."""
        return {
            note_name: self._compute_note_lines(note_name, account_lines)
            for note_name, account_lines in NOTE_ACCOUNT_MAPPING.items()
        }

    # ------------------------------------------------------------------
    # Calculs
    # ------------------------------------------------------------------

    def _compute_note_lines(
        self, note_name: str, mapping: List[Tuple[str, str, str]]
    ) -> List[NoteLineResult]:
        """Calcule les valeurs (Initial, Augm, Dim, Final) pour chaque ligne de Note."""
        lines = []
        for prefix, label, value_type in mapping:
            # 1. Sommes N
            n = self._sum_by_prefix(self._index_n, prefix)
            
            # 2. Sommes N-1 (pour fallback net_n1)
            n1 = self._sum_by_prefix(self._index_n1, prefix)
            
            # 3. Calcul de la nature du flux (Simplification SYSCOHADA)
            classe = prefix[0]
            if classe in ("1", "4", "7"): # Passif / Produit
                augm = n["credit"]
                dim  = n["debit"]
            else: # Actif / Charge
                augm = n["debit"]
                dim  = n["credit"]

            line = NoteLineResult(
                note=note_name, account_prefix=prefix, label=label,
                value_type=value_type,
                initial_n=n["initial"],
                augm_n=augm,
                dim_n=dim,
                brut_n=n["final"],
                net_n=n["final"],
                initial_n1=n1["initial"],
                net_n1=n1["final"],
                variation=n["final"] - n1["final"]
            )
            
            if value_type == "amort" or prefix.startswith(("28", "29", "19", "39", "49", "59")):
                line.amort_n = n["final"]
                line.net_n = Decimal("0")
            
            lines.append(line)
        return lines

    def _sum_by_prefix(self, index: Dict[str, Dict[str, Decimal]], prefix: str) -> Dict[str, Decimal]:
        """Somme les composantes (initial, debit, credit, final) par préfixe."""
        totals = {"initial": Decimal("0"), "debit": Decimal("0"), "credit": Decimal("0"), "final": Decimal("0")}
        for acc, vals in index.items():
            if acc.startswith(prefix):
                for k in totals:
                    totals[k] += vals.get(k, Decimal("0"))
        return totals

    @staticmethod
    def _build_index(rows: list) -> Dict[str, Dict[str, Decimal]]:
        """
        Construit un index multidimensionnel {compte: {initial, debit, credit, final}} 
        depuis les NormalizedBalanceRow.
        """
        index: Dict[str, Dict[str, Decimal]] = {}
        for row in rows:
            try:
                compte = str(getattr(row, "compte", "") or "").strip()
                if not compte:
                    continue
                
                # Récupération des données brutes depuis les métadonnées
                raw = getattr(row, "metadata", {}).get("raw", {})
                
                initial = Decimal(str(raw.get("solde_initial", 0) or 0))
                debit   = Decimal(str(raw.get("debit", 0) or 0))
                credit  = Decimal(str(raw.get("credit", 0) or 0))
                final   = Decimal(str(getattr(row, "solde_final", 0) or 0))
                
                if compte not in index:
                    index[compte] = {
                        "initial": Decimal("0"), "debit": Decimal("0"), 
                        "credit": Decimal("0"), "final": Decimal("0")
                    }
                
                index[compte]["initial"] += initial
                index[compte]["debit"]   += debit
                index[compte]["credit"]  += credit
                index[compte]["final"]   += final
                
            except Exception as exc:
                logger.debug("Indexation compte %s : %s", getattr(row, "compte", "?"), exc)
        return index

    # ------------------------------------------------------------------
    # Détection des colonnes par header (P3-A3)
    # ------------------------------------------------------------------

    def _detect_columns(self, ws) -> Dict[str, Optional[int]]:
        """
        Détecte les colonnes de destination en lisant les headers de la feuille
        (lignes 1-12 max). Retourne un dictionnaire {rôle: numéro_colonne}.
        Rôles : "label", "brut_n", "amort_n", "net_n", "net_n1", "variation_n", "variation_n1"
        """
        key = ws.title
        if key in self._col_header_cache:
            return self._col_header_cache[key]

        role_map: Dict[str, Optional[int]] = {
            "label": None,
            "initial_n": None, "initial_n1": None,
            "augm_n": None, "dim_n": None,
            "brut_n": None, "amort_n": None,
            "net_n": None, "net_n1": None,
            "variation_n": None, "variation_n1": None,
        }

        header_texts: Dict[int, str] = {}   # {col_num: texte normalisé cumulé}
        for row_idx in range(1, 13):
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if isinstance(val, str) and val.strip():
                    existing = header_texts.get(col_idx, "")
                    header_texts[col_idx] = (existing + " " + val.strip()).strip()

        for col_idx, header_raw in header_texts.items():
            h = _normalize_label(header_raw)
            if role_map["label"] is None and any(k in h for k in ("libelle", "designation", "intitule", "nature")):
                role_map["label"] = col_idx
            elif any(k in h for k in ("ouverture", "initial", "debut", "stock au 01")) and (
                "n-1" in h or "n 1" in h or "precedent" in h or "precedant" in h
            ):
                role_map["initial_n1"] = col_idx
            elif any(k in h for k in ("ouverture", "initial", "debut", "stock au 01")) and "n-1" not in h:
                role_map["initial_n"] = col_idx
            elif any(k in h for k in ("augmentation", "acquisition", "dotation", "virement post", "entree")) and "n-1" not in h:
                role_map["augm_n"] = col_idx
            elif any(k in h for k in ("diminution", "cession", "reprise", "retrait", "sortie")) and "n-1" not in h:
                role_map["dim_n"] = col_idx
            elif "brut" in h and "n-1" not in h and "amort" not in h:
                role_map["brut_n"] = col_idx
            elif any(k in h for k in ("amort", "depreciation", "provision")) and "n-1" not in h:
                role_map["amort_n"] = col_idx
            elif ("net" in h or "valeur nette" in h or "montant" in h or "solde" in h) and (
                "n-1" in h or "n 1" in h or "precedent" in h or "precedant" in h
                or "annee prec" in h or "cloture prec" in h or "exercice prec" in h
            ):
                role_map["net_n1"] = col_idx
            elif ("net" in h or "valeur nette" in h or "montant" in h or "solde" in h) and "n-1" not in h and "precedent" not in h:
                role_map["net_n"] = col_idx
            elif ("variation" in h or "ecart" in h) and (
                "n-1" in h or "precedent" in h or "precedant" in h or "annee prec" in h
            ):
                role_map["variation_n1"] = col_idx
            elif "variation" in h or "ecart" in h:
                role_map["variation_n"] = col_idx

        # Fallback : si label pas détecté, utiliser col A ou B
        if role_map["label"] is None:
            role_map["label"] = 2  # Colonne B par défaut

        self._col_header_cache[key] = role_map
        logger.debug("Colonnes détectées pour %s : %s", ws.title, role_map)
        return role_map

    # ------------------------------------------------------------------
    # Écriture structurée (P3-A3 : ne pas écraser, colonnes par header)
    # ------------------------------------------------------------------

    def _write_note_structured(
        self,
        ws,
        lines: List[NoteLineResult],
        col_map: Dict[str, Optional[int]],
        note_name: str,
    ) -> int:
        """Écrit les lignes d'une Note dans la feuille, sans écraser les cellules non vides."""
        written = 0
        for line in lines:
            # Ne pas écrire les lignes entièrement à zéro (sauf Engagements)
            if (line.net_n == 0 and line.net_n1 == 0
                    and line.brut_n == 0 and line.amort_n == 0
                    and "ENGAGEMENTS" not in note_name.upper()):
                continue

            row_idx = self._find_row_by_label(ws, line.label)
            if row_idx is None:
                logger.debug("Libellé '%s' introuvable dans %s", line.label, ws.title)
                continue

            # Écriture des valeurs selon les rôles détectés
            col_written = 0
            
            mapping_values = [
                ("initial_n",  line.initial_n),
                ("initial_n1", line.initial_n1),
                ("augm_n",     line.augm_n),
                ("dim_n",      line.dim_n),
                ("brut_n",     line.brut_n),
                ("amort_n",    line.amort_n),
                ("net_n",      line.net_n),
                ("net_n1",     line.net_n1),
                ("variation_n",  line.variation),
                ("variation_n1", line.variation),  # variation N-1 si présente
            ]

            for role, value in mapping_values:
                col_idx = col_map.get(role)
                if col_idx is None:
                    continue
                
                cell = ws.cell(row=row_idx, column=col_idx)
                # P3-A3 : Ne jamais écraser une cellule déjà remplie (non nulle, non placeholder)
                if not _is_placeholder(cell.value):
                    continue
                
                try:
                    cell.value = float(value)
                    col_written += 1
                except Exception as exc:
                    logger.warning("Erreur écriture %s!%s%d : %s", ws.title, cell.column_letter, row_idx, exc)

            if col_written > 0:
                written += 1

        return written

    # ------------------------------------------------------------------
    # Recherche de lignes par libellé normalisé
    # ------------------------------------------------------------------

    @staticmethod
    def _find_row_by_label(ws, label: str) -> Optional[int]:
        """Cherche la ligne dont la cellule de la colonne A ou B contient le libellé (normalisé)."""
        label_norm = _normalize_label(label)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row[:3]:   # Cherche dans les 3 premières colonnes
                if isinstance(cell.value, str):
                    if label_norm in _normalize_label(cell.value):
                        return cell.row
        return None

    def _find_sheet(self, note_name: str):
        """Cherche la feuille correspondant à une Note (tolérant aux espaces/casse)."""
        note_up = note_name.upper().replace(" ", "")
        
        # P3-Q1 : Correspondance exacte d'abord (sans espaces)
        for name in self.wb.sheetnames:
            clean_name = name.upper().replace(" ", "")
            if clean_name == note_up:
                return self.wb[name]
        
        # P3-Q1 : Puis correspondance partielle intelligente (éviter de confondre 3 et 30)
        for name in self.wb.sheetnames:
            clean_name = name.upper().replace(" ", "")
            if note_up in clean_name:
                # Vérifier si ce n'est pas un faux positif (ex: NOTE 3 dans NOTE 30)
                # On accepte si c'est suivi d'une lettre (3A, 3B) mais pas d'un autre chiffre
                idx = clean_name.find(note_up)
                after = clean_name[idx + len(note_up):]
                if not after or not after[0].isdigit():
                    return self.wb[name]
        return None


# ---------------------------------------------------------------------------
# Fonction utilitaire pour le pipeline
# ---------------------------------------------------------------------------
def fill_notes_from_balance(
    wb,
    normalized_rows: list,
    previous_rows: Optional[list] = None,
) -> int:
    """
    Remplit les Notes Annexes dans le workbook DSF.
    Retourne le nombre de lignes écrites.
    """
    filler = DSFNotesFiller(wb, normalized_rows, previous_rows)
    total = filler.fill_all_notes()
    logger.info("Notes Annexes : %d lignes remplies au total", total)
    return total


__all__ = [
    "DSFNotesFiller", "NoteLineResult", "NOTE_ACCOUNT_MAPPING", "fill_notes_from_balance"
]
