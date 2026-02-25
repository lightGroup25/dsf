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
    # Note 28 — Provisions et Dépréciations inscrites au bilan
    # Structure colonnes : B=Initial(A), C=DotExpl(B), D=DotFin, E=DotHAO,
    #                      F=RepExpl(C), G=RepFin, H=RepHAO, I=Clôture(D=A+B-C)
    # Chaque entrée : (prefixe_stock, label_ligne, prefixe_dot_expl, prefixe_dot_fin,
    #                  prefixe_dot_hao, prefixe_rep_expl, prefixe_rep_fin, prefixe_rep_hao)
    "NOTE 28": [
        ("14",  "Provisions reglementees",               "6815", "6875", "83",  "7815", "7875", "84"),
        ("19",  "Provisions financieres pour risques",   "6597", "6974", "839", "7597", "7974", "849"),
        ("29",  "Depreciations des immobilisations",     "6811", "6813", "831", "7811", "7813", "841"),
        ("39",  "Depreciations des stocks",              "6591", "6593", "835", "7591", "7593", "845"),
        ("49",  "Depreciations clients",                 "6591", "6593", "835", "7591", "7593", "845"),
        ("59",  "Depreciations fournisseurs",            "6596", "6976", "838", "7596", "7976", "848"),
    ],
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
        ("14",  "Provisions reglementees",  "6815", "6875", "83",  "7815", "7875", "84"),
        ("19",  "Provisions financieres",   "6597", "6974", "839", "7597", "7974", "849"),
        ("29",  "Depreciations immos",      "6811", "6813", "831", "7811", "7813", "841"),
        ("39",  "Depreciations stocks",     "6591", "6593", "835", "7591", "7593", "845"),
        ("49",  "Depreciations creances",   "6591", "6593", "835", "7591", "7593", "845"),
        ("59",  "Depreciations tresorerie", "6596", "6976", "838", "7596", "7976", "848"),
    ],
    "C2-NOTE 28": [
        ("14",  "Provisions reglementees",  "6815", "6875", "83",  "7815", "7875", "84"),
        ("19",  "Provisions financieres",   "6597", "6974", "839", "7597", "7974", "849"),
        ("29",  "Depreciations immos",      "6811", "6813", "831", "7811", "7813", "841"),
        ("39",  "Depreciations stocks",     "6591", "6593", "835", "7591", "7593", "845"),
        ("49",  "Depreciations creances",   "6591", "6593", "835", "7591", "7593", "845"),
        ("59",  "Depreciations tresorerie", "6596", "6976", "838", "7596", "7976", "848"),
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

    # Set of note names that use the Note-28 matrix format
    _NOTE28_NAMES: set = {"NOTE 28", "C1-NOTE 28", "C2-NOTE 28"}

    def fill_all_notes(self) -> int:
        """Remplit toutes les Notes disponibles dans le workbook. Retourne le nombre de lignes écrites."""
        total = 0
        # 1. Notes standard de la balance
        for note_name, account_lines in NOTE_ACCOUNT_MAPPING.items():
            ws = self._find_sheet(note_name)
            if ws is None:
                logger.debug("Feuille non trouvée pour '%s' (feuilles: %s)", note_name, self.wb.sheetnames)
                continue

            # Note 28 uses a special matrix layout (Initial/DotExpl/DotFin/DotHAO/RepExpl/RepFin/RepHAO/Final)
            if note_name in self._NOTE28_NAMES:
                written = self._fill_note28(ws, account_lines, note_name)
            else:
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
    # Remplissage spécialisé Note 28 (Provisions et Dépréciations)
    # ------------------------------------------------------------------

    def _fill_note28(self, ws, account_lines: list, note_name: str) -> int:
        """
        Remplit la Note 28 selon sa structure matricielle :
        Col B=Initial(A), C=Dot.Expl(B), D=Dot.Fin, E=Dot.HAO,
        F=Rep.Expl(C), G=Rep.Fin, H=Rep.HAO, I=Clôture(D=A+B-C).

        account_lines contient des tuples de 8 éléments :
        (prefixe_stock, label_ligne,
         pref_dot_expl, pref_dot_fin, pref_dot_hao,
         pref_rep_expl, pref_rep_fin, pref_rep_hao)
        """
        # Détection des colonnes Note-28 par header
        col_map = self._detect_note28_columns(ws)
        logger.debug("Note 28 col_map: %s", col_map)

        written = 0
        for entry in account_lines:
            if len(entry) != 8:
                # Ancienne ligne (3 éléments), on skip
                logger.warning("Note 28: entrée ignorée (format incorrect): %s", entry)
                continue
            (
                stock_prefix, label,
                pref_dot_expl, pref_dot_fin, pref_dot_hao,
                pref_rep_expl, pref_rep_fin, pref_rep_hao,
            ) = entry

            # Calcul des valeurs depuis la balance
            stock_n  = self._sum_by_prefix(self._index_n,  stock_prefix)
            stock_n1 = self._sum_by_prefix(self._index_n1, stock_prefix)

            dot_expl = self._sum_by_prefix(self._index_n, pref_dot_expl)["final"]
            dot_fin  = self._sum_by_prefix(self._index_n, pref_dot_fin) ["final"]
            dot_hao  = self._sum_by_prefix(self._index_n, pref_dot_hao) ["final"]

            rep_expl = self._sum_by_prefix(self._index_n, pref_rep_expl)["final"]
            rep_fin  = self._sum_by_prefix(self._index_n, pref_rep_fin) ["final"]
            rep_hao  = self._sum_by_prefix(self._index_n, pref_rep_hao) ["final"]

            # Pour les comptes de provision (passif - classe 1-4) le solde créditeur est positif
            # L'initial  = solde final N-1 (ou solde initial N si dispo)
            initial   = stock_n1["final"]
            cloture   = stock_n ["final"]

            # Skip si tout est zéro
            if all(v == 0 for v in [initial, dot_expl, dot_fin, dot_hao,
                                    rep_expl, rep_fin, rep_hao, cloture]):
                continue

            row_idx = self._find_row_by_label(ws, label, min_row=13)
            if row_idx is None:
                logger.debug("Note 28: libellé introuvable '%s' dans %s", label, ws.title)
                continue

            col_written = 0
            values_to_write = [
                ("initial_n",  initial),
                ("dot_expl_n", dot_expl),
                ("dot_fin_n",  dot_fin),
                ("dot_hao_n",  dot_hao),
                ("rep_expl_n", rep_expl),
                ("rep_fin_n",  rep_fin),
                ("rep_hao_n",  rep_hao),
                # La clôture (col I) est souvent une formule D=A+B-C, on ne la force pas
                # mais on l'écrit si la cellule est vide et qu'il n'y a pas de formule.
                ("cloture_n",  cloture),
            ]
            for role, val in values_to_write:
                col_idx = col_map.get(role)
                if col_idx is None:
                    continue
                cell = ws.cell(row=row_idx, column=col_idx)
                if not _is_placeholder(cell.value):
                    continue
                # Ne pas écraser les formules
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue
                try:
                    cell.value = float(val)
                    col_written += 1
                except Exception as exc:
                    logger.warning("Note 28 écriture %s!%s%d : %s",
                                   ws.title, cell.column_letter, row_idx, exc)
            if col_written > 0:
                written += 1

        return written

    def _detect_note28_columns(self, ws) -> Dict[str, Optional[int]]:
        """
        Détecte les colonnes spécifiques de la Note 28 :
        B=Initial(A), C=DotExpl, D=DotFin, E=DotHAO,
        F=RepExpl, G=RepFin, H=RepHAO, I=Clôture(D=A+B-C).
        Utilise les lettres de colonne fixes si les headers confirment la structure,
        sinon fallback sur les lettres standard B-I.
        """
        role_map: Dict[str, Optional[int]] = {
            "initial_n":  None,
            "dot_expl_n": None,
            "dot_fin_n":  None,
            "dot_hao_n":  None,
            "rep_expl_n": None,
            "rep_fin_n":  None,
            "rep_hao_n":  None,
            "cloture_n":  None,
        }

        # Phase 1 : scanner les 15 premières lignes pour trouver les headers
        header_texts: Dict[int, str] = {}
        for row_idx in range(1, min(16, ws.max_row + 1)):
            for col_idx in range(1, min(ws.max_column + 1, 15)):
                val = ws.cell(row=row_idx, column=col_idx).value
                if isinstance(val, str) and val.strip():
                    existing = header_texts.get(col_idx, "")
                    header_texts[col_idx] = (existing + " " + val.strip()).strip()

        # Phase 2 : assignation par mots-clés
        dot_cols: List[int] = []
        rep_cols: List[int] = []

        for col_idx, raw_header in header_texts.items():
            h = _normalize_label(raw_header)
            if ("ouverture" in h or "initial" in h or "debut" in h) and role_map["initial_n"] is None:
                role_map["initial_n"] = col_idx
            elif ("cloture" in h or "fermeture" in h or "d a b c" in h or "d=a" in h) and role_map["cloture_n"] is None:
                role_map["cloture_n"] = col_idx
            elif "dotation" in h or "augmentation" in h:
                dot_cols.append(col_idx)
            elif "reprise" in h or "diminution" in h:
                rep_cols.append(col_idx)

        # Phase 3 : identifier les sous-colonnes Expl/Fin/HAO
        def _assign_sub_cols(cols: List[int], role_prefix: str, role_map: dict, ws) -> None:
            """Sous-classification des colonnes par exploitation/financiere/hao."""
            if not cols:
                return
            annotated = []
            for c in cols:
                h = _normalize_label(header_texts.get(c, ""))
                if any(k in h for k in ("exploitation", "d exploitat")):
                    annotated.append(("expl", c))
                elif any(k in h for k in ("financier", "financiere", "fin")):
                    annotated.append(("fin", c))
                elif any(k in h for k in ("hao", "hors activit", "hors activites")):
                    annotated.append(("hao", c))
                else:
                    annotated.append(("unknown", c))  # fallback

            for kind, cidx in annotated:
                role_key = f"{role_prefix}_{kind}_n" if kind != "unknown" else f"{role_prefix}_expl_n"
                if role_key in role_map and role_map[role_key] is None:
                    role_map[role_key] = cidx

        _assign_sub_cols(sorted(dot_cols), "dot", role_map, ws)
        _assign_sub_cols(sorted(rep_cols), "rep", role_map, ws)

        # Phase 4 : fallback sur colonnes fixes si la détection echoue
        # Standard Note 28 SYSCOHADA: B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9
        defaults = {
            "initial_n":  2,   # col B
            "dot_expl_n": 3,   # col C
            "dot_fin_n":  4,   # col D
            "dot_hao_n":  5,   # col E
            "rep_expl_n": 6,   # col F
            "rep_fin_n":  7,   # col G
            "rep_hao_n":  8,   # col H
            "cloture_n":  9,   # col I
        }
        for role, default_col in defaults.items():
            if role_map[role] is None:
                role_map[role] = default_col

        logger.debug("Note 28 colonnes détectées pour '%s': %s", ws.title, role_map)
        return role_map

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
        Supporte deux sources de données:
        1. metadata.raw (si dispo, de balance_transformer)
        2. debit_balance / credit_balance / solde_final (de BalanceNormalizer.iterate)
        """
        index: Dict[str, Dict[str, Decimal]] = {}
        for row in rows:
            try:
                compte = str(getattr(row, "compte", "") or "").strip()
                if not compte:
                    continue
                
                # Source 1 : métadonnées brutes (balance_transformer)
                raw = getattr(row, "metadata", {}).get("raw", {})
                
                initial = Decimal(str(raw.get("solde_initial", 0) or 0))
                debit   = Decimal(str(raw.get("debit", 0) or 0))
                credit  = Decimal(str(raw.get("credit", 0) or 0))
                final   = Decimal(str(getattr(row, "solde_final", 0) or 0))

                # Source 2 : fallback sur les champs directs (BalanceNormalizer.iterate)
                if debit == 0 and credit == 0:
                    debit  = Decimal(str(getattr(row, "debit_balance",  0) or 0))
                    credit = Decimal(str(getattr(row, "credit_balance", 0) or 0))
                
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
        (lignes 1-20). Retourne un dictionnaire {rôle: numéro_colonne}.
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
        for row_idx in range(1, min(21, ws.max_row + 1)):
            for col_idx in range(1, min(ws.max_column + 1, 40)):
                val = ws.cell(row=row_idx, column=col_idx).value
                if isinstance(val, str) and val.strip():
                    existing = header_texts.get(col_idx, "")
                    header_texts[col_idx] = (existing + " " + val.strip()).strip()

        for col_idx, header_raw in header_texts.items():
            h = _normalize_label(header_raw)
            if role_map["label"] is None and any(k in h for k in ("libelle", "designation", "intitule", "nature", "compte")):
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
            elif ("net" in h or "valeur nette" in h or "montant" in h or "solde" in h or "exercice n" in h or "cloture n" in h) and (
                "n-1" in h or "n 1" in h or "precedent" in h or "precedant" in h
                or "annee prec" in h or "cloture prec" in h or "exercice prec" in h or "exercice n 1" in h
            ):
                role_map["net_n1"] = col_idx
            elif ("net" in h or "valeur nette" in h or "montant" in h or "solde" in h or "exercice n" in h or "cloture n" in h) and "n-1" not in h and "precedent" not in h and "precedant" not in h:
                role_map["net_n"] = col_idx
            elif ("variation" in h or "ecart" in h) and (
                "n-1" in h or "precedent" in h or "precedant" in h or "annee prec" in h
            ):
                role_map["variation_n1"] = col_idx
            elif "variation" in h or "ecart" in h:
                role_map["variation_n"] = col_idx

        # Fallback : si label présent mais pas net_n/net_n1, prendre les 2 premières colonnes numériques après le label
        label_col = role_map.get("label") or 2
        if not role_map.get("net_n") or not role_map.get("net_n1"):
            numeric_cols: List[int] = []
            for col_idx in range(label_col + 1, min(label_col + 7, ws.max_column + 1)):
                for row_idx in range(1, min(ws.max_row + 1, 100)):
                    v = ws.cell(row=row_idx, column=col_idx).value
                    if v is not None and isinstance(v, (int, float)) and col_idx not in numeric_cols:
                        numeric_cols.append(col_idx)
                        break
            if not role_map.get("net_n") and len(numeric_cols) >= 1:
                role_map["net_n"] = numeric_cols[0]
            if not role_map.get("net_n1") and len(numeric_cols) >= 2:
                role_map["net_n1"] = numeric_cols[1]
            elif not role_map.get("net_n1") and len(numeric_cols) == 1:
                role_map["net_n1"] = numeric_cols[0]  # même colonne en secours

        if role_map["label"] is None:
            role_map["label"] = 2

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
    def _find_row_by_label(ws, label: str, min_row: int = 1) -> Optional[int]:
        """Cherche la ligne dont une cellule (colonnes 1-5) contient le libellé (normalisé, ou mots clés).
        
        min_row : ignorer les lignes avant ce numéro (utile pour passer les entêtes de section).
        """
        label_norm = _normalize_label(label)
        # Mots significatifs du libellé (longueur >= 2, pour matcher "et" / "au" etc. avec prudence)
        label_words = [w for w in label_norm.split() if len(w) >= 2]
        for row in ws.iter_rows(min_row=min_row, max_row=ws.max_row):
            for cell in row[:5]:   # Colonnes A à E
                if not isinstance(cell.value, str) or not cell.value.strip():
                    continue
                cell_norm = _normalize_label(cell.value)
                if label_norm in cell_norm:
                    return cell.row
                # Correspondance partielle : au moins 2 mots du libellé présents dans la cellule
                if len(label_words) >= 2 and sum(1 for w in label_words if w in cell_norm) >= 2:
                    return cell.row
                # Un seul mot long (ex: "fournisseurs") peut suffire si très spécifique
                if label_words and len(label_words[0]) >= 6 and label_words[0] in cell_norm:
                    return cell.row
        return None

    def _find_sheet(self, note_name: str):
        """Cherche la feuille correspondant à une Note (tolérant espaces, tirets, casse)."""
        # Variantes : avec/sans espaces, avec/sans tirets (ex: C1-NOTE 25 / C1 NOTE 25)
        note_up = note_name.upper().replace(" ", "").replace("-", "")

        for name in self.wb.sheetnames:
            clean_name = name.upper().replace(" ", "").replace("-", "")
            if clean_name == note_up:
                return self.wb[name]
            # NOTE 4, NOTE 5, NOTE 25, etc. : feuille peut s'appeler "NOTE 4 - Stocks" -> "NOTE4STOCKS"
            if clean_name.startswith(note_up):
                # Éviter NOTE 3 = NOTE 30 : si note_up est "NOTE3", pas accepter "NOTE30"
                after = clean_name[len(note_up):]
                if not after or not after[0].isdigit():
                    return self.wb[name]

        # Correspondance partielle (note_up contenu dans le nom)
        for name in self.wb.sheetnames:
            clean_name = name.upper().replace(" ", "").replace("-", "")
            if note_up in clean_name:
                idx = clean_name.find(note_up)
                after = clean_name[idx + len(note_up):]
                if not after or not after[0].isdigit():
                    return self.wb[name]
        # Feuille nommée uniquement par le numéro/sous-numéro (ex: "15A", "16A", "17")
        for name in self.wb.sheetnames:
            clean_name = name.upper().replace(" ", "").replace("-", "")
            if len(clean_name) >= 2 and note_up.endswith(clean_name):
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
