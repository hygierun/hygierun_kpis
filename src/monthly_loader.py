"""Chargement et validation des fichiers input du Bilan Mensuel (section par section)."""

import warnings
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd

MONTHS_FR = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
]


def month_column(year: int, month: int, suffix: str = "HT") -> str:
    """Nom de colonne mensuelle des extracts, ex. (2026, 8) -> '08/26 HT'."""
    return f"{month:02d}/{year % 100:02d} {suffix}"


def previous_month(year: int, month: int) -> Tuple[int, int]:
    return (year - 1, 12) if month == 1 else (year, month - 1)


COMMERCE_REQUIRED_COLUMNS: Dict[str, List[str]] = {
    "Cdes_ALivr": ["Date", "Livraison", "Rlq", "Représentant", "A livrer Net"],
    "Cdes_Arch": ["Date", "Représentant", "Total HT"],
    "Fact_Arch": ["Date", "Représentant", "Total HT"],
    "Livr_Arch": ["Date", "Tournée", "Représentant", "Total HT"],
    "Stats_NewClients": ["Client", "1F année", "1F nb mois", "Représentant"],
    "top10": ["Article réf", "Article"],
}

PREPARATION_REQUIRED_COLUMNS: Dict[str, List[str]] = {
    "Livr_Arch": ["Date", "Tournée", "Client", "Total HT"],
    "Ach_Rcp_Arch": ["Fournisseur", "Total HT", "FFR", "Container", "Réception"],
}

LIVRAISON_COMPTA_REQUIRED_COLUMNS: Dict[str, List[str]] = {
    "Tournees_Arch": ["Date", "Désignation", "Camion", "Total HT", "Nb bl A", "Nb fact"],
    "Livr_Arch": ["Date", "Tournée", "Transporteur", "Total HT"],
    "Delais Bis": ["Date", "Liv. souhaitée", "Tournée", "Date Creation Cde", "Fact date"],
    "Cdes_Arch": ["Livraison", "Tournée", "Représentant", "Nb bls"],
    "Fact_Dues": ["Date", "Client (réf.)", "Nb JEch", "Restant dû"],
    "Clients": ["Désignation", "Qualification", "Solde cpta"],
    "GPS_Livr": ["Véhicule", "Date", "Arrivée", "Carnet arrivée", "Arrêt", "Km", "Adresse arrivée"],
}

SAV_ACHAT_REQUIRED_COLUMNS: Dict[str, List[str]] = {
    "Fact_Arch": ["Date", "Salarié", "Représentant", "PosteEtats", "Type Doc Nul", "Total HT"],
    "Devis_Arch": ["Date", "Salarié", "Total HT"],
    "Devis_EnCours": ["Date", "Salarié", "Total HT"],
    "MO+Depl": ["Article réf", "Représentant"],
    "Bilan_Fiches": ["Equipe", "Nb interventions", "Nb heures en intervention"],
    "Absences_Tech": ["Noms", "Nb jours Absences"],
    "Stock_Invent": ["Dépot", "Famille"],
    "Ach_Arti": ["Réappro", "Dispo"],
}

# Feuilles chargées sans restriction de colonnes (colonne dynamique du jour, ex. "T 04/09/26").
FULL_SHEETS = {"Stock_Invent"}

DETAIL_COLUMNS: Dict[str, List[str]] = {
    "Cdes_ALivr": ["N°", "Client", "Total HT"],
    "Cdes_Arch": ["N°", "Client"],
    "Fact_Arch": ["N°", "Client"],
    "Livr_Arch": ["N°", "Client", "Représentant"],
    "Ach_Rcp_Arch": ["N°", "Date"],
    "Tournees_Arch": ["N°", "Chauffeur"],
    "Delais Bis": ["N°"],
    "Fact_Dues": ["N°", "Client"],
    "Clients": ["Référence"],
    "Devis_Arch": ["N°", "Client"],
    "Devis_EnCours": ["N°", "Client"],
    "MO+Depl": ["Article"],
    "Ach_Arti": ["Fournisseur", "Désignation", "Référence", "Famille"],
}

DATE_COLUMNS = ["Date", "Livraison", "Liv. souhaitée", "Réception", "Fact date", "Date Creation Cde"]


def _coerce_mixed_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Une cellule datée glissée par erreur dans une colonne normalement numérique (Total HT, Heures,
    Restant dû...) fait tomber toute la colonne en dtype 'object' avec un mélange float/datetime, ce
    qui casse .sum() (TypeError: unsupported operand type(s) for +: 'float' and 'datetime.datetime').
    On détecte précisément ce mélange (jamais légitime : aucune colonne n'est censée contenir à la
    fois des nombres et des dates) et on force la colonne en numérique - la date invalide devient NaN,
    comme n'importe quelle autre valeur non numérique déjà tolérée ailleurs."""
    for col in df.columns:
        series = df[col]
        if series.dtype != object:
            continue
        has_number = series.map(lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)).any()
        has_datetime = series.map(lambda v: isinstance(v, datetime)).any()
        if has_number and has_datetime:
            df[col] = pd.to_numeric(series, errors="coerce")
    return df


def commerce_required_columns(year: int, month: int) -> Dict[str, List[str]]:
    """Colonnes obligatoires par feuille pour le mois traité (colonnes mensuelles incluses)."""
    py, pm = previous_month(year, month)
    required = {sheet: list(cols) for sheet, cols in COMMERCE_REQUIRED_COLUMNS.items()}
    required["Stats_NewClients"] += [month_column(year, month), month_column(py, pm)]
    required["top10"] += [month_column(year, month, "HT"), month_column(year, month, "Qté")]
    return required


def livraison_compta_required_columns(year: int, month: int) -> Dict[str, List[str]]:
    required = {sheet: list(cols) for sheet, cols in LIVRAISON_COMPTA_REQUIRED_COLUMNS.items()}
    required["Clients"].append(f"Vtes {year}")
    return required


def preparation_required_columns(year: int, month: int) -> Dict[str, List[str]]:
    return {sheet: list(cols) for sheet, cols in PREPARATION_REQUIRED_COLUMNS.items()}


def sav_achat_required_columns(year: int, month: int) -> Dict[str, List[str]]:
    """MO+Depl a besoin des colonnes Qté du mois, de N-1 et de Mois-1 (feuille = tout l'historique en colonnes).
    Ach_Arti a besoin de la colonne Ventes de l'année en cours (ex. 'Vtes 2026')."""
    required = {sheet: list(cols) for sheet, cols in SAV_ACHAT_REQUIRED_COLUMNS.items()}
    py, pm = previous_month(year, month)
    required["MO+Depl"] += [
        month_column(year, month, "Qté"), month_column(py, pm, "Qté"), month_column(year - 1, month, "Qté"),
    ]
    required["Ach_Arti"] += [f"Vtes {year}"]
    return required


def bilan_total_required_columns(year: int, month: int) -> Dict[str, List[str]]:
    """Union des feuilles/colonnes des 4 sections pour le fichier Input Mensuel consolidé.

    3 feuilles sont partagées entre sections (Cdes_Arch, Fact_Arch, Livr_Arch) : on garde à chaque fois
    l'union des colonnes demandées par chaque section qui s'en sert, jamais un sous-ensemble - dérivé
    directement des fonctions *_required_columns existantes, pas dupliqué à la main, pour ne jamais
    diverger si une section change ses colonnes requises."""
    merged: Dict[str, List[str]] = {}
    for required in (
        commerce_required_columns(year, month),
        preparation_required_columns(year, month),
        livraison_compta_required_columns(year, month),
        sav_achat_required_columns(year, month),
    ):
        for sheet, cols in required.items():
            existing = merged.setdefault(sheet, [])
            for col in cols:
                if col not in existing:
                    existing.append(col)
    return merged


class MonthlyDataLoader:
    """Charge les feuilles d'un Input mensuel et vérifie la présence des colonnes requises."""

    def __init__(self, excel_path: str):
        self.excel_path = str(excel_path)
        self.dfs: Dict[str, pd.DataFrame] = {}
        self.errors: List[str] = []

    def load_commerce(self, year: int, month: int) -> bool:
        return self._load(commerce_required_columns(year, month))

    def load_preparation(self, year: int, month: int) -> bool:
        return self._load(preparation_required_columns(year, month))

    def load_livraison_compta(self, year: int, month: int) -> bool:
        return self._load(livraison_compta_required_columns(year, month))

    def load_sav_achat(self, year: int, month: int) -> bool:
        return self._load(sav_achat_required_columns(year, month))

    def load_bilan_total(self, year: int, month: int) -> bool:
        return self._load(bilan_total_required_columns(year, month))

    def _load(self, required: Dict[str, List[str]]) -> bool:
        self.errors = []

        if not Path(self.excel_path).exists():
            self.errors.append(f"Fichier introuvable : {self.excel_path}")
            return False

        warnings.filterwarnings("ignore", module="openpyxl")
        with pd.ExcelFile(self.excel_path) as xl:
            missing_sheets = [s for s in required if s not in xl.sheet_names]
            if missing_sheets:
                self.errors.append("Feuilles manquantes : " + ", ".join(missing_sheets))
                return False

            for sheet, cols in required.items():
                if sheet in FULL_SHEETS:
                    df = xl.parse(sheet)
                else:
                    wanted = set(cols) | set(DETAIL_COLUMNS.get(sheet, []))
                    df = xl.parse(sheet, usecols=lambda c, w=wanted: c in w)
                missing_cols = [c for c in cols if c not in df.columns]
                if missing_cols:
                    self.errors.append(f"Feuille '{sheet}' : colonnes manquantes : {', '.join(missing_cols)}")
                    continue
                for col in DATE_COLUMNS:
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col], errors="coerce")
                self.dfs[sheet] = _coerce_mixed_numeric_columns(df)

        return not self.errors
