"""Chargement et validation des fichiers input du Bilan Mensuel (section par section)."""

import warnings
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
    "Ach_Recep": ["Fournisseur", "Total HT", "FFR", "Container", "Réception"],
}

LIVRAISON_COMPTA_REQUIRED_COLUMNS: Dict[str, List[str]] = {
    "Tournees": ["Date", "Désignation", "Camion", "Total HT", "Nb bl A", "Nb fact"],
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
}

DETAIL_COLUMNS: Dict[str, List[str]] = {
    "Cdes_ALivr": ["N°", "Client", "Total HT"],
    "Cdes_Arch": ["N°", "Client"],
    "Fact_Arch": ["N°", "Client"],
    "Livr_Arch": ["N°", "Client", "Représentant"],
    "Ach_Recep": ["N°", "Date"],
    "Tournees": ["N°", "Chauffeur"],
    "Delais Bis": ["N°"],
    "Fact_Dues": ["N°", "Client"],
    "Clients": ["Référence"],
    "Devis_Arch": ["N°", "Client"],
    "Devis_EnCours": ["N°", "Client"],
    "MO+Depl": ["Article"],
}

DATE_COLUMNS = ["Date", "Livraison", "Liv. souhaitée", "Réception", "Fact date", "Date Creation Cde"]


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
    """MO+Depl a besoin des colonnes Qté du mois, de N-1 et de Mois-1 (feuille = tout l'historique en colonnes)."""
    required = {sheet: list(cols) for sheet, cols in SAV_ACHAT_REQUIRED_COLUMNS.items()}
    py, pm = previous_month(year, month)
    required["MO+Depl"] += [
        month_column(year, month, "Qté"), month_column(py, pm, "Qté"), month_column(year - 1, month, "Qté"),
    ]
    return required


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
                wanted = set(cols) | set(DETAIL_COLUMNS.get(sheet, []))
                df = xl.parse(sheet, usecols=lambda c, w=wanted: c in w)
                missing_cols = [c for c in cols if c not in df.columns]
                if missing_cols:
                    self.errors.append(f"Feuille '{sheet}' : colonnes manquantes : {', '.join(missing_cols)}")
                    continue
                for col in DATE_COLUMNS:
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col], errors="coerce")
                self.dfs[sheet] = df

        return not self.errors
