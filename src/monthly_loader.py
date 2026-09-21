"""Chargement et validation des fichiers input du Bilan Mensuel (section par section)."""

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

COMMERCE_DETAIL_COLUMNS: Dict[str, List[str]] = {
    "Cdes_ALivr": ["N°", "Client", "Total HT"],
    "Cdes_Arch": ["N°", "Client"],
    "Fact_Arch": ["N°", "Client"],
    "Livr_Arch": ["N°", "Client"],
}

COMMERCE_DATE_COLUMNS = ["Date", "Livraison", "Liv. souhaitée"]


def commerce_required_columns(year: int, month: int) -> Dict[str, List[str]]:
    """Colonnes obligatoires par feuille pour le mois traité (colonnes mensuelles incluses)."""
    py, pm = previous_month(year, month)
    required = {sheet: list(cols) for sheet, cols in COMMERCE_REQUIRED_COLUMNS.items()}
    required["Stats_NewClients"] += [month_column(year, month), month_column(py, pm)]
    required["top10"] += [month_column(year, month, "HT"), month_column(year, month, "Qté")]
    return required


class MonthlyDataLoader:
    """Charge les feuilles d'un Input mensuel et vérifie la présence des colonnes requises."""

    def __init__(self, excel_path: str):
        self.excel_path = str(excel_path)
        self.dfs: Dict[str, pd.DataFrame] = {}
        self.errors: List[str] = []

    def load_commerce(self, year: int, month: int) -> bool:
        required = commerce_required_columns(year, month)
        self.errors = []

        if not Path(self.excel_path).exists():
            self.errors.append(f"Fichier introuvable : {self.excel_path}")
            return False

        with pd.ExcelFile(self.excel_path) as xl:
            missing_sheets = [s for s in required if s not in xl.sheet_names]
            if missing_sheets:
                self.errors.append("Feuilles manquantes : " + ", ".join(missing_sheets))
                return False

            for sheet, cols in required.items():
                wanted = set(cols) | set(COMMERCE_DETAIL_COLUMNS.get(sheet, []))
                df = xl.parse(sheet, usecols=lambda c, w=wanted: c in w)
                missing_cols = [c for c in cols if c not in df.columns]
                if missing_cols:
                    self.errors.append(f"Feuille '{sheet}' : colonnes manquantes : {', '.join(missing_cols)}")
                    continue
                for col in COMMERCE_DATE_COLUMNS:
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col], errors="coerce")
                self.dfs[sheet] = df

        return not self.errors
