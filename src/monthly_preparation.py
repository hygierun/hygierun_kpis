"""KPI mensuels de la section Préparation (règles validées avec Antoine, mode TCD).

Les dates de préparation ne sont pas dans Neo : les volumes préparés sont estimés
à partir des dates de livraison (feuille Livr_Arch).
"""

from typing import Dict, Optional

import pandas as pd

from .monthly_commerce import evolution, month_bounds
from .monthly_loader import previous_month

TOURNEES_EXCLUES = (0, 524)
CLIENT_REASSORT = "DEPOT SHOWROOM"

# Clients "dépôt technicien" (SAV = virement de dépôt) -> date de fin de validité (None = toujours actif)
DEPOTS_SAV: Dict[str, Optional[str]] = {
    "DEPOT NICOLAS": None,
    "DEPOT DANIEL": None,
    "DEPOT YANN": None,
    "DEPOT CHRISTOPHE": None,
    "DEPOT JOEL": None,
    "DEPOT JULIEN B": "2025-07-31",
}

MOT_CLE_AUTOLIQUIDATION = "AUTOLIQUIDATION"


class PreparationCalculator:
    """Calcule les KPI Préparation d'un mois, avec N-1 et Mois-1."""

    def __init__(self, dfs: Dict[str, pd.DataFrame]):
        self.dfs = dfs

    # ------------------------------------------------------------------ Commandes préparées (livraisons archivées)
    def _livraisons_mois(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Livr_Arch"].copy()
        df["Tournée"] = pd.to_numeric(df["Tournée"], errors="coerce")
        start, end = month_bounds(year, month)
        return df[(df["Date"] >= start) & (df["Date"] <= end)]

    @staticmethod
    def _is_sav(df: pd.DataFrame) -> pd.Series:
        mask = pd.Series(False, index=df.index)
        for client, valid_until in DEPOTS_SAV.items():
            match = df["Client"] == client
            if valid_until:
                match &= df["Date"] <= pd.Timestamp(valid_until)
            mask |= match
        return mask

    def commandes_preparees_rows(self, year: int, month: int) -> pd.DataFrame:
        """Livraisons du mois classées en 'Cdes clients' / 'Réassort' / 'SAV'."""
        df = self._livraisons_mois(year, month)
        reassort = df["Client"] == CLIENT_REASSORT
        sav = self._is_sav(df)
        clients = df["Tournée"].notna() & ~df["Tournée"].isin(TOURNEES_EXCLUES) & ~reassort & ~sav

        parts = []
        for label, mask in (("Cdes clients", clients), ("Réassort", reassort), ("SAV", sav)):
            rows = df[mask].copy()
            rows["Catégorie"] = label
            parts.append(rows)
        return pd.concat(parts, ignore_index=True)

    def commandes_preparees(self, year: int, month: int) -> Dict[str, float]:
        rows = self.commandes_preparees_rows(year, month)
        count = rows["Catégorie"].value_counts()
        return {
            "clients": int(count.get("Cdes clients", 0)),
            "reassort": int(count.get("Réassort", 0)),
            "sav": int(count.get("SAV", 0)),
            "total": len(rows),
            "ca": rows["Total HT"].sum(),
        }

    # ------------------------------------------------------------------ Dépotage conteneurs (réceptions d'achats)
    def conteneurs_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Ach_Recep"]
        start, end = month_bounds(year, month)
        autoliquidation = df["Fournisseur"].astype(str).str.upper().str.contains(MOT_CLE_AUTOLIQUIDATION, na=False)
        return df[(df["Réception"] >= start) & (df["Réception"] <= end)
                  & df["Container"].notna() & (df["FFR"] != "X") & ~autoliquidation]

    def conteneurs(self, year: int, month: int) -> Dict[str, float]:
        rows = self.conteneurs_rows(year, month)
        return {"nb": int(rows["Container"].nunique()), "ca": rows["Total HT"].sum()}

    # ------------------------------------------------------------------ Assemblage
    def compute(self, year: int, month: int) -> Dict[str, dict]:
        py, pm = previous_month(year, month)
        periods = {"cur": (year, month), "n1": (year - 1, month), "m1": (py, pm)}
        data = {
            key: {"commandes": self.commandes_preparees(y, m), "conteneurs": self.conteneurs(y, m)}
            for key, (y, m) in periods.items()
        }

        cur = data["cur"]
        evolutions = {}
        for ref in ("n1", "m1"):
            for field in ("total", "ca", "clients", "reassort", "sav"):
                evolutions[f"{field}_{ref}"] = evolution(cur["commandes"][field], data[ref]["commandes"][field])
            evolutions[f"conteneurs_nb_{ref}"] = evolution(cur["conteneurs"]["nb"], data[ref]["conteneurs"]["nb"])
            evolutions[f"conteneurs_ca_{ref}"] = evolution(cur["conteneurs"]["ca"], data[ref]["conteneurs"]["ca"])

        return {"year": year, "month": month, "periods": periods, "data": data, "evolutions": evolutions}
