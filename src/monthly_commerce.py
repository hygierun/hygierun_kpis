"""KPI mensuels de la section Commerce (règles validées avec Antoine, mode TCD)."""

from typing import Dict, List, Optional, Tuple

import pandas as pd

from .kpi_calculator import KPICalculator
from .monthly_loader import month_column, previous_month

FRANCK = "LECLANCHER Franck"
TEAM_2025_EXTRA = ["MARIE-LOUISE Caroline"]

TOURNEES_EXCLUES = (0, 524)
SEUILS_LIVRAISON = (100, 150)

MIN_QTE_CONSOMMABLE = 30
MOTS_CLES_NON_CONSOMMABLES = ["MAIN D'OEUVRE", "DEPLACEMENT", "LIVRAISON", "CONTRAT", "LOYER", "INSTALLATION"]


def evolution(current: Optional[float], reference: Optional[float]) -> Optional[float]:
    """Variation relative en % (None si la référence est nulle ou absente)."""
    if current is None or reference is None or reference == 0:
        return None
    return (current / reference - 1) * 100


def team_x5(year: int) -> List[str]:
    return KPICalculator.SALES_REPS_5 + (TEAM_2025_EXTRA if year < 2026 else [])


def team_x6(year: int) -> List[str]:
    return team_x5(year) + [FRANCK]


def month_bounds(year: int, month: int) -> Tuple[pd.Timestamp, pd.Timestamp]:
    start = pd.Timestamp(year, month, 1)
    return start, start + pd.offsets.MonthEnd(0)


class CommerceCalculator:
    """Calcule tous les KPI Commerce d'un mois, avec N-1 et Mois-1."""

    def __init__(self, dfs: Dict[str, pd.DataFrame]):
        self.dfs = dfs

    @staticmethod
    def _in_month(df: pd.DataFrame, col: str, year: int, month: int) -> pd.Series:
        start, end = month_bounds(year, month)
        return (df[col] >= start) & (df[col] <= end)

    @staticmethod
    def _reps(df: pd.DataFrame, reps: List[str], with_blank: bool) -> pd.Series:
        mask = df["Représentant"].isin(reps)
        return mask | df["Représentant"].isna() if with_blank else mask

    # ------------------------------------------------------------------ Commandes passées
    def commandes_rows(self, year: int, month: int) -> pd.DataFrame:
        parts = []
        for sheet, source, amount_col in (("Cdes_Arch", "Archivée", "Total HT"), ("Cdes_ALivr", "À livrer", "A livrer Net")):
            df = self.dfs[sheet]
            rows = df[self._in_month(df, "Date", year, month) & self._reps(df, team_x6(year), True)].copy()
            rows["Source"] = source
            rows["Montant retenu"] = rows[amount_col]
            parts.append(rows)
        return pd.concat(parts, ignore_index=True)

    def commandes(self, year: int, month: int) -> Dict[str, float]:
        rows = self.commandes_rows(year, month)
        arch, alivr = rows[rows["Source"] == "Archivée"], rows[rows["Source"] == "À livrer"]
        return {
            "nb": len(rows), "nb_arch": len(arch), "nb_alivr": len(alivr),
            "ca": rows["Montant retenu"].sum(),
            "ca_arch": arch["Montant retenu"].sum(), "ca_alivr": alivr["Montant retenu"].sum(),
        }

    # ------------------------------------------------------------------ Factures / paniers moyens
    def factures_rows(self, year: int, month: int, with_franck: bool) -> pd.DataFrame:
        df = self.dfs["Fact_Arch"]
        reps = team_x6(year) if with_franck else team_x5(year)
        return df[self._in_month(df, "Date", year, month) & self._reps(df, reps, with_franck)]

    def factures(self, year: int, month: int) -> Dict[str, float]:
        avec = self.factures_rows(year, month, True)
        sans = self.factures_rows(year, month, False)
        ca_avec, ca_sans = avec["Total HT"].sum(), sans["Total HT"].sum()
        return {
            "nb": len(avec), "ca": ca_avec,
            "pm_avec": ca_avec / len(avec) if len(avec) else None,
            "nb_sans": len(sans), "ca_sans": ca_sans,
            "pm_sans": ca_sans / len(sans) if len(sans) else None,
        }

    # ------------------------------------------------------------------ Commandes en attente
    def en_attente_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Cdes_ALivr"]
        _, end = month_bounds(year, month)
        rows = df[(df["Livraison"] >= pd.Timestamp(year, 1, 1)) & (df["Livraison"] <= end)
                  & self._reps(df, team_x6(year), True)].copy()
        rows["Type"] = rows["Rlq"].astype(str).str.strip().str.upper().map(lambda v: "Cde avec reliquats" if v == "X" else "Commande initiale")
        rows["Hors Franck"] = rows["Représentant"].isin(team_x5(year))
        return rows

    def en_attente(self, year: int, month: int) -> Dict[str, Dict[str, float]]:
        rows = self.en_attente_rows(year, month)
        out = {}
        for key, sub in (("initiale", rows[rows["Type"] == "Commande initiale"]),
                         ("reliquats", rows[rows["Type"] == "Cde avec reliquats"]),
                         ("total", rows)):
            out[key] = {
                "nb": len(sub),
                "net_avec": sub["A livrer Net"].sum(),
                "net_sans": sub[sub["Hors Franck"]]["A livrer Net"].sum(),
            }
        return out

    # ------------------------------------------------------------------ Seuils de livraison
    def seuils_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Livr_Arch"].copy()
        df["Tournée"] = pd.to_numeric(df["Tournée"], errors="coerce")
        return df[df["Tournée"].notna() & ~df["Tournée"].isin(TOURNEES_EXCLUES)
                  & self._in_month(df, "Date", year, month) & self._reps(df, team_x6(year), False)]

    def seuils(self, year: int, month: int) -> Dict[str, float]:
        rows = self.seuils_rows(year, month)
        total = len(rows)
        out = {"total": total}
        for seuil in SEUILS_LIVRAISON:
            nb = int((rows["Total HT"] < seuil).sum())
            out[f"nb_{seuil}"] = nb
            out[f"part_{seuil}"] = nb / total * 100 if total else None
        return out

    # ------------------------------------------------------------------ Nouveaux clients
    def nouveaux_clients_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Stats_NewClients"]
        col = month_column(year, month)
        rows = df[(df["1F année"] == year) & (df["1F nb mois"] == 0) & df["Représentant"].isin(team_x6(year))]
        return rows[rows[col] > 0]

    def nouveaux_clients(self, year: int, month: int) -> Dict[str, float]:
        rows = self.nouveaux_clients_rows(year, month)
        return {"nb": len(rows), "ca": rows[month_column(year, month)].sum()}

    # ------------------------------------------------------------------ Top ventes consommables
    def top_ventes_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["top10"].copy()
        col_ht, col_qte = month_column(year, month, "HT"), month_column(year, month, "Qté")
        df["Article réf"] = df["Article réf"].astype(str)
        nom = df["Article"].astype(str).str.upper()
        service = nom.apply(lambda n: any(k in n for k in MOTS_CLES_NON_CONSOMMABLES))
        conso = df[(df[col_qte] >= MIN_QTE_CONSOMMABLE) & ~service]
        return conso.rename(columns={col_ht: "HT", col_qte: "Quantité"})[["Article réf", "Article", "HT", "Quantité"]]

    def top_ventes(self, year: int, month: int, n: int = 10) -> Dict[str, pd.DataFrame]:
        rows = self.top_ventes_rows(year, month)
        ca_ref = self.factures(year, month)["ca_sans"]
        rows = rows.assign(part_ca=rows["HT"] / ca_ref * 100 if ca_ref else None)
        return {
            "valeur": rows.sort_values("HT", ascending=False).head(n).reset_index(drop=True),
            "volume": rows.sort_values("Quantité", ascending=False).head(n).reset_index(drop=True),
        }

    # ------------------------------------------------------------------ Assemblage
    def compute(self, year: int, month: int) -> Dict[str, dict]:
        """Tous les KPI du mois + périodes de comparaison (N-1 et Mois-1) et évolutions."""
        py, pm = previous_month(year, month)
        periods = {"cur": (year, month), "n1": (year - 1, month), "m1": (py, pm)}

        data = {}
        for key, (y, m) in periods.items():
            data[key] = {
                "commandes": self.commandes(y, m),
                "factures": self.factures(y, m),
                "seuils": self.seuils(y, m),
                "nouveaux_clients": self.nouveaux_clients(y, m) if key != "n1" else None,
            }
        cur = data["cur"]
        cur["en_attente"] = self.en_attente(year, month)
        cur["top_ventes"] = self.top_ventes(year, month)

        def evo(ref_key: str, block: str, field: str) -> Optional[float]:
            ref = data[ref_key][block]
            return evolution(cur[block][field], ref[field]) if ref else None

        evolutions = {
            "commandes_nb_n1": evo("n1", "commandes", "nb"),
            "commandes_ca_n1": evo("n1", "commandes", "ca"),
            "factures_nb_n1": evo("n1", "factures", "nb"),
            "factures_ca_n1": evo("n1", "factures", "ca"),
            "pm_avec_n1": evo("n1", "factures", "pm_avec"),
            "pm_sans_n1": evo("n1", "factures", "pm_sans"),
            "nouveaux_clients_nb_m1": evo("m1", "nouveaux_clients", "nb"),
            "nouveaux_clients_ca_m1": evo("m1", "nouveaux_clients", "ca"),
        }
        for seuil in SEUILS_LIVRAISON:
            for ref_key in ("n1", "m1"):
                evolutions[f"seuil_{seuil}_{ref_key}"] = evo(ref_key, "seuils", f"part_{seuil}")

        return {"year": year, "month": month, "periods": periods, "data": data, "evolutions": evolutions}
