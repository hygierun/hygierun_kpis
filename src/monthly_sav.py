"""KPI mensuels de la section SAV (règles validées avec Antoine, mode TCD).

Seuls 4 blocs sont implémentés pour l'instant : Devis SAV, Facturation vs Docs Nuls, Main d'œuvre et
Déplacement. Productivité, Nombre d'interventions, Factures émises et la section Achat/Appro viendront
dans une prochaine passe.
"""

from typing import Dict, List, Optional

import pandas as pd

from .monthly_commerce import evolution, month_bounds
from .monthly_loader import month_column, previous_month

DEVIS_SALARIES = ["DALLEAU Jimmy", "DANVIN Joel"]

DOCS_NULS_SALARIES = ["Laetitia COINTREL", "Aimana ABDEREMANE"]
DOCS_NULS_TYPES = [
    "LIVRAISON ET INSTAL", "CONTRAT", "GARANTIE", "MISE A DISPO",
    "LOCATION", "FULL SERVICE", "PRET MACHINE", "PREPA MACHINES",
]
SAV_TECHNICIENS_5 = ["DANVIN Joel", "GRONDIN Daniel", "BAREGE YANN", "BIZEUL Christophe", "ALBANY Nicolas"]

# Main d'œuvre / déplacement : le Représentant force la catégorie quand c'est un technicien connu ;
# sinon on se rabat sur la référence article (13xxx -> SAV, 40xxx -> EBC).
MO_DEPL_EBC_REPS = ["BAREGE YANN", "DANVIN Joel", "BIZEUL Christophe"]
MO_DEPL_SAV_REPS = ["GRONDIN Daniel", "ALBANY Nicolas", "DALLEAU Jimmy"]
MAIN_OEUVRE_ARTICLES = {"EBC": ["40741"], "SAV": ["13082"]}
DEPLACEMENT_ARTICLES = {"EBC": ["40742"], "SAV": ["13045", "13047", "13049", "13080"]}


def _categorie_mo_depl(representant, article_ref, ebc_articles: List[str], sav_articles: List[str]) -> Optional[str]:
    if representant in MO_DEPL_EBC_REPS:
        return "EBC"
    if representant in MO_DEPL_SAV_REPS:
        return "SAV"
    ref = str(article_ref)
    if ref in ebc_articles:
        return "EBC"
    if ref in sav_articles:
        return "SAV"
    return None


class SavCalculator:
    """Calcule les KPI SAV du mois (Devis, Facturation vs Docs Nuls, Main d'œuvre, Déplacement), avec N-1 et Mois-1."""

    def __init__(self, dfs: Dict[str, pd.DataFrame]):
        self.dfs = dfs

    @staticmethod
    def _in_month(df: pd.DataFrame, col: str, year: int, month: int) -> pd.Series:
        start, end = month_bounds(year, month)
        return (df[col] >= start) & (df[col] <= end)

    # ------------------------------------------------------------------ Devis SAV
    def devis_rows(self, year: int, month: int) -> pd.DataFrame:
        parts = [
            df[self._in_month(df, "Date", year, month) & df["Salarié"].isin(DEVIS_SALARIES)]
            for sheet, df in self.dfs.items() if sheet in ("Devis_Arch", "Devis_EnCours")
        ]
        return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=["Total HT"])

    def devis(self, year: int, month: int) -> Dict[str, float]:
        rows = self.devis_rows(year, month)
        return {"nb": len(rows), "ca": rows["Total HT"].sum()}

    # ------------------------------------------------------------------ Facturation vs Docs Nuls
    def docs_nuls_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Fact_Arch"]
        mois = self._in_month(df, "Date", year, month)

        dn = df[mois & df["Salarié"].isin(DOCS_NULS_SALARIES) & (df["PosteEtats"] == "DN")
                & df["Type Doc Nul"].isin(DOCS_NULS_TYPES)].copy()
        dn["Catégorie"] = "DN"

        non_dn = df[mois & df["Représentant"].isin(SAV_TECHNICIENS_5) & df["PosteEtats"].isna()].copy()
        non_dn["Catégorie"] = "Non DN"

        return pd.concat([dn, non_dn], ignore_index=True)

    def docs_nuls(self, year: int, month: int) -> Dict[str, float]:
        rows = self.docs_nuls_rows(year, month)
        dn, non_dn = rows[rows["Catégorie"] == "DN"], rows[rows["Catégorie"] == "Non DN"]
        return {"total": len(rows), "ca": rows["Total HT"].sum(), "dn": len(dn), "non_dn": len(non_dn)}

    # ------------------------------------------------------------------ Main d'œuvre / Déplacement (feuille MO+Depl)
    def _mo_depl_rows(self, year: int, month: int, articles: Dict[str, List[str]]) -> pd.DataFrame:
        df = self.dfs["MO+Depl"]
        col = month_column(year, month, "Qté")
        if col not in df.columns:
            return pd.DataFrame(columns=["Représentant", "Article réf", "Catégorie", col])
        rows = df[df["Article réf"].astype(str).isin(articles["EBC"] + articles["SAV"])].copy()
        rows["Catégorie"] = [
            _categorie_mo_depl(rep, ref, articles["EBC"], articles["SAV"])
            for rep, ref in zip(rows["Représentant"], rows["Article réf"])
        ]
        rows["Heures"] = rows[col]
        return rows

    def _mo_depl(self, year: int, month: int, articles: Dict[str, List[str]]) -> Dict[str, float]:
        rows = self._mo_depl_rows(year, month, articles)
        ebc = rows.loc[rows["Catégorie"] == "EBC", "Heures"].sum()
        sav = rows.loc[rows["Catégorie"] == "SAV", "Heures"].sum()
        return {"ebc": ebc, "sav": sav, "total": ebc + sav}

    def main_oeuvre_rows(self, year: int, month: int) -> pd.DataFrame:
        return self._mo_depl_rows(year, month, MAIN_OEUVRE_ARTICLES)

    def main_oeuvre(self, year: int, month: int) -> Dict[str, float]:
        return self._mo_depl(year, month, MAIN_OEUVRE_ARTICLES)

    def deplacement_rows(self, year: int, month: int) -> pd.DataFrame:
        return self._mo_depl_rows(year, month, DEPLACEMENT_ARTICLES)

    def deplacement(self, year: int, month: int) -> Dict[str, float]:
        return self._mo_depl(year, month, DEPLACEMENT_ARTICLES)

    # ------------------------------------------------------------------ Assemblage
    def compute(self, year: int, month: int) -> Dict[str, dict]:
        py, pm = previous_month(year, month)
        periods = {"cur": (year, month), "n1": (year - 1, month), "m1": (py, pm)}
        data = {
            key: {
                "devis": self.devis(y, m),
                "docs_nuls": self.docs_nuls(y, m),
                "main_oeuvre": self.main_oeuvre(y, m),
                "deplacement": self.deplacement(y, m),
            }
            for key, (y, m) in periods.items()
        }

        cur = data["cur"]
        evolutions: Dict[str, Optional[float]] = {}
        for ref in ("n1", "m1"):
            other = data[ref]
            evolutions[f"devis_nb_{ref}"] = evolution(cur["devis"]["nb"], other["devis"]["nb"])
            for champ in ("total", "ca", "dn", "non_dn"):
                evolutions[f"docs_nuls_{champ}_{ref}"] = evolution(cur["docs_nuls"][champ], other["docs_nuls"][champ])
            for bloc in ("main_oeuvre", "deplacement"):
                for champ in ("total", "ebc", "sav"):
                    evolutions[f"{bloc}_{champ}_{ref}"] = evolution(cur[bloc][champ], other[bloc][champ])

        return {"year": year, "month": month, "periods": periods, "data": data, "evolutions": evolutions}
