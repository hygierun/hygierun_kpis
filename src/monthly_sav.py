"""KPI mensuels des sections SAV et Achat/Appro (règles validées avec Antoine, mode TCD).

Blocs implémentés : Devis SAV, Facturation vs Docs Nuls, Main d'œuvre, Déplacement, Productivité,
Nombre d'interventions (SAV) et Valorisation du stock (Achat/Appro). Le nombre d'interventions et les
heures en intervention ne sont pas recalculables depuis Input_SAV_Achat.xlsx : ils viennent de la
lecture manuelle des fiches papier (feuille Bilan_Fiches, remplie chaque mois via le prompt
docs/prompt_analyse_fiches_sav.md). Idem pour leur Mois-1, lu depuis le PowerPoint du mois précédent
(voir monthly_sav_mois1.py) faute d'historique recalculable.
"""

import calendar
from typing import Dict, List, Optional

import pandas as pd

from .monthly_commerce import evolution, month_bounds
from .monthly_loader import month_column, previous_month

DEVIS_SALARIES = ["DALLEAU Jimmy", "DANVIN Joel"]

DOCS_NULS_SALARIES = ["Laetitia COINTREL", "Aimana ABDEREMANE"]
DOCS_NULS_TYPES = ["CONTRAT", "FULL SERVICE", "GARANTIE", "LIVRAISON ET INSTAL", "PREPA MACHINES", "PRET MACHINE", "LOCATION"]
# Mise à dispo (et l'absence de Type Doc Nul) ne compte en DN que si un des 5 techniciens SAV est le Représentant.
DOCS_NULS_TYPES_SI_TECH = ["MISE A DISPO"]
SAV_TECHNICIENS_5 = ["DANVIN Joel", "GRONDIN Daniel", "BAREGE YANN", "BIZEUL Christophe", "ALBANY Nicolas"]

# Main d'œuvre / déplacement : le Représentant force la catégorie quand c'est un technicien connu ;
# sinon on se rabat sur la référence article (13xxx -> SAV, 40xxx -> EBC).
MO_DEPL_EBC_REPS = ["BAREGE YANN", "DANVIN Joel", "BIZEUL Christophe"]
MO_DEPL_SAV_REPS = ["GRONDIN Daniel", "ALBANY Nicolas", "DALLEAU Jimmy"]
MAIN_OEUVRE_ARTICLES = {"EBC": ["40741"], "SAV": ["13082"]}
DEPLACEMENT_ARTICLES = {"EBC": ["40742"], "SAV": ["13045", "13047", "13049", "13080"]}

# Productivité : heures travaillées = (jours ouvrés du mois - absences) x 7h, sommées par équipe.
# Joël ne compte pas dans ce calcul (2 techniciens par équipe) - règle validée avec Antoine.
PRODUCTIVITE_TECHNICIENS = {"ebc": ["Yann", "Christophe"], "sav": ["Daniel", "Nicolas"]}
HEURES_PAR_JOUR = 7

# Articles à épuisement (Achat/Appro) : feuille Ach_Arti, Réappro = "A épuis.".
ARTICLES_EPUISEMENT_REAPPRO = "A épuis."

# Valorisation du stock (Achat/Appro) : Dépôt + Showroom, 8 familles de la feuille Stock_Invent.
STOCK_DEPOTS = ["DEPOT", "SHOWROOM"]
STOCK_FAMILLES = [
    ("Machine de nettoyage", "HYGIENE Machine de nettoyage"),
    ("Produits de Nettoyage", "HYGIENE Produits de Nettoyage"),
    ("Papier", "HYGIENE Papier"),
    ("Matériel Manuel", "HYGIENE Matériel Manuel"),
    ("Accessoires Manuel", "HYGIENE Accessoires Manuel"),
    ("Brosserie", "HYGIENE Brosserie"),
    ("Achats Local", "HYGIENE Achats Local"),
    ("Disque d'Entretien", "HYGIENE Disque d'Entretien*"),
]


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
        is_tech = df["Représentant"].isin(SAV_TECHNICIENS_5)

        type_ok = df["Type Doc Nul"].isin(DOCS_NULS_TYPES) | (
            is_tech & (df["Type Doc Nul"].isin(DOCS_NULS_TYPES_SI_TECH) | df["Type Doc Nul"].isna())
        )
        dn = df[mois & df["Salarié"].isin(DOCS_NULS_SALARIES) & (df["PosteEtats"] == "DN") & type_ok].copy()
        dn["Catégorie"] = "DN"

        non_dn = df[mois & is_tech & df["PosteEtats"].isna()].copy()
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

    # ------------------------------------------------------------------ Productivité (Nb interventions + heures)
    @staticmethod
    def _jours_ouvres(year: int, month: int) -> int:
        cal = calendar.Calendar()
        return sum(1 for d in cal.itermonthdates(year, month) if d.month == month and d.weekday() < 5)

    def heures_travaillees(self, year: int, month: int) -> Dict[str, float]:
        df = self.dfs["Absences_Tech"]
        absences = dict(zip(df["Noms"], df["Nb jours Absences"]))
        jours = self._jours_ouvres(year, month)
        result = {
            equipe: sum((jours - absences.get(nom, 0)) * HEURES_PAR_JOUR for nom in noms)
            for equipe, noms in PRODUCTIVITE_TECHNICIENS.items()
        }
        result["total"] = result["ebc"] + result["sav"]
        return result

    def _bilan_fiches(self, year: int, month: int, champ: str) -> Dict[str, float]:
        df = self.dfs["Bilan_Fiches"].set_index("Equipe")
        ebc = float(df.loc["EBC", champ]) if "EBC" in df.index else 0.0
        sav = float(df.loc["SAV", champ]) if "SAV" in df.index else 0.0
        return {"ebc": ebc, "sav": sav, "total": ebc + sav}

    def nb_interventions(self, year: int, month: int) -> Dict[str, float]:
        rows = self._bilan_fiches(year, month, "Nb interventions")
        return {k: int(v) for k, v in rows.items()}

    def heures_intervention(self, year: int, month: int) -> Dict[str, float]:
        return self._bilan_fiches(year, month, "Nb heures en intervention")

    def productivite(self, year: int, month: int) -> Dict[str, dict]:
        travaillees = self.heures_travaillees(year, month)
        interventions = self.heures_intervention(year, month)

        def pct(inter: float, trav: float) -> Optional[float]:
            return (inter / trav * 100) if trav else None

        return {
            "heures_travaillees": travaillees,
            "heures_intervention": interventions,
            "pct": {equipe: pct(interventions[equipe], travaillees[equipe]) for equipe in ("ebc", "sav", "total")},
        }

    # ------------------------------------------------------------------ Articles à épuisement (Achat/Appro)
    def articles_epuisement(self, year: int, month: int) -> Dict[str, Optional[float]]:
        df = self.dfs["Ach_Arti"]
        epuis = df[df["Réappro"] == ARTICLES_EPUISEMENT_REAPPRO]
        total = len(epuis)
        col_ventes = f"Vtes {year}"
        ventes = int((epuis[col_ventes].fillna(0) != 0).sum())
        dispo = int((epuis["Dispo"].fillna(0) != 0).sum())

        def pct(n: int) -> Optional[float]:
            return (n / total * 100) if total else None

        return {"total": total, "ventes": ventes, "dispo": dispo,
                "pct_ventes": pct(ventes), "pct_dispo": pct(dispo)}

    # ------------------------------------------------------------------ Valorisation du stock (Achat/Appro)
    def valorisation_stock(self, year: int, month: int) -> Dict[str, object]:
        df = self.dfs["Stock_Invent"]
        col_t = next(c for c in df.columns if isinstance(c, str) and c.startswith("T "))
        sub = df[df["Dépot"].isin(STOCK_DEPOTS) & df["Famille"].isin([valeur for _, valeur in STOCK_FAMILLES])]
        piv = (sub.pivot_table(index="Famille", columns="Dépot", values=col_t, aggfunc="sum")
                  .reindex(index=[valeur for _, valeur in STOCK_FAMILLES], columns=STOCK_DEPOTS)
                  .fillna(0.0))

        rows, total_depot, total_showroom = [], 0.0, 0.0
        for label, valeur in STOCK_FAMILLES:
            depot, showroom = float(piv.loc[valeur, "DEPOT"]), float(piv.loc[valeur, "SHOWROOM"])
            rows.append({"famille": label, "depot": depot, "showroom": showroom, "total": depot + showroom})
            total_depot += depot
            total_showroom += showroom

        return {"rows": rows, "total_depot": total_depot, "total_showroom": total_showroom,
                "total": total_depot + total_showroom, "colonne": col_t}

    # ------------------------------------------------------------------ Assemblage
    def compute(self, year: int, month: int, mois1_reference: Optional[dict] = None) -> Dict[str, dict]:
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
        data["cur"]["productivite"] = self.productivite(year, month)
        data["cur"]["nb_interventions"] = self.nb_interventions(year, month)
        data["cur"]["valorisation_stock"] = self.valorisation_stock(year, month)
        data["cur"]["articles_epuisement"] = self.articles_epuisement(year, month)

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

        # Nombre d'interventions : pas d'historique recalculable, Mois-1 lu depuis un PPTX si fourni.
        m1_interventions = (mois1_reference or {}).get("nb_interventions", {})
        for champ in ("total", "ebc", "sav"):
            evolutions[f"nb_interventions_{champ}_m1"] = evolution(cur["nb_interventions"][champ], m1_interventions.get(champ))

        return {"year": year, "month": month, "periods": periods, "data": data, "evolutions": evolutions,
                "mois1_reference": mois1_reference}
