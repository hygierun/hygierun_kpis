"""KPI mensuels des sections Livraison et Comptabilité (règles validées avec Antoine, mode TCD)."""

from typing import Dict, List, Optional

import pandas as pd

from .monthly_commerce import evolution, month_bounds
from .monthly_loader import previous_month

TOURNEE_CAMION_MIN = 2200
TOURNEES_EXCLUES = (0, 524)
TOURNEE_ENLEVEMENT, TRANSPORTEUR_ENLEVEMENT = 327, "ENLEVEMENT CLIENTS"
TOURNEE_COMMERCIAUX, TRANSPORTEUR_COMMERCIAUX = 100, "LIVRAISON VRP"

DELAI_MAX_JOURS = 14
MULTI_BL_LIGNES = (0, 1, 2, 3)  # + "4 et plus"

QUALIFICATIONS_BLOQUEES = [
    "Client Public - Bloqué", "ND Cover - Bloqué", "ND Cover refusé - Bloqué", "Sans assurance - Bloqué",
]
IMPAYES_SEUIL_JOURS = 60

# GPS : plaque -> chauffeur (ordre d'affichage du tableau du template)
GPS_CHAUFFEURS = ["Nathan", "Patrick", "Alain (cellule)"]
GPS_PLAQUES = {"HE-451-KX": "Nathan", "GA-850-JB": "Patrick", "GS-993-QR": "Alain (cellule)"}

# Référence manuelle pour les KPI "photo du jour" (clients bloqués, factures dues, GPS) qu'on ne peut pas
# recalculer depuis l'historique du fichier. À défaut de mieux, ces valeurs viennent du dernier deck PowerPoint
# produit pour ce mois — jamais des KPI recalculables (délais, camions, enlèvements, multiples), où notre propre
# calcul sur les données du mois reste la référence fiable et cohérente avec les règles validées.
MOIS1_MANUEL: Dict[tuple, dict] = {
    (2026, 7): {
        "clients_bloques": {"total": 148, "actifs": 73, "ca": 157_000, "solde": None},
        "factures_dues": {"nb": 416, "montant": 227_985, "nb_60": 83, "montant_60": 39_000, "clients_60": 51},
        "gps": {
            "Nathan": {"Nb Arrêts Moy": 8.8, "Distance Moy (kms)": 110, "Nb de jours travail": 21},
            "Patrick": {"Nb Arrêts Moy": 10.0, "Distance Moy (kms)": 126, "Nb de jours travail": 21},
            "Alain (cellule)": {"Nb Arrêts Moy": 9.8, "Distance Moy (kms)": 147, "Nb de jours travail": 22},
            "Moy pond": {"Nb Arrêts Moy": 9.5, "Distance Moy (kms)": 128, "Nb de jours travail": 21.3},
        },
    },
}


# Pause déjeuner : arrêt de 11h à 14h d'au moins 30 min, sans plafond de durée (souvent près du dépôt Hygierun).
# Toute pause de plus d'1h30, à n'importe quelle heure, est aussi exclue (ce n'est pas un arrêt client).
PAUSE_HEURE_MIN, PAUSE_HEURE_MAX = 11.0, 14.0
PAUSE_DUREE_MIN = pd.Timedelta(minutes=30)
PAUSE_DUREE_LONGUE = pd.Timedelta(minutes=90)


def categorie_tournee(camion, designation) -> str:
    """Camion d'abord ; si vide, on lit la désignation (microstor / transp ext)."""
    cam = "" if pd.isna(camion) else str(camion).upper().strip()
    des = "" if pd.isna(designation) else str(designation).upper()
    if "MICRO" in cam:
        return "Microstor"
    if not cam:
        if "MICRO" in des.replace(" ", ""):
            return "Microstor"
        if "TRANSP EXT" in des:
            return "Sous-traitance externe"
    return "Camion"


class LivraisonComptaCalculator:
    """Calcule les KPI Livraison (avec N-1 et Mois-1) et Comptabilité (photo du jour)."""

    def __init__(self, dfs: Dict[str, pd.DataFrame], jours_recalage: int = 0):
        """jours_recalage : retire N jours à Nb JEch de Fact_Dues (utile uniquement pour rejouer un export ancien)."""
        self.dfs = dfs
        self.jours_recalage = jours_recalage

    @staticmethod
    def _in_month(df: pd.DataFrame, col: str, year: int, month: int) -> pd.Series:
        start, end = month_bounds(year, month)
        return (df[col] >= start) & (df[col] <= end)

    # ------------------------------------------------------------------ Tournées (camions)
    def tournees_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Tournees"]
        rows = df[self._in_month(df, "Date", year, month)].copy()
        rows["Catégorie"] = [categorie_tournee(c, d) for c, d in zip(rows["Camion"], rows["Désignation"])]
        return rows

    def livraisons_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Livr_Arch"].copy()
        df["Tournée"] = pd.to_numeric(df["Tournée"], errors="coerce")
        return df[self._in_month(df, "Date", year, month)]

    def livraison_camions(self, year: int, month: int) -> Dict[str, float]:
        tournees = self.tournees_rows(year, month)
        livraisons = self.livraisons_rows(year, month)
        ca_livre = tournees["Total HT"].sum()
        ca_total = livraisons[livraisons["Tournée"].notna() & ~livraisons["Tournée"].isin(TOURNEES_EXCLUES)]["Total HT"].sum()
        categories = tournees["Catégorie"].value_counts()
        return {
            "nb_tournees": len(tournees),
            "ca_livre": ca_livre,
            "ca_livre_total": ca_total,
            "part_ca_livre": ca_livre / ca_total * 100 if ca_total else None,
            "bla_moy": tournees["Nb bl A"].mean() if len(tournees) else None,
            "fact_moy": tournees["Nb fact"].mean() if len(tournees) else None,
            "microstor": int(categories.get("Microstor", 0)),
            "sous_traitance": int(categories.get("Sous-traitance externe", 0)),
        }

    def enlevements_commerciaux_rows(self, year: int, month: int) -> pd.DataFrame:
        rows = self.livraisons_rows(year, month).copy()
        transporteur = rows["Transporteur"].fillna("").astype(str).str.upper().str.strip()
        tournee = rows["Tournée"]
        rows["Catégorie"] = None
        rows.loc[(tournee == TOURNEE_ENLEVEMENT) | ((tournee == 0) & (transporteur == TRANSPORTEUR_ENLEVEMENT)), "Catégorie"] = "Enlèvement client"
        rows.loc[(tournee == TOURNEE_COMMERCIAUX) | ((tournee == 0) & (transporteur == TRANSPORTEUR_COMMERCIAUX)), "Catégorie"] = "Livraison commercial"
        return rows[rows["Catégorie"].notna()]

    def enlevements_commerciaux(self, year: int, month: int) -> Dict[str, int]:
        count = self.enlevements_commerciaux_rows(year, month)["Catégorie"].value_counts()
        return {"enlevements": int(count.get("Enlèvement client", 0)), "commerciaux": int(count.get("Livraison commercial", 0))}

    # ------------------------------------------------------------------ Délais (colonnes reconstruites depuis les dates)
    def delais_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Delais Bis"].copy()
        df["Tournée"] = pd.to_numeric(df["Tournée"], errors="coerce")
        df = df[self._in_month(df, "Date", year, month) & (df["Tournée"] > TOURNEE_CAMION_MIN)].copy()
        souhaite = (df["Liv. souhaitée"] - df["Date Creation Cde"]).dt.days
        df["Délai souhaité (j)"] = souhaite
        df["Délai cde → livraison (j)"] = (df["Date"] - df["Date Creation Cde"]).dt.days - (souhaite - 1)
        df["Délai livraison → facture (j)"] = (df["Fact date"] - df["Date"]).dt.days
        return df[df["Délai cde → livraison (j)"].between(0, DELAI_MAX_JOURS)]

    def delais(self, year: int, month: int) -> Dict[str, Optional[float]]:
        rows = self.delais_rows(year, month)
        return {
            "cde_livraison": rows["Délai cde → livraison (j)"].mean() if len(rows) else None,
            "livraison_facture": rows["Délai livraison → facture (j)"].mean() if rows["Délai livraison → facture (j)"].notna().any() else None,
            "nb": len(rows),
        }

    # ------------------------------------------------------------------ Multiples livraisons (toutes les commandes archivées)
    def multiples_rows(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Cdes_Arch"].copy()
        df["Tournée"] = pd.to_numeric(df["Tournée"], errors="coerce")
        return df[self._in_month(df, "Livraison", year, month) & df["Tournée"].notna() & ~df["Tournée"].isin(TOURNEES_EXCLUES)]

    def multiples(self, year: int, month: int) -> Dict[str, float]:
        rows = self.multiples_rows(year, month)
        total = len(rows)
        nb_bls = rows["Nb bls"]
        buckets = {str(n): int((nb_bls == n).sum()) for n in MULTI_BL_LIGNES}
        buckets["4 et plus"] = int((nb_bls >= 4).sum())
        return {
            "total": total, "buckets": buckets,
            "nb_1bl": buckets["1"], "part_1bl": buckets["1"] / total * 100 if total else None,
        }

    # ------------------------------------------------------------------ GPS (nb d'arrêts et distance par tournée = véhicule + jour)
    def gps_tournees_rows(self, year: int, month: int) -> pd.DataFrame:
        """Une ligne par (véhicule, date) du mois avec son nombre d'arrêts et sa distance parcourue.

        Un arrêt à Hygierun (retour dépôt, ou passage en cours de tournée) n'est jamais compté comme un arrêt.
        Une pause (11h-14h et >= 30 min, ou n'importe quelle heure si > 1h30) est exclue du nombre d'arrêts
        mais son kilométrage reste dans la distance totale, comme le trajet retour.
        """
        if "GPS_Livr" not in self.dfs:
            return pd.DataFrame(columns=["Véhicule", "Date", "Chauffeur", "Nb arrêts", "Distance (km)"])

        df = self.dfs["GPS_Livr"].copy()
        df = df[self._in_month(df, "Date", year, month)]
        df["Chauffeur"] = df["Véhicule"].apply(lambda v: next((n for p, n in GPS_PLAQUES.items() if p in str(v)), None))
        heure_arrivee = pd.to_datetime(df["Arrivée"], format="%H:%M:%S", errors="coerce")
        heure_decimale = heure_arrivee.dt.hour + heure_arrivee.dt.minute / 60
        duree_arret = pd.to_timedelta(df["Arrêt"].astype(str), errors="coerce")
        est_hygierun = df["Carnet arrivée"] == "HYGIERUN"
        est_pause = ~est_hygierun & (
            (heure_decimale.between(PAUSE_HEURE_MIN, PAUSE_HEURE_MAX) & (duree_arret >= PAUSE_DUREE_MIN))
            | (duree_arret > PAUSE_DUREE_LONGUE)
        )

        return (df.assign(**{"Arrêt commercial": ~est_hygierun & ~est_pause})
                  .groupby(["Véhicule", "Date"])
                  .apply(lambda g: pd.Series({
                      "Chauffeur": g["Chauffeur"].iloc[0],
                      "Nb arrêts": int(g["Arrêt commercial"].sum()),
                      "Distance (km)": g["Km"].sum(),
                  }))
                  .reset_index())

    def gps_par_chauffeur(self, year: int, month: int) -> Optional[pd.DataFrame]:
        """Nb d'arrêts moyen, distance moyenne et nb de jours travaillés par chauffeur, + une ligne 'Moy pond'."""
        rows = self.gps_tournees_rows(year, month)
        if rows.empty:
            return None

        par_chauffeur = rows.groupby("Chauffeur").agg(
            **{"Nb Arrêts Moy": ("Nb arrêts", "mean"), "Distance Moy (kms)": ("Distance (km)", "mean"),
               "Nb de jours travail": ("Date", "count")}
        ).reindex(GPS_CHAUFFEURS)

        pondere = pd.DataFrame([{
            "Nb Arrêts Moy": rows["Nb arrêts"].mean(), "Distance Moy (kms)": rows["Distance (km)"].mean(),
            "Nb de jours travail": rows["Date"].nunique(),
        }], index=["Moy pond"])
        return pd.concat([par_chauffeur, pondere])

    # ------------------------------------------------------------------ Comptabilité (photo du jour de l'export)
    def clients_bloques_rows(self, year: int) -> pd.DataFrame:
        df = self.dfs["Clients"]
        return df[df["Qualification"].isin(QUALIFICATIONS_BLOQUEES)]

    def clients_bloques(self, year: int) -> Dict[str, float]:
        bloques = self.clients_bloques_rows(year)
        actifs = bloques[bloques[f"Vtes {year}"] > 0]
        return {
            "total": len(bloques), "actifs": len(actifs),
            "ca": actifs[f"Vtes {year}"].sum(), "solde": actifs["Solde cpta"].sum(),
        }

    def _factures_dues(self, year: int, month: int) -> pd.DataFrame:
        df = self.dfs["Fact_Dues"].copy()
        _, end = month_bounds(year, month)
        df = df[df["Date"] <= end]
        df["Nb JEch retenu"] = df["Nb JEch"] - self.jours_recalage
        return df

    def factures_dues_rows(self, year: int, month: int) -> pd.DataFrame:
        """Factures impayées (Nb JEch > 0, montant >= 0) avec un indicateur > 60 j (Nb JEch >= 60, montant > 0)."""
        df = self._factures_dues(year, month)
        impayees = df[(df["Nb JEch retenu"] > 0) & (df["Restant dû"] >= 0)].copy()
        impayees[f"> {IMPAYES_SEUIL_JOURS} j"] = (impayees["Nb JEch retenu"] >= IMPAYES_SEUIL_JOURS) & (impayees["Restant dû"] > 0)
        return impayees

    def factures_dues(self, year: int, month: int) -> Dict[str, float]:
        impayees = self.factures_dues_rows(year, month)
        anciennes = impayees[impayees[f"> {IMPAYES_SEUIL_JOURS} j"]]
        de_l_annee = impayees[impayees["Date"].dt.year == year]
        nb, montant = len(impayees), impayees["Restant dû"].sum()
        return {
            "nb": nb, "montant": montant,
            "part_annee_nb": len(de_l_annee) / nb * 100 if nb else None,
            "part_annee_montant": de_l_annee["Restant dû"].sum() / montant * 100 if montant else None,
            "nb_60": len(anciennes), "montant_60": anciennes["Restant dû"].sum(),
            "clients_60": int(anciennes["Client (réf.)"].nunique()),
            "part_60_nb": len(anciennes) / nb * 100 if nb else None,
            "part_60_montant": anciennes["Restant dû"].sum() / montant * 100 if montant else None,
        }

    # ------------------------------------------------------------------ Assemblage
    def compute(self, year: int, month: int) -> Dict[str, dict]:
        py, pm = previous_month(year, month)
        periods = {"cur": (year, month), "n1": (year - 1, month), "m1": (py, pm)}
        data = {
            key: {
                "camions": self.livraison_camions(y, m),
                "enlevements": self.enlevements_commerciaux(y, m),
                "delais": self.delais(y, m),
                "multiples": self.multiples(y, m),
            }
            for key, (y, m) in periods.items()
        }
        data["cur"]["clients_bloques"] = self.clients_bloques(year)
        data["cur"]["factures_dues"] = self.factures_dues(year, month)
        data["cur"]["gps"] = self.gps_par_chauffeur(year, month)

        manuel = MOIS1_MANUEL.get((py, pm), {})
        data["m1"]["clients_bloques"] = manuel.get("clients_bloques")
        data["m1"]["factures_dues"] = manuel.get("factures_dues")
        data["m1"]["gps"] = pd.DataFrame(manuel["gps"]).T if "gps" in manuel else None

        cur = data["cur"]
        evolutions: Dict[str, Optional[float]] = {}
        m1_cb, m1_fd, m1_gps = data["m1"]["clients_bloques"], data["m1"]["factures_dues"], data["m1"]["gps"]
        if m1_cb:
            evolutions["clients_bloques_total_m1"] = evolution(cur["clients_bloques"]["total"], m1_cb["total"])
            evolutions["clients_bloques_ca_m1"] = evolution(cur["clients_bloques"]["ca"], m1_cb["ca"])
        if m1_fd:
            evolutions["factures_dues_nb_m1"] = evolution(cur["factures_dues"]["nb"], m1_fd["nb"])
            evolutions["factures_dues_montant_m1"] = evolution(cur["factures_dues"]["montant"], m1_fd["montant"])
            evolutions["factures_dues_60_nb_m1"] = evolution(cur["factures_dues"]["nb_60"], m1_fd["nb_60"])
            evolutions["factures_dues_60_montant_m1"] = evolution(cur["factures_dues"]["montant_60"], m1_fd["montant_60"])
        if m1_gps is not None:
            evolutions["gps_arrets_m1"] = evolution(cur["gps"].loc["Moy pond", "Nb Arrêts Moy"], m1_gps.loc["Moy pond", "Nb Arrêts Moy"])
            evolutions["gps_distance_m1"] = evolution(cur["gps"].loc["Moy pond", "Distance Moy (kms)"], m1_gps.loc["Moy pond", "Distance Moy (kms)"])

        for ref in ("n1", "m1"):
            other = data[ref]
            evolutions[f"ca_livre_{ref}"] = evolution(cur["camions"]["ca_livre"], other["camions"]["ca_livre"])
            evolutions[f"bla_{ref}"] = evolution(cur["camions"]["bla_moy"], other["camions"]["bla_moy"])
            evolutions[f"fact_{ref}"] = evolution(cur["camions"]["fact_moy"], other["camions"]["fact_moy"])
            evolutions[f"enlevements_{ref}"] = evolution(cur["enlevements"]["enlevements"], other["enlevements"]["enlevements"])
            evolutions[f"commerciaux_{ref}"] = evolution(cur["enlevements"]["commerciaux"], other["enlevements"]["commerciaux"])
            evolutions[f"delai_cde_livraison_{ref}"] = evolution(cur["delais"]["cde_livraison"], other["delais"]["cde_livraison"])
            evolutions[f"delai_livraison_facture_{ref}"] = evolution(cur["delais"]["livraison_facture"], other["delais"]["livraison_facture"])
            evolutions[f"part_1bl_{ref}"] = evolution(cur["multiples"]["part_1bl"], other["multiples"]["part_1bl"])

        return {"year": year, "month": month, "periods": periods, "data": data, "evolutions": evolutions}
