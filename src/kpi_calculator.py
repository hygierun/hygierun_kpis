"""
Moteur de calcul des KPIs Hygierun - Phase 3

Implémente tous les KPIs validés avec les formules exactes.
Chaque KPI suit la recette précise documentée et validée avec Antoine.
"""

import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, Optional, Tuple, List


class KPICalculator:
    """Calculateur de KPIs pour Hygierun"""

    # Noms des commerciaux (5 sans Franck)
    SALES_REPS_5 = [
        "HUGUET Anthony",
        "STRAZZULLA Francesco",
        "QUERE Julien",
        "SUMATRA Alexandre",
        "CAZAL Damien"
    ]

    # 6 commerciaux (avec Franck)
    SALES_REPS_6 = SALES_REPS_5 + ["LECLANCHER Franck"]

    def __init__(self, data_loader):
        """
        Initialise le calculateur avec les DataFrames chargés

        Args:
            data_loader: Instance de DataLoader avec toutes les feuilles chargées
        """
        self.loader = data_loader
        self.dfs = data_loader.dfs
        self.start_date = None
        self.end_date = None

    def set_period(self, start_date: datetime, end_date: datetime):
        """Définit la période de calcul"""
        self.start_date = start_date
        self.end_date = end_date

    # ========================================================================
    # SECTION COMMANDES
    # ========================================================================

    def calc_commandes_avec_franck(self) -> Tuple[int, float]:
        """
        Commandes avec Franck
        Source: Commandes_ALivrer + Commandes_entournées + Commandes_Arch
        Filtres: Date période + Représentant (6 commerciaux + vides)
        """
        dfs = []
        for sheet in ["Commandes_ALivrer", "Commandes_entournées", "Commandes_Arch"]:
            if sheet in self.dfs:
                dfs.append(self.dfs[sheet].copy())

        if not dfs:
            return 0, 0

        df = pd.concat(dfs, ignore_index=True)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

        # Filtre date
        df = df[(df['Date'] >= self.start_date) & (df['Date'] <= self.end_date)]

        # Filtre représentants: 6 + vides
        df = df[(df['Représentant'].isin(self.SALES_REPS_6)) | (df['Représentant'].isna())]

        nb = len(df)
        ca = df['Total HT'].sum() if 'Total HT' in df.columns else 0

        return nb, ca

    def calc_commandes_sans_franck(self) -> Tuple[int, float]:
        """Commandes sans Franck (SEULEMENT les 5, pas de vides)"""
        dfs = []
        for sheet in ["Commandes_ALivrer", "Commandes_entournées", "Commandes_Arch"]:
            if sheet in self.dfs:
                dfs.append(self.dfs[sheet].copy())

        if not dfs:
            return 0, 0

        df = pd.concat(dfs, ignore_index=True)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

        # Filtre date
        df = df[(df['Date'] >= self.start_date) & (df['Date'] <= self.end_date)]

        # Filtre représentants: 5 seulement (sans Franck, sans vides)
        df = df[df['Représentant'].isin(self.SALES_REPS_5)]

        nb = len(df)
        ca = df['Total HT'].sum() if 'Total HT' in df.columns else 0

        return nb, ca

    # ========================================================================
    # SECTION FACTURES
    # ========================================================================

    def calc_factures_avec_franck(self) -> Tuple[int, float]:
        """
        Factures avec Franck
        Source: Fact_Integrer + Fact_Arch
        Filtres: Date période + Représentant (6 + vides)
        """
        dfs = []
        for sheet in ["Fact_Integrer", "Fact_Arch"]:
            if sheet in self.dfs:
                dfs.append(self.dfs[sheet].copy())

        if not dfs:
            return 0, 0

        df = pd.concat(dfs, ignore_index=True)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

        # Filtre date
        df = df[(df['Date'] >= self.start_date) & (df['Date'] <= self.end_date)]

        # Filtre représentants: 6 + vides
        df = df[(df['Représentant'].isin(self.SALES_REPS_6)) | (df['Représentant'].isna())]

        nb = len(df)
        ca = df['Total HT'].sum() if 'Total HT' in df.columns else 0

        return nb, ca

    def calc_factures_sans_franck(self) -> Tuple[int, float]:
        """Factures sans Franck (5 seulement, sans vides)"""
        dfs = []
        for sheet in ["Fact_Integrer", "Fact_Arch"]:
            if sheet in self.dfs:
                dfs.append(self.dfs[sheet].copy())

        if not dfs:
            return 0, 0

        df = pd.concat(dfs, ignore_index=True)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

        # Filtre date
        df = df[(df['Date'] >= self.start_date) & (df['Date'] <= self.end_date)]

        # Filtre représentants: 5 seulement
        df = df[df['Représentant'].isin(self.SALES_REPS_5)]

        nb = len(df)
        ca = df['Total HT'].sum() if 'Total HT' in df.columns else 0

        return nb, ca

    # ========================================================================
    # SECTION LIVRAISONS
    # ========================================================================

    def calc_livraisons_avec_franck(self) -> Tuple[int, float]:
        """
        Livraisons avec Franck
        Source: Livr_AFact + Livr_Arch
        Filtres: Date + Tournée > 0 + Représentant (6 + vides)
        """
        dfs = []
        for sheet in ["Livr_AFact", "Livr_Arch"]:
            if sheet in self.dfs:
                dfs.append(self.dfs[sheet].copy())

        if not dfs:
            return 0, 0

        df = pd.concat(dfs, ignore_index=True)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df['Tournée'] = pd.to_numeric(df['Tournée'], errors='coerce')

        # Filtres
        df = df[(df['Date'] >= self.start_date) & (df['Date'] <= self.end_date)]
        df = df[df['Tournée'] > 0]
        df = df[(df['Représentant'].isin(self.SALES_REPS_6)) | (df['Représentant'].isna())]

        nb = len(df)
        ca = df['Total HT'].sum() if 'Total HT' in df.columns else 0

        return nb, ca

    def calc_livraisons_sans_franck(self) -> Tuple[int, float]:
        """Livraisons sans Franck (5 seulement)"""
        dfs = []
        for sheet in ["Livr_AFact", "Livr_Arch"]:
            if sheet in self.dfs:
                dfs.append(self.dfs[sheet].copy())

        if not dfs:
            return 0, 0

        df = pd.concat(dfs, ignore_index=True)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df['Tournée'] = pd.to_numeric(df['Tournée'], errors='coerce')

        # Filtres
        df = df[(df['Date'] >= self.start_date) & (df['Date'] <= self.end_date)]
        df = df[df['Tournée'] > 0]
        df = df[df['Représentant'].isin(self.SALES_REPS_5)]

        nb = len(df)
        ca = df['Total HT'].sum() if 'Total HT' in df.columns else 0

        return nb, ca

    def calc_delai_livraison(self) -> float:
        """
        Délai moyen Commande → Livraison
        Source: Expl_Loc_Delais

        Formule:
        - Calcule le délai souhaité à partir de (Liv. souhaitée - Date Creation Cde)
        - delai_cde_2_livr = (Date - Date Creation Cde) - (délai livraison souhaité - 1)

        Filtres:
        - Date [start_date; end_date]
        - délai livraison souhaité >= 0
        - delai_cde_2_livr [0; 14]
        - Tournée > 3000
        """
        df = self.dfs.get("Expl_Loc_Delais", pd.DataFrame()).copy()
        if df.empty:
            return 0

        # Convertir dates
        for col in ['Date', 'Date Creation Cde', 'Liv. souhaitée']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        # Convertir nombres
        df['Tournée'] = pd.to_numeric(df['Tournée'], errors='coerce')

        # Calculer le délai livraison souhaité (en jours)
        # délai livraison souhaité = (Liv. souhaitée - Date Creation Cde).days
        df['délai_livraison_souhaité'] = (
            df['Liv. souhaitée'] - df['Date Creation Cde']
        ).dt.days

        # Calculer le délai cde 2 livr
        # delai_cde_2_livr = (Date - Date Creation Cde) - (délai livraison souhaité - 1)
        df['delai_cde_2_livr'] = (
            (df['Date'] - df['Date Creation Cde']).dt.days -
            (df['délai_livraison_souhaité'] - 1)
        )

        # Appliquer filtres
        df = df[(df['Date'] >= self.start_date) & (df['Date'] <= self.end_date)]
        df = df[df['délai_livraison_souhaité'] >= 0]
        df = df[(df['delai_cde_2_livr'] >= 0) & (df['delai_cde_2_livr'] <= 14)]
        df = df[df['Tournée'] > 3000]

        if df.empty:
            return 0

        return df['delai_cde_2_livr'].mean()

    def calc_livraisons_en_attente(self) -> Tuple[int, float]:
        """
        Livraisons en attente
        Source: Commandes_ALivrer

        Filtres:
        - Livraison [01/01/2026; end_date]
        - Représentant: 6 commerciaux (sans vides)
        """
        df = self.dfs.get("Commandes_ALivrer", pd.DataFrame()).copy()
        if df.empty:
            return 0, 0

        df['Livraison'] = pd.to_datetime(df['Livraison'], errors='coerce')

        # Filtres
        date_min = datetime(2026, 1, 1)
        df = df[(df['Livraison'] >= date_min) & (df['Livraison'] <= self.end_date)]
        df = df[df['Représentant'].isin(self.SALES_REPS_6)]

        nb = len(df)
        ca = df['Total TTC'].sum() if 'Total TTC' in df.columns else df['Total HT'].sum()

        return nb, ca

    # ========================================================================
    # SECTION CRÉANCES
    # ========================================================================

    def calc_creances(self) -> Tuple[int, float]:
        """
        Créances échues
        Source: Creances
        Filtre: Nb JEch > 0 et pas vide
        """
        df = self.dfs.get("Creances", pd.DataFrame()).copy()
        if df.empty:
            return 0, 0

        df['Nb JEch'] = pd.to_numeric(df['Nb JEch'], errors='coerce')

        # Filtre
        df = df[(df['Nb JEch'] > 0) & (df['Nb JEch'].notna())]

        nb = len(df)
        ca = df['Restant dû'].sum() if 'Restant dû' in df.columns else 0

        return nb, ca

    # ========================================================================
    # UTILITAIRES
    # ========================================================================

    def get_all_kpis(self) -> Dict[str, any]:
        """
        Calcule et retourne tous les KPIs pour la période

        Returns:
            Dict avec tous les KPIs organisés par section
        """
        # Commandes
        nb_cdes_avec, ca_cdes_avec = self.calc_commandes_avec_franck()
        nb_cdes_sans, ca_cdes_sans = self.calc_commandes_sans_franck()

        # Factures
        nb_fact_avec, ca_fact_avec = self.calc_factures_avec_franck()
        nb_fact_sans, ca_fact_sans = self.calc_factures_sans_franck()

        # Livraisons
        nb_livr_avec, ca_livr_avec = self.calc_livraisons_avec_franck()
        nb_livr_sans, ca_livr_sans = self.calc_livraisons_sans_franck()
        delai = self.calc_delai_livraison()
        nb_attente, ca_attente = self.calc_livraisons_en_attente()

        # Créances
        nb_creances, ca_creances = self.calc_creances()

        return {
            "commandes": {
                "nb_avec": nb_cdes_avec,
                "ca_avec": ca_cdes_avec,
                "ca_sans": ca_cdes_sans,
            },
            "factures": {
                "nb_avec": nb_fact_avec,
                "ca_avec": ca_fact_avec,
                "nb_sans": nb_fact_sans,
                "ca_sans": ca_fact_sans,
            },
            "livraisons": {
                "nb_avec": nb_livr_avec,
                "ca_avec": ca_livr_avec,
                "nb_sans": nb_livr_sans,
                "ca_sans": ca_livr_sans,
                "delai": delai,
                "en_attente_nb": nb_attente,
                "en_attente_ca": ca_attente,
            },
            "creances": {
                "nb": nb_creances,
                "ca": ca_creances,
            }
        }

    def get_formatted_kpis(self) -> Dict[str, any]:
        """Retourne les KPIs formatés pour l'affichage et l'export"""
        kpis = self.get_all_kpis()

        return {
            "commandes": {
                "nb_avec": f"{kpis['commandes']['nb_avec']:.0f}",
                "ca_avec": f"{kpis['commandes']['ca_avec']/1000:.1f}",  # en k€
                "ca_sans": f"{kpis['commandes']['ca_sans']/1000:.1f}",
            },
            "factures": {
                "nb_avec": f"{kpis['factures']['nb_avec']:.0f}",
                "ca_avec": f"{kpis['factures']['ca_avec']/1000:.1f}",
                "nb_sans": f"{kpis['factures']['nb_sans']:.0f}",
                "ca_sans": f"{kpis['factures']['ca_sans']/1000:.1f}",
            },
            "livraisons": {
                "nb_avec": f"{kpis['livraisons']['nb_avec']:.0f}",
                "ca_avec": f"{kpis['livraisons']['ca_avec']/1000:.1f}",
                "nb_sans": f"{kpis['livraisons']['nb_sans']:.0f}",
                "ca_sans": f"{kpis['livraisons']['ca_sans']/1000:.1f}",
                "delai": f"{kpis['livraisons']['delai']:.2f}",
                "en_attente_nb": f"{kpis['livraisons']['en_attente_nb']:.0f}",
                "en_attente_ca": f"{kpis['livraisons']['en_attente_ca']/1000:.1f}",
            },
            "creances": {
                "nb": f"{kpis['creances']['nb']:.0f}",
                "ca": f"{kpis['creances']['ca']/1000:.1f}",
            }
        }
