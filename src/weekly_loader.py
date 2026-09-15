"""
DataLoader simplifié pour fichiers hebdomadaires
Charge uniquement les feuilles essentielles requises pour les calculs KPI
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Optional


class WeeklyDataLoader:
    """Charge les données d'un fichier Excel hebdomadaire"""

    # Feuilles essentielles pour une extraction hebdomadaire
    REQUIRED_SHEETS = [
        "Commandes_ALivrer",
        "Commandes_entournées",
        "Commandes_Arch",
        "Fact_Integrer",
        "Fact_Arch",
        "Livr_AFact",
        "Livr_Arch",
        "Expl_Loc_Delais",
        "Creances",
    ]

    def __init__(self, file_path: str):
        """
        Initialise le loader

        Args:
            file_path: Chemin du fichier Excel
        """
        self.file_path = file_path
        self.dfs: Dict[str, pd.DataFrame] = {}
        self.available_sheets = []

    def load_and_validate(self) -> bool:
        """
        Charge et valide les données

        Returns:
            True si succès, False sinon
        """

        if not Path(self.file_path).exists():
            print(f"❌ Fichier non trouvé: {self.file_path}")
            return False

        print(f"\n📂 CHARGEMENT DES DONNÉES")
        print("=" * 80)
        print(f"📁 Fichier: {self.file_path}")

        try:
            # Lister les feuilles disponibles
            xl = pd.ExcelFile(self.file_path)
            self.available_sheets = xl.sheet_names
            print(f"\n🔍 Feuilles disponibles: {len(self.available_sheets)}")
            for sheet in self.available_sheets:
                print(f"   ✓ {sheet}")

            # Charger les feuilles essentielles
            print(f"\n📥 Chargement des feuilles essentielles:")
            for sheet in self.REQUIRED_SHEETS:
                if sheet in self.available_sheets:
                    df = pd.read_excel(self.file_path, sheet_name=sheet)
                    self.dfs[sheet] = df
                    print(f"   ✓ {sheet:30s} ({len(df):6d} lignes)")
                else:
                    print(f"   ✗ {sheet:30s} (non trouvée)")

            # Vérifier qu'au moins les feuilles critiques sont présentes
            critical_sheets = [
                "Commandes_ALivrer",
                "Commandes_Arch",
                "Fact_Integrer",
                "Fact_Arch",
                "Livr_AFact",
                "Livr_Arch",
                "Expl_Loc_Delais",
                "Creances",
            ]

            missing = [s for s in critical_sheets if s not in self.dfs]
            if missing:
                print(f"\n❌ Feuilles critiques manquantes: {missing}")
                return False

            print(f"\n✅ Données chargées avec succès")
            return True

        except Exception as e:
            print(f"\n❌ Erreur lors du chargement: {e}")
            return False

    def get_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """Retourne une feuille chargée"""
        return self.dfs.get(sheet_name)

    def print_data_summary(self) -> None:
        """Affiche un résumé des données chargées"""
        print(f"\n📊 RÉSUMÉ DES DONNÉES")
        print("=" * 80)
        for sheet, df in self.dfs.items():
            print(f"\n{sheet}:")
            print(f"   Lignes: {len(df)}")
            print(f"   Colonnes: {len(df.columns)}")
            if len(df.columns) > 0:
                print(f"   Champs: {', '.join(df.columns[:5])}...")
