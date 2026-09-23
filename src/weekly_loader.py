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

    # Colonnes réellement utilisées par KPICalculator pour chaque feuille - source de vérité pour le
    # template Input téléchargeable et l'aide affichée dans l'appli (pas de validation stricte ici,
    # WeeklyDataLoader ne vérifie que la présence des feuilles, pas des colonnes).
    REQUIRED_COLUMNS: Dict[str, list] = {
        "Commandes_ALivrer": ["Date", "Représentant", "Total HT", "Livraison", "Total TTC"],
        "Commandes_entournées": ["Date", "Représentant", "Total HT"],
        "Commandes_Arch": ["Date", "Représentant", "Total HT"],
        "Fact_Integrer": ["Date", "Représentant", "Total HT"],
        "Fact_Arch": ["Date", "Représentant", "Total HT"],
        "Livr_AFact": ["Date", "Tournée", "Représentant", "Total HT"],
        "Livr_Arch": ["Date", "Tournée", "Représentant", "Total HT"],
        "Expl_Loc_Delais": ["Date", "Date Creation Cde", "Liv. souhaitée", "Tournée"],
        "Creances": ["Nb JEch", "Restant dû"],
    }

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
                    # Gestion spéciale pour Expl_Loc_Delais
                    if sheet == "Expl_Loc_Delais":
                        df = pd.read_excel(self.file_path, sheet_name=sheet)
                        # Note: Si le fichier a des colonnes vides au début, les pandas les ignore automatiquement
                    else:
                        df = pd.read_excel(self.file_path, sheet_name=sheet)
                    self.dfs[sheet] = df
                    print(f"   ✓ {sheet:30s} ({len(df):6d} lignes)")
                else:
                    print(f"   ✗ {sheet:30s} (non trouvée)")

            # Vérifier qu'au moins les feuilles critiques EXISTENT (peuvent être vides pour données passées)
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

            missing = [s for s in critical_sheets if s not in self.available_sheets]
            if missing:
                print(f"\n❌ Feuilles critiques manquantes: {missing}")
                return False

            print(f"\n✅ Feuilles critiques présentes (peuvent être vides pour données archivées)")

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
