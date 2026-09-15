"""Validation de la qualité des données d'entrée"""

import pandas as pd
from typing import Dict, List, Tuple
from src.config import EXCEL_SHEETS_CONFIG


class DataValidator:
    """Valide la structure et la qualité des données Excel"""

    def __init__(self):
        self.errors = []
        self.warnings = []

    def validate_excel_file(self, excel_path: str) -> Tuple[bool, List[str], List[str]]:
        """
        Valide le fichier Excel complet

        Returns:
            (est_valide, erreurs, avertissements)
        """
        self.errors = []
        self.warnings = []

        try:
            # Lire les feuilles disponibles
            xls = pd.ExcelFile(excel_path)
            available_sheets = set(xls.sheet_names)

        except Exception as e:
            self.errors.append(f"❌ Impossible de lire le fichier Excel: {str(e)}")
            return False, self.errors, self.warnings

        # Vérifier les feuilles requises
        required_sheets = set(EXCEL_SHEETS_CONFIG.keys())
        missing_sheets = required_sheets - available_sheets

        if missing_sheets:
            self.errors.append(
                f"❌ Feuilles manquantes : {', '.join(sorted(missing_sheets))}"
            )

        # Valider chaque feuille
        for sheet_name in required_sheets & available_sheets:
            try:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                self._validate_sheet(sheet_name, df)
            except Exception as e:
                self.errors.append(
                    f"❌ Erreur lors de la lecture de '{sheet_name}': {str(e)}"
                )

        is_valid = len(self.errors) == 0

        return is_valid, self.errors, self.warnings

    def _validate_sheet(self, sheet_name: str, df: pd.DataFrame):
        """Valide une feuille spécifique"""

        config = EXCEL_SHEETS_CONFIG[sheet_name]
        required_columns = config["required_columns"]

        # Vérifier les colonnes
        missing_columns = set(required_columns) - set(df.columns)

        if missing_columns:
            self.errors.append(
                f"❌ Feuille '{sheet_name}': colonnes manquantes : "
                f"{', '.join(sorted(missing_columns))}"
            )
            return

        # Avertissements pour données manquantes (dans les colonnes clés)
        key_columns = required_columns[:3]  # Les 3 premières colonnes
        missing_values = df[key_columns].isnull().sum()

        if missing_values.any():
            for col, count in missing_values[missing_values > 0].items():
                pct = (count / len(df)) * 100
                if pct > 10:
                    self.warnings.append(
                        f"⚠️  Feuille '{sheet_name}': {count} valeurs manquantes "
                        f"dans '{col}' ({pct:.1f}%)"
                    )

    def print_validation_report(self, is_valid: bool):
        """Affiche un rapport de validation"""

        print("\n" + "=" * 80)
        print("📋 RAPPORT DE VALIDATION DES DONNÉES")
        print("=" * 80)

        if is_valid:
            print("✅ Validation réussie - Toutes les données sont correctes !")
        else:
            print("❌ Validation échouée - Veuillez corriger les erreurs suivantes:\n")
            for error in self.errors:
                print(f"   {error}")

        if self.warnings:
            print("\n⚠️  Avertissements:")
            for warning in self.warnings:
                print(f"   {warning}")

        print("\n" + "=" * 80)

        return is_valid
