"""Data Loader - Ingestion et préparation des données"""

import pandas as pd
import os
from datetime import datetime
from typing import Dict, Optional
from pathlib import Path

from src.config import EXCEL_SHEETS_CONFIG
from src.validators import DataValidator


class DataLoader:
    """Charge et prépare les données d'entrée"""

    def __init__(self, excel_path: str, sav_zip_path: Optional[str] = None):
        """
        Initialise le loader

        Args:
            excel_path: Chemin vers Input_Donnees_Brutes.xlsx
            sav_zip_path: Chemin optionnel vers FIN_AOUT_FICHES_SAV.zip
        """
        self.excel_path = excel_path
        self.sav_zip_path = sav_zip_path
        self.data = {}  # Dictionnaire de DataFrames
        self.validator = DataValidator()

    def load_and_validate(self) -> bool:
        """
        Charge et valide les données

        Returns:
            True si la charge est réussie, False sinon
        """
        print("\n" + "=" * 80)
        print("📂 CHARGEMENT DES DONNÉES")
        print("=" * 80)

        # Valider le fichier
        print(f"\n🔍 Validation du fichier: {os.path.basename(self.excel_path)}")
        is_valid, errors, warnings = self.validator.validate_excel_file(self.excel_path)

        if not is_valid:
            self.validator.print_validation_report(False)
            return False

        # Charger les feuilles requises
        print("\n📥 Chargement des feuilles Excel...")
        try:
            for sheet_name in EXCEL_SHEETS_CONFIG.keys():
                try:
                    df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
                    self.data[sheet_name] = self._prepare_dataframe(df, sheet_name)
                    print(f"   ✓ {sheet_name:<25} | {len(df):>6} lignes")
                except Exception as e:
                    print(f"   ❌ {sheet_name:<25} | Erreur: {str(e)}")
                    return False

        except Exception as e:
            print(f"❌ Erreur lors du chargement: {str(e)}")
            return False

        # Afficher le rapport de validation
        self.validator.print_validation_report(is_valid)

        print(f"\n✅ Données chargées avec succès!")
        print(f"   Total de feuilles chargées: {len(self.data)}")

        return True

    def _prepare_dataframe(self, df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
        """
        Prépare un DataFrame pour le traitement

        - Nettoie les noms de colonnes
        - Convertit les types de données
        - Supprime les lignes vides
        """
        # Supprimer les lignes complètement vides
        df = df.dropna(how='all')

        # Convertir les colonnes de dates
        date_columns = ['Date', 'Liv. souhaitée', 'Tournée', 'Echéance', 'RDate']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        # Convertir les colonnes numériques
        numeric_columns = [col for col in df.columns if 'Total' in col or 'Nb' in col
                          or '%' in col or 'Restant' in col or 'HT' in col or 'TTC' in col]
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        return df.reset_index(drop=True)

    def get_dataframe(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """Récupère un DataFrame chargé"""
        return self.data.get(sheet_name)

    def get_all_data(self) -> Dict[str, pd.DataFrame]:
        """Récupère tous les DataFrames"""
        return self.data.copy()

    # ========================================================================
    # MÉTHODES DE FILTRAGE
    # ========================================================================

    def filter_by_month(self, year: int, month: int) -> 'DataLoader':
        """Filtre tous les DataFrames par mois"""
        from datetime import datetime
        from src.utils import DateHelper

        start_date, end_date = DateHelper.get_month_period(year, month)

        for sheet_name, df in self.data.items():
            if 'Date' in df.columns:
                self.data[sheet_name] = DateHelper.filter_by_date(
                    df, 'Date', start_date, end_date
                )

        return self

    def filter_by_week(self, date: datetime) -> 'DataLoader':
        """Filtre tous les DataFrames par semaine"""
        from src.utils import DateHelper

        start_date, end_date = DateHelper.get_week_period(date)

        for sheet_name, df in self.data.items():
            if 'Date' in df.columns:
                self.data[sheet_name] = DateHelper.filter_by_date(
                    df, 'Date', start_date, end_date
                )

        return self

    # ========================================================================
    # STATS ET RAPPORTS
    # ========================================================================

    def print_data_summary(self):
        """Affiche un résumé des données chargées"""
        print("\n" + "=" * 80)
        print("📊 RÉSUMÉ DES DONNÉES")
        print("=" * 80)

        for sheet_name, df in self.data.items():
            print(f"\n{sheet_name:<25} | {len(df):>6} lignes | {len(df.columns):>2} colonnes")
            if len(df) > 0 and 'Date' in df.columns:
                min_date = df['Date'].min()
                max_date = df['Date'].max()
                print(f"{'':25} | Période: {min_date.date()} à {max_date.date()}")

        print("\n" + "=" * 80)


class SAVFichesParser:
    """Parser pour les fiches d'intervention SAV en ZIP"""

    def __init__(self, zip_path: str):
        """
        Initialise le parser SAV

        Args:
            zip_path: Chemin vers FIN_AOUT_FICHES_SAV.zip
        """
        self.zip_path = zip_path
        self.fiches_data = []

    def parse_fiches(self) -> bool:
        """
        Parse les fiches d'intervention du ZIP

        Pour l'instant, enregistre juste les métadonnées des PDFs.
        L'extraction du contenu texte des PDFs est à implémenter en Phase 3.

        Returns:
            True si succès
        """
        import zipfile

        try:
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                pdf_files = [f for f in zip_ref.namelist() if f.endswith('.pdf')]

                print(f"\n📑 Fiches d'intervention SAV trouvées: {len(pdf_files)}")

                for pdf_file in pdf_files:
                    info = zip_ref.getinfo(pdf_file)
                    self.fiches_data.append({
                        'filename': pdf_file,
                        'size': info.file_size,
                        'date': info.date_time
                    })

                print(f"✅ {len(self.fiches_data)} fiches traitées (métadonnées)")
                return True

        except Exception as e:
            print(f"❌ Erreur lors du parsing du ZIP SAV: {str(e)}")
            return False

    def get_fiches_count(self) -> int:
        """Retourne le nombre de fiches traitées"""
        return len(self.fiches_data)
