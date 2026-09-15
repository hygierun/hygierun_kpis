"""Utilitaires pour formatage et calculs"""

import pandas as pd
from datetime import datetime, timedelta
from typing import Optional, Tuple


class DateHelper:
    """Gère les périodes et filtres de dates"""

    @staticmethod
    def get_month_period(year: int, month: int) -> Tuple[datetime, datetime]:
        """Retourne (date_debut, date_fin) d'un mois"""
        start = datetime(year, month, 1)
        if month == 12:
            end = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            end = datetime(year, month + 1, 1) - timedelta(days=1)
        return start, end

    @staticmethod
    def get_week_period(date: datetime) -> Tuple[datetime, datetime]:
        """Retourne (lundi, dimanche) de la semaine d'une date"""
        start = date - timedelta(days=date.weekday())
        end = start + timedelta(days=6)
        return start, end

    @staticmethod
    def filter_by_date(df: pd.DataFrame, column: str, start_date: datetime,
                       end_date: datetime) -> pd.DataFrame:
        """Filtre un DataFrame par période de dates"""
        df_copy = df.copy()
        df_copy[column] = pd.to_datetime(df_copy[column], errors='coerce')
        return df_copy[(df_copy[column] >= start_date) & (df_copy[column] <= end_date)]


class FormatHelper:
    """Formate les valeurs pour les rapports"""

    @staticmethod
    def format_currency(value: float, symbol: str = "€") -> str:
        """Formate une valeur en devise"""
        if pd.isna(value):
            return "N/A"
        if value >= 1000:
            return f"{value/1000:.1f} k{symbol}"
        return f"{value:.0f} {symbol}"

    @staticmethod
    def format_percentage(value: float, decimals: int = 1) -> str:
        """Formate une valeur en pourcentage"""
        if pd.isna(value):
            return "N/A"
        return f"{value:.{decimals}f}%"

    @staticmethod
    def format_number(value: float, decimals: int = 0) -> str:
        """Formate un nombre"""
        if pd.isna(value):
            return "N/A"
        return f"{value:.{decimals}f}"

    @staticmethod
    def format_unit(value: float, unit: str = "", decimals: int = 1) -> str:
        """Formate une valeur avec unité"""
        if pd.isna(value):
            return "N/A"
        return f"{value:.{decimals}f} {unit}".strip()

    @staticmethod
    def get_trend_symbol(current: float, previous: float) -> Tuple[str, str]:
        """
        Retourne (symbole, couleur) pour une tendance

        Returns:
            ("▲ +10%", "green") ou ("▼ -5%", "red")
        """
        if pd.isna(current) or pd.isna(previous) or previous == 0:
            return "—", "gray"

        change = ((current - previous) / abs(previous)) * 100

        if change > 0:
            return f"▲ +{change:.0f}%", "green"
        elif change < 0:
            return f"▼ {change:.0f}%", "red"
        else:
            return "= 0%", "gray"


class KPICalculator:
    """Calculs de base pour les KPIs"""

    @staticmethod
    def count_records(df: pd.DataFrame, filter_col: Optional[str] = None,
                      filter_value: Optional[str] = None) -> int:
        """Compte les enregistrements avec filtrage optionnel"""
        if filter_col and filter_value:
            return len(df[df[filter_col] == filter_value])
        return len(df)

    @staticmethod
    def sum_column(df: pd.DataFrame, column: str) -> float:
        """Somme une colonne"""
        return df[column].sum() if column in df.columns else 0

    @staticmethod
    def average_column(df: pd.DataFrame, column: str) -> float:
        """Moyenne d'une colonne"""
        return df[column].mean() if column in df.columns else 0

    @staticmethod
    def percentage_of_total(value: float, total: float) -> float:
        """Calcule le pourcentage par rapport au total"""
        if total == 0:
            return 0
        return (value / total) * 100

    @staticmethod
    def calculate_difference(current: float, previous: float) -> Tuple[float, float]:
        """
        Calcule la différence et le pourcentage

        Returns:
            (difference_absolue, pourcentage)
        """
        diff = current - previous
        if previous == 0:
            pct = 0
        else:
            pct = (diff / abs(previous)) * 100
        return diff, pct
