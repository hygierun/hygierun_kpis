"""Hygierun KPI Tool - Automated KPI calculation and report generation"""

from .weekly_loader import WeeklyDataLoader
from .kpi_calculator import KPICalculator

__version__ = "1.0.0"
__author__ = "Hygierun Analytics"

__all__ = ["WeeklyDataLoader", "KPICalculator"]
