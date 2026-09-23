"""Assemble les 4 Excel de sortie regroupés du Bilan Total (Commerce / Préparation+Livraison /
Compta+Achats / SAV), validés avec Antoine. Réutilise telles quelles les lignes de synthèse et les
feuilles de détail déjà en place pour chaque section (factorisées en fonctions composables dans
monthly_preparation_excel / monthly_livraison_excel / monthly_sav_excel) - aucune règle de calcul
dupliquée ou modifiée, seul l'assemblage change.
"""

from pathlib import Path

from openpyxl import Workbook

from . import monthly_preparation_excel
from .excel_helpers import write_summary
from .monthly_livraison import LivraisonComptaCalculator
from .monthly_livraison_excel import compta_rows, livraison_ops_rows, write_compta_sheets, write_livraison_ops_sheets
from .monthly_loader import MONTHS_FR
from .monthly_preparation import PreparationCalculator
from .monthly_sav import SavCalculator
from .monthly_sav_excel import achat_rows, sav_rows, write_achat_sheets, write_sav_sheets


def generate_preparation_livraison_excel(prep_calc: PreparationCalculator, prep_result: dict,
                                         livraison_calc: LivraisonComptaCalculator, livraison_result: dict,
                                         output_path: str) -> str:
    year, month = prep_result["year"], prep_result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    rows = monthly_preparation_excel.summary_rows(prep_result) + [None] + livraison_ops_rows(livraison_result)
    write_summary(ws, f"Préparation et Livraison - {MONTHS_FR[month - 1]} {year}", rows)

    monthly_preparation_excel.write_preparation_sheets(wb, prep_calc, year, month)
    write_livraison_ops_sheets(wb, livraison_calc, year, month)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def generate_compta_achats_excel(livraison_calc: LivraisonComptaCalculator, livraison_result: dict,
                                 sav_calc: SavCalculator, sav_result: dict, output_path: str) -> str:
    year, month = livraison_result["year"], livraison_result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    rows = compta_rows(livraison_result) + [None] + achat_rows(sav_result)
    write_summary(ws, f"Compta et Achats - {MONTHS_FR[month - 1]} {year}", rows)

    write_compta_sheets(wb, livraison_calc, year, month)
    write_achat_sheets(wb, sav_calc, year, month)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def generate_sav_only_excel(sav_calc: SavCalculator, sav_result: dict, output_path: str) -> str:
    year, month = sav_result["year"], sav_result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    write_summary(ws, f"SAV - {MONTHS_FR[month - 1]} {year}", sav_rows(sav_result))

    write_sav_sheets(wb, sav_calc, year, month)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
