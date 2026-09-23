"""Excel de sortie de la section Préparation : synthèse chiffrée + feuilles de détail."""

from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook

from .excel_helpers import select_columns, write_frame, write_summary
from .monthly_loader import MONTHS_FR
from .monthly_preparation import PreparationCalculator


def summary_rows(result: dict) -> List[Optional[list]]:
    """Lignes de la feuille Synthèse : [libellé, valeur, N-1, évol N-1, Mois-1, évol Mois-1]."""
    d, e = result["data"], result["evolutions"]
    cur, n1, m1 = d["cur"], d["n1"], d["m1"]

    def line(label: str, block: str, field: str, evo_key: str) -> list:
        return [label, cur[block][field], n1[block][field], e[f"{evo_key}_n1"], m1[block][field], e[f"{evo_key}_m1"]]

    return [
        ["Commandes préparées par la logistique (estimées sur les dates de livraison)"],
        line("Total commandes préparées", "commandes", "total", "total"),
        line("CA HT (€)", "commandes", "ca", "ca"),
        line("  dont cdes clients", "commandes", "clients", "clients"),
        line("  dont réassort (DEPOT SHOWROOM)", "commandes", "reassort", "reassort"),
        line("  dont SAV (virement de dépôt)", "commandes", "sav", "sav"),
        None,
        ["Dépotage conteneurs"],
        line("Nb conteneurs", "conteneurs", "nb", "conteneurs_nb"),
        line("Total HT réceptions conteneurs (€)", "conteneurs", "ca", "conteneurs_ca"),
    ]


def write_preparation_sheets(wb: Workbook, calc: PreparationCalculator, year: int, month: int):
    livraisons = select_columns(calc.commandes_preparees_rows(year, month),
                                ["Catégorie", "Date", "N°", "Tournée", "Client", "Total HT"])
    write_frame(wb, "Commandes_preparees", livraisons)
    write_frame(wb, "Conteneurs", select_columns(calc.conteneurs_rows(year, month),
                ["Container", "Réception", "N°", "Fournisseur", "Total HT", "FFR"]).sort_values(["Container", "Réception"]))


def generate_preparation_excel(calc: PreparationCalculator, result: dict, output_path: str) -> str:
    year, month = result["year"], result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    write_summary(ws, f"Préparation - {MONTHS_FR[month - 1]} {year}", summary_rows(result))

    write_preparation_sheets(wb, calc, year, month)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
