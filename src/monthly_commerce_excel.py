"""Excel de sortie de la section Commerce : synthèse chiffrée + feuilles de détail."""

from pathlib import Path
from typing import List, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .monthly_commerce import CommerceCalculator, SEUILS_LIVRAISON
from .monthly_loader import MONTHS_FR

HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FONT = Font(bold=True, color="366092")

SUMMARY_HEADERS = ["Indicateur", "Valeur", "N-1", "Évol. N-1 (%)", "Mois-1", "Évol. Mois-1 (%)"]


def summary_rows(result: dict) -> List[Optional[list]]:
    """Lignes de la feuille Synthèse : [libellé, valeur, N-1, évol N-1, Mois-1, évol Mois-1]. None = ligne titre/vide."""
    d, e = result["data"], result["evolutions"]
    cur, n1, m1 = d["cur"], d["n1"], d["m1"]
    ea = cur["en_attente"]

    rows: List[Optional[list]] = [
        ["Commandes clients passées"],
        ["Nb commandes", cur["commandes"]["nb"], n1["commandes"]["nb"], e["commandes_nb_n1"], m1["commandes"]["nb"], None],
        ["  dont archivées", cur["commandes"]["nb_arch"], n1["commandes"]["nb_arch"], None, m1["commandes"]["nb_arch"], None],
        ["  dont à livrer", cur["commandes"]["nb_alivr"], n1["commandes"]["nb_alivr"], None, m1["commandes"]["nb_alivr"], None],
        ["CA HT commandes (€)", cur["commandes"]["ca"], n1["commandes"]["ca"], e["commandes_ca_n1"], m1["commandes"]["ca"], None],
        ["  dont archivées (Total HT)", cur["commandes"]["ca_arch"], n1["commandes"]["ca_arch"], None, m1["commandes"]["ca_arch"], None],
        ["  dont à livrer (A livrer Net)", cur["commandes"]["ca_alivr"], n1["commandes"]["ca_alivr"], None, m1["commandes"]["ca_alivr"], None],
        None,
        ["Facturation"],
        ["Nb factures (équipe + vide)", cur["factures"]["nb"], n1["factures"]["nb"], e["factures_nb_n1"], m1["factures"]["nb"], None],
        ["CA HT facturé (€)", cur["factures"]["ca"], n1["factures"]["ca"], e["factures_ca_n1"], m1["factures"]["ca"], None],
        ["Panier moyen avec Franck (€)", cur["factures"]["pm_avec"], n1["factures"]["pm_avec"], e["pm_avec_n1"], m1["factures"]["pm_avec"], None],
        ["Panier moyen hors Franck (€)", cur["factures"]["pm_sans"], n1["factures"]["pm_sans"], e["pm_sans_n1"], m1["factures"]["pm_sans"], None],
        ["CA HT facturé hors Franck (€)", cur["factures"]["ca_sans"], n1["factures"]["ca_sans"], None, m1["factures"]["ca_sans"], None],
        None,
        ["Commandes sous le seuil de livraison"],
        ["Nb livraisons du mois", cur["seuils"]["total"], n1["seuils"]["total"], None, m1["seuils"]["total"], None],
    ]
    for seuil in SEUILS_LIVRAISON:
        rows += [
            [f"Livraisons < {seuil} € (nb)", cur["seuils"][f"nb_{seuil}"], n1["seuils"][f"nb_{seuil}"], None, m1["seuils"][f"nb_{seuil}"], None],
            [f"Livraisons < {seuil} € (% des livraisons)", cur["seuils"][f"part_{seuil}"], n1["seuils"][f"part_{seuil}"],
             e[f"seuil_{seuil}_n1"], m1["seuils"][f"part_{seuil}"], e[f"seuil_{seuil}_m1"]],
        ]
    rows += [
        None,
        ["Nouveaux clients (1ère commande)"],
        ["Nb nouveaux clients", cur["nouveaux_clients"]["nb"], None, None, m1["nouveaux_clients"]["nb"], e["nouveaux_clients_nb_m1"]],
        ["CA HT nouveaux clients (€)", cur["nouveaux_clients"]["ca"], None, None, m1["nouveaux_clients"]["ca"], e["nouveaux_clients_ca_m1"]],
        None,
        ["Commandes en attente de livraison (A livrer Net, €)"],
    ]
    for label, key in (("Commande initiale", "initiale"), ("Cde avec reliquats", "reliquats"), ("TOTAL", "total")):
        rows += [
            [f"{label} - nb (Franck + vide)", ea[key]["nb"], None, None, None, None],
            [f"{label} - montant avec Franck (+ vide)", ea[key]["net_avec"], None, None, None, None],
            [f"{label} - montant hors Franck", ea[key]["net_sans"], None, None, None, None],
        ]
    return rows


def _write_summary(ws, result: dict):
    year, month = result["year"], result["month"]
    ws.append([f"Commerce - {MONTHS_FR[month - 1]} {year}"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(SUMMARY_HEADERS)
    for col in range(1, len(SUMMARY_HEADERS) + 1):
        cell = ws.cell(row=3, column=col)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL

    for row in summary_rows(result):
        if row is None:
            ws.append([])
        elif len(row) == 1:
            ws.append(row)
            ws.cell(row=ws.max_row, column=1).font = SECTION_FONT
        else:
            ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
            for col in (2, 3, 5):
                ws.cell(row=ws.max_row, column=col).number_format = "#,##0.0" if isinstance(row[col - 1], float) else "#,##0"
            for col in (4, 6):
                ws.cell(row=ws.max_row, column=col).number_format = "+0.0;-0.0;0.0"

    ws.column_dimensions["A"].width = 52
    for col in "BCDEF":
        ws.column_dimensions[col].width = 16


def _write_frame(wb: Workbook, title: str, df: pd.DataFrame):
    ws = wb.create_sheet(title)
    for col_idx, name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        ws.column_dimensions[get_column_letter(col_idx)].width = max(14, min(50, len(str(name)) + 4))
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            if isinstance(value, float) and pd.isna(value):
                value = None
            elif isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if hasattr(value, "year"):
                cell.number_format = "dd/mm/yyyy"


def generate_commerce_excel(calc: CommerceCalculator, result: dict, output_path: str) -> str:
    year, month = result["year"], result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    _write_summary(ws, result)

    def cols(df: pd.DataFrame, wanted: List[str]) -> pd.DataFrame:
        return df[[c for c in wanted if c in df.columns]]

    _write_frame(wb, "Commandes", cols(calc.commandes_rows(year, month),
                 ["Source", "Date", "N°", "Client", "Représentant", "Total HT", "A livrer Net", "Montant retenu"]))
    _write_frame(wb, "Factures", cols(calc.factures_rows(year, month, True), ["Date", "N°", "Client", "Représentant", "Total HT"]))
    _write_frame(wb, "Livraisons_seuils", cols(calc.seuils_rows(year, month), ["Date", "N°", "Tournée", "Client", "Représentant", "Total HT"]))
    _write_frame(wb, "En_attente", cols(calc.en_attente_rows(year, month),
                 ["Type", "Date", "Livraison", "N°", "Client", "Représentant", "Rlq", "Total HT", "A livrer Net"]))
    _write_frame(wb, "Nouveaux_clients", calc.nouveaux_clients_rows(year, month))

    top = result["data"]["cur"]["top_ventes"]
    _write_frame(wb, "Top_valeur", top["valeur"])
    _write_frame(wb, "Top_volume", top["volume"])

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
