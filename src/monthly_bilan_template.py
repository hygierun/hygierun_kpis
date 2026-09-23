"""Génère le fichier Excel 'template' d'Input Mensuel consolidé : une feuille par feuille attendue,
avec les colonnes obligatoires (+ les colonnes de détail utiles) déjà en en-tête, et - quand Antoine
l'a précisé - une note sous les en-têtes indiquant où aller chercher les données dans l'ERP. Reste à
l'utilisateur de coller les exports dans chaque feuille (et de renommer les colonnes datées comme
'T [date export]' ou '08/26 HT' selon le mois traité).
"""

from io import BytesIO
from typing import Dict, List

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from .monthly_loader import DETAIL_COLUMNS, bilan_total_required_columns

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
NOTE_FONT = Font(italic=True, color="808080")
LINK_FONT = Font(italic=True, color="0563C1", underline="single")

# Stock_Invent : colonnes réelles nécessaires au calcul au-delà des colonnes obligatoires (Dépot,
# Famille) - la colonne "T <date>" est dynamique (retrouvée par son préfixe "T "), à renommer chaque
# mois avec la date du jour de l'export (ex. "T 04/09/26").
EXTRA_COLUMNS: Dict[str, List[str]] = {
    "Stock_Invent": ["Désignation", "Référence", "Ref. fourn.", "Unité", "T [date export]", "Nb jours", "Anc Ref"],
}

# Où aller chercher chaque export dans l'ERP (précisé par Antoine) - une ligne de note par élément,
# placée juste sous les en-têtes.
SHEET_NOTES: Dict[str, List[str]] = {
    "Stats_NewClients": ["Chemin : Ventes/Stat_Générales/KPI/NewClient Input / Resultats"],
    "top10": ["Chemin : Ventes/Stat_Générales/KPI/Top10Input /Choisir le mois d'intérêt) Resultats"],
    "Delais Bis": ["Livraison / Archivées / filtre sur date / 2025 jusque today / CTRL+H / Delais Bis"],
    "Fact_Dues": ["Clients / CLQ+D / Comptabilité / Fact Dues"],
    "GPS_Livr": ["Concatener tous les fichiers reçu par mail fleet-alert sur le mois en question"],
    "MO+Depl": ["Chemin : Ventes/Stat_Générales/KPI/MO+Depl Input / 24derniers mois /Resultats"],
    "Bilan_Fiches": [
        "Faire tourner Claude avec le prompt dédié et les signatures des tech dédiées : voir le kit fiches SAV à télécharger",
        "Importer dans Claude le zip avec toutes les fiches SAV du mois (demander à Laetitia de partager les scans, ex. Z:COMMUNS/STAGIAIRES)",
    ],
    "Stock_Invent": ["Chemin : Stock / Stock / CLQ+D / Inventaires / CLQ sur \"Inventaire complet du dépôt DEPOT\" / CLQ Stock au date du jour / CLQ Achat Moyen pondéré"],
}

# Absences_Tech : les 4 techniciens dont l'absence compte dans la productivité (Joël est exclu de ce
# calcul) - prérempli pour éviter les fautes de frappe sur les noms.
ABSENCES_TECH_NOMS = ["Yann", "Christophe", "Nicolas", "Daniel"]
ABSENCES_TECH_NOTE = "À remplir sur le mois d'intérêt avec les infos du lien des congés Hygierun :"
ABSENCES_TECH_LINK = ("Congés Hygierun", "https://microstor.notion.site/Cong-s-Hygierun-39feff3e50128054bb11eae748d02ddf?pvs=74")


def _merge_columns(*groups: List[str]) -> List[str]:
    merged: List[str] = []
    for group in groups:
        for col in group:
            if col not in merged:
                merged.append(col)
    return merged


def generate_bilan_template_excel(year: int, month: int) -> bytes:
    """Fichier Excel vierge (en-têtes seulement) prêt à être rempli pour le mois donné."""
    required = bilan_total_required_columns(year, month)
    wb = Workbook()
    wb.remove(wb.active)

    for sheet, cols in required.items():
        columns = _merge_columns(cols, DETAIL_COLUMNS.get(sheet, []), EXTRA_COLUMNS.get(sheet, []))
        ws = wb.create_sheet(sheet)
        for col_idx, name in enumerate(columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=name)
            cell.font, cell.fill = HEADER_FONT, HEADER_FILL
            ws.column_dimensions[cell.column_letter].width = max(14, min(28, len(str(name)) + 4))
        ws.freeze_panes = "A2"

        if sheet == "Absences_Tech":
            for row_idx, nom in enumerate(ABSENCES_TECH_NOMS, 2):
                ws.cell(row=row_idx, column=1, value=nom)
            note_row = len(ABSENCES_TECH_NOMS) + 3
            ws.cell(row=note_row, column=1, value=ABSENCES_TECH_NOTE).font = NOTE_FONT
            link_cell = ws.cell(row=note_row + 1, column=1, value=ABSENCES_TECH_LINK[0])
            link_cell.hyperlink, link_cell.font = ABSENCES_TECH_LINK[1], LINK_FONT
            continue

        for note_idx, note in enumerate(SHEET_NOTES.get(sheet, [])):
            ws.cell(row=3 + note_idx, column=1, value=note).font = NOTE_FONT

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
