"""Excel de sortie de la section Commerce : synthèse chiffrée + feuilles de détail."""

from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook

from .excel_helpers import SUMMARY_HEADERS, select_columns, write_frame, write_summary  # noqa: F401
from .monthly_commerce import CommerceCalculator, SEUILS_LIVRAISON
from .monthly_loader import MONTHS_FR


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
            [f"{label} - nb (hors Franck)", ea[key]["nb"], None, None, None, None],
            [f"{label} - montant avec Franck (+ vide)", ea[key]["net_avec"], None, None, None, None],
            [f"{label} - montant hors Franck", ea[key]["net_sans"], None, None, None, None],
        ]
    return rows


def generate_commerce_excel(calc: CommerceCalculator, result: dict, output_path: str) -> str:
    year, month = result["year"], result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    write_summary(ws, f"Commerce - {MONTHS_FR[month - 1]} {year}", summary_rows(result))

    write_frame(wb, "Commandes", select_columns(calc.commandes_rows(year, month),
                 ["Source", "Date", "N°", "Client", "Représentant", "Total HT", "A livrer Net", "Montant retenu"]))
    write_frame(wb, "Factures", select_columns(calc.factures_rows(year, month, True), ["Date", "N°", "Client", "Représentant", "Total HT"]))
    write_frame(wb, "Livraisons_seuils", select_columns(calc.seuils_rows(year, month), ["Date", "N°", "Tournée", "Client", "Représentant", "Total HT"]))
    write_frame(wb, "En_attente", select_columns(calc.en_attente_rows(year, month),
                 ["Type", "Date", "Livraison", "N°", "Client", "Représentant", "Rlq", "Total HT", "A livrer Net"]))
    write_frame(wb, "Nouveaux_clients", calc.nouveaux_clients_rows(year, month))

    top = result["data"]["cur"]["top_ventes"]
    write_frame(wb, "Top_valeur", top["valeur"])
    write_frame(wb, "Top_volume", top["volume"])

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
