"""Excel de sortie des sections Livraison et Comptabilité : synthèse chiffrée + feuilles de détail."""

from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook

from .excel_helpers import select_columns, write_frame, write_summary
from .monthly_livraison import LivraisonComptaCalculator
from .monthly_loader import MONTHS_FR


def livraison_ops_rows(result: dict) -> List[Optional[list]]:
    """Lignes Synthèse propres aux opérations de livraison (camions, enlèvements, délais, multiples)."""
    d, e = result["data"], result["evolutions"]
    cur, n1, m1 = d["cur"], d["n1"], d["m1"]

    def line(label: str, block: str, field: str, evo_key: Optional[str] = None) -> list:
        return [label, cur[block][field], n1[block][field], e.get(f"{evo_key}_n1") if evo_key else None,
                m1[block][field], e.get(f"{evo_key}_m1") if evo_key else None]

    rows: List[Optional[list]] = [
        ["Livraison - CA livré par camion"],
        line("CA livré par camion (€)", "camions", "ca_livre", "ca_livre"),
        line("CA livré total hors tournées 0 et 524 (€)", "camions", "ca_livre_total"),
        line("Part du CA livré par camion (%)", "camions", "part_ca_livre"),
        line("Nb tournées", "camions", "nb_tournees"),
        line("Nb BL A moyen par tournée", "camions", "bla_moy", "bla"),
        line("Nb factures moyen par tournée", "camions", "fact_moy", "fact"),
        line("Camion microstor (nb tournées)", "camions", "microstor"),
        line("Sous-traitances externes (nb tournées)", "camions", "sous_traitance"),
        line("Enlèvements clients", "enlevements", "enlevements", "enlevements"),
        line("Livraisons commerciaux", "enlevements", "commerciaux", "commerciaux"),
        None,
        ["Livraison - délais (tournées camion, délai cde → livraison entre 0 et 14 j)"],
        line("Délai commande → livraison (jours)", "delais", "cde_livraison", "delai_cde_livraison"),
        line("Délai livraison → facturation (jours)", "delais", "livraison_facture", "delai_livraison_facture"),
        line("Nb lignes retenues", "delais", "nb"),
        None,
        ["Livraison - multiples livraisons (toutes les commandes archivées)"],
        line("Nb commandes", "multiples", "total"),
        line("Commandes livrées en 1 BL", "multiples", "nb_1bl"),
        line("Part des commandes livrées en 1 BL (%)", "multiples", "part_1bl", "part_1bl"),
    ]
    for bucket in cur["multiples"]["buckets"]:
        rows.append([f"  commandes avec {bucket} BL", cur["multiples"]["buckets"][bucket],
                     n1["multiples"]["buckets"][bucket], None, m1["multiples"]["buckets"][bucket], None])
    return rows


def compta_rows(result: dict) -> List[Optional[list]]:
    """Lignes Synthèse propres à la compta (clients bloqués, factures dues) - photo du jour de l'export."""
    cb, fd = result["data"]["cur"]["clients_bloques"], result["data"]["cur"]["factures_dues"]
    return [
        ["Compta - clients bloqués (photo du jour de l'export)"],
        ["Clients bloqués", cb["total"], None, None, None, None],
        ["  dont actifs sur l'année (ventes > 0)", cb["actifs"], None, None, None, None],
        ["  CA de l'année des clients actifs (€)", cb["ca"], None, None, None, None],
        ["  Solde compta des clients actifs (€)", cb["solde"], None, None, None, None],
        None,
        ["Compta - factures impayées (photo du jour de l'export)"],
        ["Factures impayées : nb (Nb JEch > 0, montant >= 0)", fd["nb"], None, None, None, None],
        ["Factures impayées : montant (€)", fd["montant"], None, None, None, None],
        ["  part des factures de l'année : nb (%)", fd["part_annee_nb"], None, None, None, None],
        ["  part des factures de l'année : montant (%)", fd["part_annee_montant"], None, None, None, None],
        ["Factures impayées >= 60 j : nb (montant > 0)", fd["nb_60"], None, None, None, None],
        ["Factures impayées >= 60 j : montant (€)", fd["montant_60"], None, None, None, None],
        ["Factures impayées >= 60 j : clients distincts", fd["clients_60"], None, None, None, None],
        ["  part du total impayé : nb (%)", fd["part_60_nb"], None, None, None, None],
        ["  part du total impayé : montant (%)", fd["part_60_montant"], None, None, None, None],
    ]


def summary_rows(result: dict) -> List[Optional[list]]:
    """Lignes de la feuille Synthèse : [libellé, valeur, N-1, évol N-1, Mois-1, évol Mois-1]."""
    return livraison_ops_rows(result) + [None] + compta_rows(result)


def write_livraison_ops_sheets(wb: Workbook, calc: LivraisonComptaCalculator, year: int, month: int):
    write_frame(wb, "Tournees", select_columns(calc.tournees_rows(year, month),
                ["Date", "N°", "Désignation", "Chauffeur", "Camion", "Catégorie", "Total HT", "Nb bl A", "Nb fact"]))
    write_frame(wb, "Enlevements_commerciaux", select_columns(calc.enlevements_commerciaux_rows(year, month),
                ["Catégorie", "Date", "N°", "Tournée", "Transporteur", "Client", "Total HT"]))
    write_frame(wb, "Delais", select_columns(calc.delais_rows(year, month),
                ["Date", "N°", "Tournée", "Liv. souhaitée", "Date Creation Cde", "Fact date",
                 "Délai souhaité (j)", "Délai cde → livraison (j)", "Délai livraison → facture (j)"]))
    write_frame(wb, "Multiples_livraisons", select_columns(calc.multiples_rows(year, month),
                ["Livraison", "N°", "Client", "Représentant", "Tournée", "Nb bls"]))
    gps = calc.gps_tournees_rows(year, month)
    if not gps.empty:
        write_frame(wb, "GPS_tournees", gps)
        write_frame(wb, "GPS_par_chauffeur", calc.gps_par_chauffeur(year, month).reset_index(names="Chauffeur"))


def write_compta_sheets(wb: Workbook, calc: LivraisonComptaCalculator, year: int, month: int):
    write_frame(wb, "Clients_bloques", select_columns(calc.clients_bloques_rows(year),
                ["Référence", "Désignation", "Qualification", f"Vtes {year}", "Solde cpta"]))
    write_frame(wb, "Factures_dues", select_columns(calc.factures_dues_rows(year, month),
                ["N°", "Date", "Client", "Client (réf.)", "Echéance", "Nb JEch", "Nb JEch retenu", "Restant dû", "> 60 j"]))


def generate_livraison_excel(calc: LivraisonComptaCalculator, result: dict, output_path: str) -> str:
    year, month = result["year"], result["month"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    write_summary(ws, f"Livraison et Compta - {MONTHS_FR[month - 1]} {year}", summary_rows(result))

    write_livraison_ops_sheets(wb, calc, year, month)
    write_compta_sheets(wb, calc, year, month)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
