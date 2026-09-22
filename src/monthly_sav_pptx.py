"""Remplit le template PowerPoint SAV (diapo 1, 4 blocs) avec les KPI calculés.

Les blocs pas encore implémentés (Productivité, Nombre d'interventions) sont affichés en "n/a".
La diapo 2 (Achat/Appro) n'est pas touchée : elle garde ses placeholders "à collecter" du template.
"""

from copy import deepcopy
from pathlib import Path

from pptx import Presentation

from .pptx_helpers import arrow, delta_color, find_shape, fr, fr_k, fr_pct, set_cell, set_delta, set_delta_na, set_subtitle, set_text


def _heures(value: float) -> str:
    """Pas de décimale pour un nombre d'heures entier (144), une décimale sinon (107,5)."""
    return fr(value, 0) if float(value).is_integer() else fr(value, 1)


def _set_devis_mois1(shape, nb: int, pct):
    """Cas particulier : 3 runs ('Mois', ' -1 : N devis', 'flèche %'), pas le schéma standard à 2 runs."""
    runs = shape.text_frame.paragraphs[0].runs
    runs[1].text = f" -1 : {nb} devis"
    runs[2].text = f"{arrow(pct)} {fr_pct(pct)}"
    runs[2].font.color.rgb = delta_color(pct)


def _ensure_second_paragraph(shape, template_shape):
    """La case Mois-1 du CA Docs Nuls n'a qu'un paragraphe (placeholder vide) : on clone la 2e ligne
    (grande, colorée) d'une case N-1/Mois-1 déjà stylée avant d'y écrire l'évolution normalement."""
    if len(shape.text_frame.paragraphs) < 2:
        modele = deepcopy(template_shape.text_frame.paragraphs[1]._p)
        shape.text_frame.paragraphs[0]._p.addnext(modele)


def _fill_devis(slide, result: dict):
    d, e = result["data"], result["evolutions"]
    cur, n1, m1 = d["cur"]["devis"], d["n1"]["devis"], d["m1"]["devis"]

    set_text(find_shape(slide, 100), f"{cur['nb']} devis")
    set_text(find_shape(slide, 102), fr_k(cur["ca"], 1))
    set_delta(find_shape(slide, 93), f"N-1 : {n1['nb']} devis", e["devis_nb_n1"])
    _set_devis_mois1(find_shape(slide, 83), m1["nb"], e["devis_nb_m1"])


def _fill_docs_nuls(slide, result: dict):
    d, e = result["data"], result["evolutions"]
    cur, n1, m1 = d["cur"]["docs_nuls"], d["n1"]["docs_nuls"], d["m1"]["docs_nuls"]

    set_text(find_shape(slide, 112), f"{cur['total']} Fact")
    set_text(find_shape(slide, 118), f"{cur['dn']} DN")
    set_text(find_shape(slide, 120), f"{cur['non_dn']} non DN")
    set_text(find_shape(slide, 104), fr_k(cur["ca"], 1))

    set_delta(find_shape(slide, 129), f"Mois-1 : {m1['total']}", e["docs_nuls_total_m1"])
    set_delta(find_shape(slide, 122), f"Mois-1 : {m1['dn']} DN", e["docs_nuls_dn_m1"])
    set_delta(find_shape(slide, 116), f"Mois-1 : {m1['non_dn']}", e["docs_nuls_non_dn_m1"])

    ca_mois1 = find_shape(slide, 148)
    ca_n1 = find_shape(slide, 106)
    _ensure_second_paragraph(ca_mois1, ca_n1)
    set_delta(ca_n1, f"N-1 : {fr_k(n1['ca'], 1)}", e["docs_nuls_ca_n1"])
    set_delta(ca_mois1, f"Mois-1 : {fr_k(m1['ca'], 1)}", e["docs_nuls_ca_m1"])


def _fill_main_oeuvre(slide, result: dict):
    d, e = result["data"], result["evolutions"]
    cur, n1 = d["cur"]["main_oeuvre"], d["n1"]["main_oeuvre"]

    set_text(find_shape(slide, 72), f"{_heures(cur['total'])}h")
    set_text(find_shape(slide, 78), f"{_heures(cur['ebc'])} EBC")
    set_text(find_shape(slide, 79), f"{_heures(cur['sav'])} SAV")
    set_delta(find_shape(slide, 82), f"N-1 : {_heures(n1['total'])}", e["main_oeuvre_total_n1"])
    set_delta(find_shape(slide, 77), f"N-1 : {_heures(n1['ebc'])}", e["main_oeuvre_ebc_n1"])
    set_delta(find_shape(slide, 80), f"N-1 : {_heures(n1['sav'])}", e["main_oeuvre_sav_n1"])


def _fill_deplacement(slide, result: dict):
    d, e = result["data"], result["evolutions"]
    cur, n1 = d["cur"]["deplacement"], d["n1"]["deplacement"]

    set_text(find_shape(slide, 130), f"{cur['total']:.0f} u")
    set_text(find_shape(slide, 138), f"{cur['ebc']:.0f} EBC")
    set_text(find_shape(slide, 140), f"{cur['sav']:.0f} SAV")
    set_delta(find_shape(slide, 146), f"N-1 : {n1['total']:.0f}", e["deplacement_total_n1"])
    set_delta(find_shape(slide, 134), f"N-1 : {n1['ebc']:.0f}", e["deplacement_ebc_n1"])
    set_delta(find_shape(slide, 142), f"N-1 : {n1['sav']:.0f}", e["deplacement_sav_n1"])


def _fill_non_implemente(slide):
    """Productivité et Nombre d'interventions : pas encore calculés."""
    table = find_shape(slide, 9).table
    for row in range(1, len(table.rows)):
        for col in range(1, len(table.columns)):
            set_cell(table.cell(row, col), "n/a")

    set_text(find_shape(slide, 34), "n/a")
    for shape_id in (51, 53):
        set_text(find_shape(slide, shape_id), "n/a")
    for shape_id in (28, 50, 55):
        set_delta_na(find_shape(slide, shape_id), "Mois-1")


def generate_sav_pptx(result: dict, template_path: str, output_path: str) -> str:
    prs = Presentation(template_path)
    slide1 = prs.slides[0]
    set_subtitle(find_shape(slide1, 5), result["year"], result["month"])
    _fill_devis(slide1, result)
    _fill_docs_nuls(slide1, result)
    _fill_main_oeuvre(slide1, result)
    _fill_deplacement(slide1, result)
    _fill_non_implemente(slide1)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path
