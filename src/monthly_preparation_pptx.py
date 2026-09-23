"""Remplit le template PowerPoint Préparation (1 diapo) avec les KPI calculés."""

from pathlib import Path

from pptx import Presentation

from .pptx_helpers import COMMENT_PLACEHOLDER, find_shape, fr_k, set_delta, set_subtitle, set_text


def _fill_slide1(slide, result: dict):
    d, e = result["data"], result["evolutions"]
    cur, n1 = d["cur"], d["n1"]
    cmd, cmd1 = cur["commandes"], n1["commandes"]
    cont, cont1 = cur["conteneurs"], n1["conteneurs"]

    set_text(find_shape(slide, 9), f"{cmd['total']} cdes")
    set_delta(find_shape(slide, 11), f"N-1 : {cmd1['total']} cdes", e["total_n1"])
    set_text(find_shape(slide, 12), fr_k(cmd["ca"]))
    set_delta(find_shape(slide, 14), f"N-1 : {fr_k(cmd1['ca'])}", e["ca_n1"])

    set_text(find_shape(slide, 54), f"{cmd['clients']} cdes")
    set_delta(find_shape(slide, 32), f"N-1 : {cmd1['clients']} cdes", e["clients_n1"])
    set_text(find_shape(slide, 24), f"{cmd['reassort']} cdes")
    set_delta(find_shape(slide, 26), f"N-1 : {cmd1['reassort']} cdes", e["reassort_n1"])
    set_text(find_shape(slide, 74), f"{cmd['sav']} cdes")
    set_delta(find_shape(slide, 20), f"N-1 : {cmd1['sav']} cdes", e["sav_n1"])

    set_text(find_shape(slide, 37), f"{cont['nb']} ctainer")
    set_delta(find_shape(slide, 28), f"N-1 : {cont1['nb']}", e["conteneurs_nb_n1"])
    set_text(find_shape(slide, 47), fr_k(cont["ca"]))
    set_delta(find_shape(slide, 29), f"N-1 : {fr_k(cont1['ca'])}", e["conteneurs_ca_n1"])

    set_text(find_shape(slide, 33), COMMENT_PLACEHOLDER)


def generate_preparation_pptx(result: dict, template_path: str, output_path: str) -> str:
    prs = Presentation(template_path)
    slide = prs.slides[0]
    set_subtitle(find_shape(slide, 5), result["year"], result["month"])
    _fill_slide1(slide, result)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path
