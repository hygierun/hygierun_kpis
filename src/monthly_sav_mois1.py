"""Lecture du PowerPoint SAV du mois précédent pour le Mois-1 du bloc Nombre d'interventions.

Comme pour Livraison & Compta, ce KPI n'est pas recalculable depuis l'historique (les fiches papier ne
sont lues qu'une fois, pour le mois en cours) : son évolution Mois-1 se récupère en relisant le deck
généré le mois précédent par cette appli.
"""

import re
from typing import Dict, Optional

from pptx import Presentation

from .pptx_helpers import find_shape

NB_INTERVENTIONS_SHAPES = {"total": 34, "ebc": 51, "sav": 53}


def _parse_int(text: str) -> Optional[int]:
    match = re.search(r"(\d+)", text)
    return int(match.group(1)) if match else None


def read_sav_mois1_reference(pptx_path: str) -> Dict[str, Dict[str, Optional[int]]]:
    prs = Presentation(pptx_path)
    slide = prs.slides[0]
    nb_interventions = {
        champ: _parse_int(find_shape(slide, shape_id).text_frame.text)
        for champ, shape_id in NB_INTERVENTIONS_SHAPES.items()
    }
    return {"nb_interventions": nb_interventions}
