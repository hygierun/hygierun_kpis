"""Lit le Mois-1 depuis le PowerPoint Bilan Mensuel *consolidé* (11 diapos), en réutilisant le même
cœur de lecture que les decks autonomes - seuls les index de diapo changent (diapo 1 = Livraison GPS et
diapo 3 = Compta deviennent les diapos 6 et 8 du deck fusionné ; SAV/Achat = diapos 9/10)."""

from typing import Dict, Optional

from pptx import Presentation

from .monthly_livraison_mois1 import _read_from_slides as _read_livraison_slides
from .monthly_sav_mois1 import _read_from_slides as _read_sav_slides

LIVRAISON_GPS_SLIDE, LIVRAISON_COMPTA_SLIDE = 6, 8
SAV_SLIDE, ACHAT_SLIDE = 9, 10


def read_bilan_mois1_reference(pptx_path: str) -> Dict[str, Optional[dict]]:
    """Retourne {'livraison': {...} | None, 'sav': {...} | None} - chacun au même format que les
    lecteurs autonomes, prêt à passer tel quel en mois1_reference des calculators concernés."""
    try:
        prs = Presentation(pptx_path)
        slides = prs.slides
        if len(slides) < 11:
            return {"livraison": None, "sav": None}
    except Exception:
        return {"livraison": None, "sav": None}

    try:
        livraison = _read_livraison_slides(slides[LIVRAISON_GPS_SLIDE], slides[LIVRAISON_COMPTA_SLIDE])
    except Exception:
        livraison = None

    try:
        sav = _read_sav_slides(slides[SAV_SLIDE], slides[ACHAT_SLIDE])
    except Exception:
        sav = None

    return {"livraison": livraison, "sav": sav}
